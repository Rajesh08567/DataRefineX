using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using DataRefineX.Models;
using MiniExcelLibs;

namespace DataRefineX.Services;

public sealed class ProcessingResult
{
    public bool Success { get; init; }
    public string? OutputPath { get; init; }
    public long TotalRowsRead { get; init; }
    public long DuplicatesRemoved { get; init; }
    public long InvalidRemoved { get; init; }
    public long ValidRecords { get; init; }
    public TimeSpan Elapsed { get; init; }
    public string? ErrorMessage { get; init; }
}

public enum SheetSelectionMode
{
    All,
    ByName,
    FirstN
}

public enum DedupKeyMode
{
    /// <summary>Use a single column as the dedup key (case-insensitive trim+lowercase).</summary>
    SingleColumn,
    /// <summary>Concatenate multiple columns to form the dedup key.</summary>
    MultipleColumns,
    /// <summary>Treat the entire row (all column values joined) as the dedup key.</summary>
    EntireRow,
    /// <summary>No dedup — keep every row, just split out invalid/empty rows if validation enabled.</summary>
    None
}

[Flags]
public enum ValidationMode
{
    None     = 0,
    Email    = 1 << 0,
    NotEmpty = 1 << 1
}

public enum OutputFormat
{
    Xlsx,
    Csv
}

public enum OutputDestination
{
    /// <summary>Write a single new file (xlsx with multiple sheets, or one csv per bucket).</summary>
    NewFile,
    /// <summary>Overwrite each source file in place — unique rows replace the original sheets, duplicates/invalid go to extra sheets.</summary>
    InPlace
}

public sealed class ProcessingOptions
{
    public SheetSelectionMode SheetMode { get; init; } = SheetSelectionMode.All;
    public string[] SheetNames { get; init; } = Array.Empty<string>();
    public int FirstNSheets { get; init; } = 1;

    public DedupKeyMode DedupMode { get; init; } = DedupKeyMode.SingleColumn;
    public string[] DedupColumns { get; init; } = Array.Empty<string>();

    /// <summary>If true, cells in the dedup column containing the delimiters are split into multiple logical rows (one per value).</summary>
    public bool SplitMultiValueCells { get; init; } = false;
    /// <summary>Characters that separate multiple values in a single cell. Default: ; , | newline.</summary>
    public string MultiValueDelimiters { get; init; } = ";,|\n\r";

    public ValidationMode Validation { get; init; } = ValidationMode.None;
    /// <summary>Column to validate when <see cref="Validation"/> != None. Empty = uses first dedup column.</summary>
    public string ValidationColumn { get; init; } = "";

    public bool WriteUniqueSheet { get; init; } = true;
    public bool WriteDuplicatesSheet { get; init; } = false;
    public bool WriteInvalidSheet { get; init; } = false;

    public string UniqueSheetName { get; init; } = "Unique";
    public string DuplicatesSheetName { get; init; } = "Duplicates";
    public string InvalidSheetName { get; init; } = "Invalid";

    public string? OutputDirectory { get; init; }
    public OutputFormat OutputFormat { get; init; } = OutputFormat.Xlsx;
    public OutputDestination Destination { get; init; } = OutputDestination.NewFile;
    /// <summary>When true, output keeps the original sheet names (one output sheet per input sheet, dedup per-sheet). When false, all unique rows merge into a single 'Unique' sheet.</summary>
    public bool PreserveSourceSheets { get; init; } = true;
    /// <summary>If true, dedup matching is case-sensitive (e.g. 'Foo' and 'foo' are different). Default false (case-insensitive — typical for emails, usernames, etc.).</summary>
    public bool CaseSensitive { get; init; } = false;
    public int MaxDegreeOfParallelism { get; init; } = Math.Max(2, Environment.ProcessorCount);
}

public sealed class ExcelProcessor
{
    private const string SourceColumnName = "_Source";
    private const string ReasonColumnName = "_Reason";

    private static readonly Regex EmailRegex = new(
        @"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public async Task<ProcessingResult> ProcessAsync(
        IReadOnlyList<FileItem> files,
        ProcessingOptions options,
        IProgress<ProgressUpdate> progress,
        Action<LogLevel, string> log,
        CancellationToken cancellationToken)
    {
        // In-place writes operate per-file (dedup within each file independently).
        if (options.Destination == OutputDestination.InPlace && files.Count > 1)
        {
            return await ProcessPerFileAsync(files, options, progress, log, cancellationToken);
        }

        return await ProcessCombinedAsync(files, options, progress, log, cancellationToken);
    }

    private async Task<ProcessingResult> ProcessPerFileAsync(
        IReadOnlyList<FileItem> files,
        ProcessingOptions options,
        IProgress<ProgressUpdate> progress,
        Action<LogLevel, string> log,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        var totalRows = 0L;
        var totalDups = 0L;
        var totalInvalid = 0L;
        var totalUnique = 0L;
        string? lastOutput = null;

        for (var i = 0; i < files.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var single = new[] { files[i] };
            var r = await ProcessCombinedAsync(single, options, progress, log, cancellationToken);
            if (!r.Success) continue;
            totalRows += r.TotalRowsRead;
            totalDups += r.DuplicatesRemoved;
            totalInvalid += r.InvalidRemoved;
            totalUnique += r.ValidRecords;
            lastOutput = r.OutputPath ?? lastOutput;
        }

        sw.Stop();
        return new ProcessingResult
        {
            Success = true,
            OutputPath = lastOutput,
            TotalRowsRead = totalRows,
            DuplicatesRemoved = totalDups,
            InvalidRemoved = totalInvalid,
            ValidRecords = totalUnique,
            Elapsed = sw.Elapsed
        };
    }

    private async Task<ProcessingResult> ProcessCombinedAsync(
        IReadOnlyList<FileItem> files,
        ProcessingOptions options,
        IProgress<ProgressUpdate> progress,
        Action<LogLevel, string> log,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();

        try
        {
            if (files.Count == 0)
            {
                return new ProcessingResult { Success = false, ErrorMessage = "No files to process." };
            }

            log(LogLevel.Info, $"Starting processing of {files.Count} file(s) with {options.MaxDegreeOfParallelism} workers.");
            log(LogLevel.Info, $"Sheet mode: {options.SheetMode} • Dedup: {DescribeDedup(options)} • Validation: {options.Validation}");

            var totalRowsRead = 0L;
            var duplicatesRemoved = 0L;
            var invalidRemoved = 0L;
            var filesDoneCount = 0;

            var columnOrder = new List<string>();
            var columnOrderLock = new object();
            var columnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Per-sheet seen-keys when preserving structure; otherwise one global set.
            var seenKeysGlobal = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);
            var seenKeysPerSheet = new ConcurrentDictionary<string, ConcurrentDictionary<string, byte>>(StringComparer.OrdinalIgnoreCase);

            // Per-sheet row buckets — sheetName -> bag of unique rows.
            var uniquePerSheet = new ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            var duplicatesPerSheet = new ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            var invalidPerSheet = new ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);

            // Preserve the order in which sheets are first encountered, so output sheets keep load order.
            var sheetOrder = new List<string>();
            var sheetOrderSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sheetOrderLock = new object();

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = options.MaxDegreeOfParallelism,
                CancellationToken = cancellationToken
            };

            await Parallel.ForEachAsync(files, parallelOptions, async (file, ct) =>
            {
                if (ct.IsCancellationRequested) return;

                file.Status = FileStatus.Reading;
                progress.Report(BuildProgress(filesDoneCount, files.Count, file.FileName,
                    totalRowsRead, duplicatesRemoved, invalidRemoved, CountAllRows(uniquePerSheet)));

                try
                {
                    if (!File.Exists(file.FullPath))
                    {
                        file.Status = FileStatus.Failed;
                        file.Message = "File not found";
                        log(LogLevel.Error, $"{file.FileName}: file not found.");
                        return;
                    }

                    var isCsv = IsCsvFile(file.FullPath);

                    List<string> sheetsToRead;
                    if (isCsv)
                    {
                        // CSVs have no concept of multiple sheets — always treat as a single virtual sheet.
                        sheetsToRead = new List<string> { "" };
                    }
                    else
                    {
                        var sheetsInFile = await Task.Run(() => SafeGetSheetNames(file.FullPath), ct);
                        sheetsToRead = ResolveSheetsToRead(sheetsInFile, options);
                    }

                    if (sheetsToRead.Count == 0)
                    {
                        file.Status = FileStatus.Skipped;
                        file.Message = "No matching sheets";
                        log(LogLevel.Warning, $"{file.FileName}: no sheets matched mode '{options.SheetMode}' — skipping.");
                        Interlocked.Increment(ref filesDoneCount);
                        return;
                    }

                    var fileRowsRead = 0;

                    foreach (var sheet in sheetsToRead)
                    {
                        ct.ThrowIfCancellationRequested();

                        var bucketKey = options.PreserveSourceSheets
                            ? (isCsv
                                ? Path.GetFileNameWithoutExtension(file.FullPath)
                                : sheet)
                            : options.UniqueSheetName;

                        var seenForBucket = options.PreserveSourceSheets
                            ? seenKeysPerSheet.GetOrAdd(bucketKey, _ => new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase))
                            : seenKeysGlobal;

                        var uniqueBag = uniquePerSheet.GetOrAdd(bucketKey, _ => new ConcurrentBag<Dictionary<string, object?>>());

                        lock (sheetOrderLock)
                        {
                            if (sheetOrderSet.Add(bucketKey)) sheetOrder.Add(bucketKey);
                        }

                        await Task.Run(() =>
                        {
                            var rows = isCsv
                                ? MiniExcel.Query(file.FullPath, useHeaderRow: true, excelType: MiniExcelLibs.ExcelType.CSV)
                                : MiniExcel.Query(file.FullPath, useHeaderRow: true, sheetName: sheet);

                            foreach (var raw in rows)
                            {
                                if (ct.IsCancellationRequested) break;
                                if (raw is not IDictionary<string, object> dict) continue;

                                var snapshot = SnapshotRow(dict);

                                // Skip rows where every column is empty/whitespace — they're not data at all.
                                if (IsRowEmpty(snapshot))
                                {
                                    continue;
                                }

                                Interlocked.Increment(ref totalRowsRead);
                                fileRowsRead++;

                                if (dict.Count > 0)
                                {
                                    lock (columnOrderLock)
                                    {
                                        foreach (var rawCol in dict.Keys)
                                        {
                                            if (rawCol is null) continue;
                                            var col = rawCol.Trim();
                                            if (col.Length == 0) continue;
                                            if (col == SourceColumnName || col == ReasonColumnName) continue;
                                            if (columnSet.Add(col)) columnOrder.Add(col);
                                        }
                                    }
                                }

                                // Optional: split multi-value cells in the dedup column into separate rows.
                                var rowsToProcess = options.SplitMultiValueCells
                                    ? ExpandMultiValueRow(snapshot, options)
                                    : new List<Dictionary<string, object?>> { snapshot };

                                foreach (var rowToProcess in rowsToProcess)
                                {
                                    // Validation (combined Email + NotEmpty)
                                    if (options.Validation != ValidationMode.None)
                                    {
                                        var validationCol = string.IsNullOrWhiteSpace(options.ValidationColumn)
                                            ? (options.DedupColumns.Length > 0 ? options.DedupColumns[0] : "")
                                            : options.ValidationColumn;
                                        var validationValue = LookupColumn(rowToProcess, validationCol)?.Trim() ?? "";

                                        var (ok, reason) = Validate(validationValue, options.Validation, validationCol);
                                        if (!ok)
                                        {
                                            Interlocked.Increment(ref invalidRemoved);
                                            rowToProcess[ReasonColumnName] = reason;
                                            var invBag = invalidPerSheet.GetOrAdd(bucketKey, _ => new ConcurrentBag<Dictionary<string, object?>>());
                                            invBag.Add(rowToProcess);
                                            continue;
                                        }
                                    }

                                    // Dedup — empty dedup keys are kept in the unique bag (no dedup applies).
                                    if (options.DedupMode != DedupKeyMode.None)
                                    {
                                        var key = BuildDedupKey(rowToProcess, options);
                                        if (!string.IsNullOrEmpty(key))
                                        {
                                            if (!seenForBucket.TryAdd(key, 0))
                                            {
                                                Interlocked.Increment(ref duplicatesRemoved);
                                                var dupBag = duplicatesPerSheet.GetOrAdd(bucketKey, _ => new ConcurrentBag<Dictionary<string, object?>>());
                                                dupBag.Add(rowToProcess);
                                                continue;
                                            }
                                        }
                                    }

                                    uniqueBag.Add(rowToProcess);
                                }

                                if (fileRowsRead % 5000 == 0)
                                {
                                    file.RowsRead = fileRowsRead;
                                    progress.Report(BuildProgress(filesDoneCount, files.Count, file.FileName,
                                        totalRowsRead, duplicatesRemoved, invalidRemoved, CountAllRows(uniquePerSheet)));
                                }
                            }
                        }, ct);
                    }

                    file.RowsRead = fileRowsRead;
                    file.Status = FileStatus.Processed;
                    log(LogLevel.Success, $"{file.FileName}: read {fileRowsRead:N0} rows from {sheetsToRead.Count} sheet(s).");
                }
                catch (OperationCanceledException)
                {
                    file.Status = FileStatus.Failed;
                    file.Message = "Cancelled";
                    throw;
                }
                catch (Exception ex)
                {
                    file.Status = FileStatus.Failed;
                    file.Message = ex.Message;
                    log(LogLevel.Error, $"{file.FileName}: {ex.Message}");
                }
                finally
                {
                    Interlocked.Increment(ref filesDoneCount);
                    progress.Report(BuildProgress(filesDoneCount, files.Count, file.FileName,
                        totalRowsRead, duplicatesRemoved, invalidRemoved, CountAllRows(uniquePerSheet)));
                }
            });

            cancellationToken.ThrowIfCancellationRequested();

            List<string> finalColumns;
            lock (columnOrderLock)
            {
                finalColumns = new List<string>(columnOrder);
            }

            var totalUnique = CountAllRows(uniquePerSheet);
            log(LogLevel.Info, $"Accepted {totalUnique:N0} unique rows across {uniquePerSheet.Count} sheet(s) • {duplicatesRemoved:N0} duplicates • {invalidRemoved:N0} invalid. Writing output...");

            List<string> orderedKeysForLog;
            lock (sheetOrderLock)
            {
                orderedKeysForLog = new List<string>(sheetOrder);
            }

            if (invalidRemoved > 0)
            {
                var perSheet = string.Join(", ",
                    orderedKeysForLog
                        .Where(k => invalidPerSheet.ContainsKey(k))
                        .Select(k => $"{k}={invalidPerSheet[k].Count}"));
                log(LogLevel.Info, $"Invalid rows by sheet: {perSheet}");
            }
            if (duplicatesRemoved > 0)
            {
                var perSheet = string.Join(", ",
                    orderedKeysForLog
                        .Where(k => duplicatesPerSheet.ContainsKey(k))
                        .Select(k => $"{k}={duplicatesPerSheet[k].Count}"));
                log(LogLevel.Info, $"Duplicate rows by sheet: {perSheet}");
            }

            var outDir = string.IsNullOrWhiteSpace(options.OutputDirectory)
                ? Path.GetDirectoryName(files[0].FullPath) ?? Environment.CurrentDirectory
                : options.OutputDirectory!;
            Directory.CreateDirectory(outDir);

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string basePath;
            string primaryOutputFile;
            var isInPlace = options.Destination == OutputDestination.InPlace && files.Count == 1;

            if (isInPlace)
            {
                var sourcePath = files[0].FullPath;
                var sourceDir = Path.GetDirectoryName(sourcePath) ?? outDir;
                var sourceName = Path.GetFileNameWithoutExtension(sourcePath);
                basePath = Path.Combine(sourceDir, sourceName);
                primaryOutputFile = sourcePath; // overwrite the source file directly
            }
            else
            {
                basePath = Path.Combine(outDir, $"Processed_{stamp}");
                primaryOutputFile = options.OutputFormat == OutputFormat.Csv
                    ? $"{basePath}_{options.UniqueSheetName}.csv"
                    : $"{basePath}.xlsx";
            }

            await Task.Run(() =>
            {
                var srcExt = isInPlace ? Path.GetExtension(files[0].FullPath).ToLowerInvariant() : "";
                var writeAsCsv = isInPlace ? srcExt == ".csv" : options.OutputFormat == OutputFormat.Csv;

                List<string> orderedSheetKeys;
                lock (sheetOrderLock)
                {
                    orderedSheetKeys = sheetOrder
                        .Where(k => uniquePerSheet.ContainsKey(k))
                        .ToList();
                }

                if (writeAsCsv)
                {
                    // CSV is single-table per file. With preserved sheets, each input sheet becomes its own csv file.
                    if (options.WriteUniqueSheet || isInPlace)
                    {
                        if (isInPlace && uniquePerSheet.Count == 1)
                        {
                            // Single source CSV — overwrite it directly.
                            var only = uniquePerSheet.Values.First();
                            MiniExcel.SaveAs(primaryOutputFile,
                                NormalizeRows(only, finalColumns).ToList(),
                                excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                        }
                        else
                        {
                            foreach (var sheetName in orderedSheetKeys)
                            {
                                if (!uniquePerSheet.TryGetValue(sheetName, out var bag)) continue;
                                var path = uniquePerSheet.Count > 1
                                    ? $"{basePath}_{SanitizeFilename(sheetName)}.csv"
                                    : (options.WriteDuplicatesSheet || options.WriteInvalidSheet
                                        ? $"{basePath}_{options.UniqueSheetName}.csv"
                                        : primaryOutputFile);
                                MiniExcel.SaveAs(path,
                                    NormalizeRows(bag, finalColumns).ToList(),
                                    excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                            }
                        }
                    }
                    if (options.WriteDuplicatesSheet)
                    {
                        var mergedDup = FlattenWithSource(duplicatesPerSheet, orderedSheetKeys);
                        MiniExcel.SaveAs($"{basePath}_{options.DuplicatesSheetName}.csv",
                            NormalizeRowsWithSource(mergedDup, finalColumns),
                            excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                    }
                    if (options.WriteInvalidSheet)
                    {
                        var mergedInv = FlattenWithSource(invalidPerSheet, orderedSheetKeys);
                        MiniExcel.SaveAs($"{basePath}_{options.InvalidSheetName}.csv",
                            NormalizeRowsWithSource(mergedInv, finalColumns, includeReason: true),
                            excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                    }
                }
                else
                {
                    var sheets = new Dictionary<string, object>();

                    if (options.WriteUniqueSheet || isInPlace)
                    {
                        if (options.PreserveSourceSheets && uniquePerSheet.Count > 0)
                        {
                            // One output sheet per input sheet, in the order they were loaded.
                            foreach (var sheetName in orderedSheetKeys)
                            {
                                if (!uniquePerSheet.TryGetValue(sheetName, out var bag)) continue;
                                var safeName = SanitizeSheetName(sheetName, sheets.Keys);
                                sheets[safeName] = NormalizeRows(bag, finalColumns).ToList();
                            }
                        }
                        else
                        {
                            // Merge everything into one sheet — preserve load order across sheets.
                            var merged = orderedSheetKeys
                                .Where(k => uniquePerSheet.ContainsKey(k))
                                .SelectMany(k => uniquePerSheet[k]);
                            sheets[options.UniqueSheetName] = NormalizeRows(merged, finalColumns).ToList();
                        }
                    }

                    // Duplicates and Invalid always merge into single named sheets — cleaner output.
                    // A "_Source" column is added so users can trace the original sheet of each row.
                    // Invalid sheet additionally gets a "_Reason" column explaining why each row was rejected.
                    if (options.WriteDuplicatesSheet)
                    {
                        var mergedDup = FlattenWithSource(duplicatesPerSheet, orderedSheetKeys);
                        sheets[options.DuplicatesSheetName] = NormalizeRowsWithSource(mergedDup, finalColumns);
                    }

                    if (options.WriteInvalidSheet)
                    {
                        var mergedInv = FlattenWithSource(invalidPerSheet, orderedSheetKeys);
                        sheets[options.InvalidSheetName] = NormalizeRowsWithSource(mergedInv, finalColumns, includeReason: true);
                    }

                    if (sheets.Count == 0)
                    {
                        var merged = orderedSheetKeys
                            .Where(k => uniquePerSheet.ContainsKey(k))
                            .SelectMany(k => uniquePerSheet[k]);
                        sheets[options.UniqueSheetName] = NormalizeRows(merged, finalColumns).ToList();
                    }

                    MiniExcel.SaveAs(primaryOutputFile, sheets, overwriteFile: true);
                }
            }, cancellationToken);

            var outputFile = primaryOutputFile;

            sw.Stop();

            if (uniquePerSheet.Count > 1 && options.PreserveSourceSheets)
            {
                var sheetSummary = string.Join(", ",
                    orderedKeysForLog
                        .Where(k => uniquePerSheet.ContainsKey(k))
                        .Select(k => $"{k} ({uniquePerSheet[k].Count:N0})"));
                log(LogLevel.Info, $"Output sheets: {sheetSummary}");
            }

            log(LogLevel.Success,
                $"Wrote {totalUnique:N0} unique rows to {Path.GetFileName(outputFile)} in {sw.Elapsed.TotalSeconds:0.0}s.");

            return new ProcessingResult
            {
                Success = true,
                OutputPath = outputFile,
                TotalRowsRead = Interlocked.Read(ref totalRowsRead),
                DuplicatesRemoved = Interlocked.Read(ref duplicatesRemoved),
                InvalidRemoved = Interlocked.Read(ref invalidRemoved),
                ValidRecords = totalUnique,
                Elapsed = sw.Elapsed
            };
        }
        catch (OperationCanceledException)
        {
            sw.Stop();
            log(LogLevel.Warning, "Processing cancelled.");
            return new ProcessingResult { Success = false, ErrorMessage = "Cancelled", Elapsed = sw.Elapsed };
        }
        catch (Exception ex)
        {
            sw.Stop();
            log(LogLevel.Error, $"Fatal: {ex.Message}");
            return new ProcessingResult { Success = false, ErrorMessage = ex.Message, Elapsed = sw.Elapsed };
        }
    }

    /// <summary>
    /// Reads workbook headers without loading the data rows. Used by the UI to populate column pickers.
    /// </summary>
    public async Task<IReadOnlyList<string>> ScanHeadersAsync(IReadOnlyList<FileItem> files, ProcessingOptions options, CancellationToken ct = default)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ordered = new List<string>();

        await Task.Run(() =>
        {
            foreach (var file in files)
            {
                if (ct.IsCancellationRequested) return;
                if (!File.Exists(file.FullPath)) continue;

                var isCsv = IsCsvFile(file.FullPath);
                List<string> sheetsToRead;

                if (isCsv)
                {
                    sheetsToRead = new List<string> { "" };
                }
                else
                {
                    IReadOnlyList<string> sheetsInFile;
                    try { sheetsInFile = MiniExcel.GetSheetNames(file.FullPath).ToList(); }
                    catch { continue; }

                    sheetsToRead = ResolveSheetsToRead(sheetsInFile, options);
                    if (sheetsToRead.Count == 0) sheetsToRead = sheetsInFile.Take(1).ToList();
                }

                foreach (var sheet in sheetsToRead)
                {
                    try
                    {
                        var query = isCsv
                            ? MiniExcel.Query(file.FullPath, useHeaderRow: true, excelType: MiniExcelLibs.ExcelType.CSV)
                            : MiniExcel.Query(file.FullPath, useHeaderRow: true, sheetName: sheet);

                        var firstRow = query.Cast<IDictionary<string, object>>().FirstOrDefault();
                        if (firstRow is null) continue;
                        foreach (var rawCol in firstRow.Keys)
                        {
                            if (rawCol is null) continue;
                            var col = rawCol.Trim();
                            if (col.Length == 0) continue;
                            if (seen.Add(col)) ordered.Add(col);
                        }
                    }
                    catch
                    {
                        // Skip unreadable sheets silently — the main processing pass will surface the error.
                    }
                }
            }
        }, ct);

        return ordered;
    }

    private static string DescribeDedup(ProcessingOptions o) => o.DedupMode switch
    {
        DedupKeyMode.None => "off",
        DedupKeyMode.EntireRow => "entire row",
        DedupKeyMode.MultipleColumns => $"columns [{string.Join(", ", o.DedupColumns)}]",
        _ => o.DedupColumns.Length > 0 ? $"column '{o.DedupColumns[0]}'" : "single column"
    };

    private static List<Dictionary<string, object?>> ExpandMultiValueRow(
        Dictionary<string, object?> row, ProcessingOptions options)
    {
        // Only single-column or multi-column dedup modes need splitting; whole-row/none don't.
        if (options.DedupMode != DedupKeyMode.SingleColumn && options.DedupMode != DedupKeyMode.MultipleColumns)
            return new List<Dictionary<string, object?>> { row };

        if (options.DedupColumns.Length == 0)
            return new List<Dictionary<string, object?>> { row };

        var delimiters = (options.MultiValueDelimiters ?? "").ToCharArray();
        if (delimiters.Length == 0)
            return new List<Dictionary<string, object?>> { row };

        // Find the actual key in the row matching the (case-insensitive) dedup column name.
        // Only split the FIRST dedup column. Other columns stay untouched.
        var targetCol = options.DedupColumns[0];
        var actualKey = row.Keys.FirstOrDefault(k =>
            string.Equals(k, targetCol, StringComparison.OrdinalIgnoreCase));
        if (actualKey is null)
            return new List<Dictionary<string, object?>> { row };

        var raw = row[actualKey]?.ToString();
        if (string.IsNullOrWhiteSpace(raw))
            return new List<Dictionary<string, object?>> { row };

        var parts = raw.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToArray();

        if (parts.Length <= 1)
            return new List<Dictionary<string, object?>> { row };

        var expanded = new List<Dictionary<string, object?>>(parts.Length);
        foreach (var part in parts)
        {
            var clone = new Dictionary<string, object?>(row, StringComparer.Ordinal)
            {
                [actualKey] = part
            };
            expanded.Add(clone);
        }
        return expanded;
    }

    private static bool IsRowEmpty(IReadOnlyDictionary<string, object?> row)
    {
        foreach (var v in row.Values)
        {
            if (v is null) continue;
            var s = v.ToString();
            if (!string.IsNullOrWhiteSpace(s)) return false;
        }
        return true;
    }

    private static Dictionary<string, object?> SnapshotRow(IDictionary<string, object> dict)
    {
        var snapshot = new Dictionary<string, object?>(dict.Count, StringComparer.Ordinal);
        foreach (var kvp in dict)
        {
            snapshot[kvp.Key] = kvp.Value;
        }
        return snapshot;
    }

    private static string Normalize(string value, bool caseSensitive)
        => caseSensitive ? value.Trim() : value.Trim().ToLowerInvariant();

    private static string BuildDedupKey(IReadOnlyDictionary<string, object?> row, ProcessingOptions options)
    {
        switch (options.DedupMode)
        {
            case DedupKeyMode.EntireRow:
            {
                var parts = row.Values.Select(v => Normalize(v?.ToString() ?? "", options.CaseSensitive));
                return string.Join("", parts);
            }
            case DedupKeyMode.MultipleColumns:
            {
                var parts = options.DedupColumns
                    .Select(c => Normalize(LookupColumn(row, c) ?? "", options.CaseSensitive));
                return string.Join("", parts);
            }
            case DedupKeyMode.SingleColumn:
            default:
            {
                var col = options.DedupColumns.Length > 0 ? options.DedupColumns[0] : "";
                return Normalize(LookupColumn(row, col) ?? "", options.CaseSensitive);
            }
        }
    }

    private static string? LookupColumn(IReadOnlyDictionary<string, object?> row, string columnName)
    {
        if (string.IsNullOrEmpty(columnName)) return null;

        // Exact (case-insensitive, trimmed) match — no substring fallback to avoid
        // accidentally picking up 'firstname' when user asks for 'name'.
        var target = columnName.Trim();
        foreach (var kvp in row)
        {
            if (kvp.Key is not null &&
                string.Equals(kvp.Key.Trim(), target, StringComparison.OrdinalIgnoreCase))
                return kvp.Value?.ToString();
        }
        return null;
    }

    private static (bool Ok, string Reason) Validate(string value, ValidationMode mode, string column)
    {
        if (mode == ValidationMode.None) return (true, "");

        if (mode.HasFlag(ValidationMode.NotEmpty) && string.IsNullOrWhiteSpace(value))
        {
            return (false, $"'{column}' is empty");
        }

        if (mode.HasFlag(ValidationMode.Email))
        {
            if (string.IsNullOrEmpty(value))
                return (false, $"'{column}' is empty (expected email)");
            if (value.Length < 5 || value.Length > 320)
                return (false, $"'{column}' length out of range");
            if (!EmailRegex.IsMatch(value))
                return (false, $"'{column}' is not a valid email format");
        }

        return (true, "");
    }

    private static IEnumerable<(string Source, Dictionary<string, object?> Row)> FlattenWithSource(
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> bucket,
        IReadOnlyList<string> orderedKeys)
    {
        // First, emit rows in the canonical load order.
        var emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in orderedKeys)
        {
            if (!bucket.TryGetValue(key, out var bag)) continue;
            emitted.Add(key);
            foreach (var row in bag) yield return (key, row);
        }

        // Any keys that exist in the bucket but not in orderedKeys (shouldn't normally happen) — emit last.
        foreach (var (sheetName, bag) in bucket)
        {
            if (emitted.Contains(sheetName)) continue;
            foreach (var row in bag) yield return (sheetName, row);
        }
    }

    private static List<Dictionary<string, object?>> NormalizeRowsWithSource(
        IEnumerable<(string Source, Dictionary<string, object?> Row)> rowsWithSource,
        List<string> finalColumns,
        bool includeReason = false)
    {
        var prefixCols = includeReason
            ? new[] { SourceColumnName, ReasonColumnName }
            : new[] { SourceColumnName };

        var allCols = new List<string>(prefixCols.Length + finalColumns.Count);
        allCols.AddRange(prefixCols);
        allCols.AddRange(finalColumns);

        var result = new List<Dictionary<string, object?>>();
        foreach (var (source, row) in rowsWithSource)
        {
            var normalized = new Dictionary<string, object?>(allCols.Count, StringComparer.Ordinal)
            {
                [SourceColumnName] = source
            };
            if (includeReason)
            {
                normalized[ReasonColumnName] = row.TryGetValue(ReasonColumnName, out var r) ? r : null;
            }
            foreach (var col in finalColumns)
            {
                if (col == ReasonColumnName) continue; // already handled above
                normalized[col] = LookupColumnValue(row, col);
            }
            result.Add(normalized);
        }

        if (result.Count == 0 && allCols.Count > 0)
        {
            var headerOnly = new Dictionary<string, object?>(allCols.Count, StringComparer.Ordinal);
            foreach (var col in allCols) headerOnly[col] = null;
            result.Add(headerOnly);
        }

        return result;
    }

    private static List<Dictionary<string, object?>> NormalizeRows(
        IEnumerable<Dictionary<string, object?>> rows,
        List<string> finalColumns)
    {
        var result = new List<Dictionary<string, object?>>();
        foreach (var row in rows)
        {
            var normalized = new Dictionary<string, object?>(finalColumns.Count, StringComparer.Ordinal);
            foreach (var col in finalColumns)
            {
                normalized[col] = LookupColumnValue(row, col);
            }
            result.Add(normalized);
        }

        // If no rows, emit a single header-only row of nulls so the sheet is created with column headers visible.
        if (result.Count == 0 && finalColumns.Count > 0)
        {
            var headerOnly = new Dictionary<string, object?>(finalColumns.Count, StringComparer.Ordinal);
            foreach (var col in finalColumns) headerOnly[col] = null;
            result.Add(headerOnly);
        }

        return result;
    }

    /// <summary>Looks up a column value with case-insensitive trimmed match — handles ' Email ' vs 'Email' from messy source headers.</summary>
    private static object? LookupColumnValue(IReadOnlyDictionary<string, object?> row, string column)
    {
        if (row.TryGetValue(column, out var v)) return v;
        var target = column.Trim();
        foreach (var kvp in row)
        {
            if (kvp.Key is not null &&
                string.Equals(kvp.Key.Trim(), target, StringComparison.OrdinalIgnoreCase))
                return kvp.Value;
        }
        return null;
    }

    private static ProgressUpdate BuildProgress(
        int filesDone, int fileTotal, string? currentFile,
        long totalRowsRead, long duplicatesRemoved, long invalidRemoved,
        long validRecords)
    {
        var pct = fileTotal == 0 ? 0.0 : Math.Min(100.0, (double)filesDone / fileTotal * 100.0);
        return new ProgressUpdate(filesDone, fileTotal, currentFile,
            totalRowsRead, duplicatesRemoved, invalidRemoved, validRecords, pct);
    }

    private static bool IsCsvFile(string path)
        => string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);

    private static IReadOnlyList<string> SafeGetSheetNames(string path)
    {
        try
        {
            return MiniExcel.GetSheetNames(path).ToList();
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    private static List<string> ResolveSheetsToRead(IReadOnlyList<string> sheetsInFile, ProcessingOptions options)
    {
        switch (options.SheetMode)
        {
            case SheetSelectionMode.All:
                return sheetsInFile.ToList();

            case SheetSelectionMode.FirstN:
            {
                var n = Math.Max(1, options.FirstNSheets);
                return sheetsInFile.Take(n).ToList();
            }

            case SheetSelectionMode.ByName:
            default:
            {
                var picks = options.SheetNames ?? Array.Empty<string>();
                if (picks.Length == 0) return new List<string>();

                var matched = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var sheet in sheetsInFile)
                {
                    foreach (var pattern in picks)
                    {
                        if (string.IsNullOrWhiteSpace(pattern)) continue;
                        if (MatchesSheet(sheet, pattern.Trim()))
                        {
                            if (seen.Add(sheet)) matched.Add(sheet);
                            break;
                        }
                    }
                }
                return matched;
            }
        }
    }

    private static long CountAllRows(ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> map)
    {
        long total = 0;
        foreach (var bag in map.Values) total += bag.Count;
        return total;
    }

    private static readonly char[] InvalidSheetChars = { '\\', '/', '?', '*', '[', ']', ':' };

    private static string SanitizeSheetName(string name, IEnumerable<string> existing)
    {
        var trimmed = string.IsNullOrWhiteSpace(name) ? "Sheet" : name.Trim();
        foreach (var c in InvalidSheetChars) trimmed = trimmed.Replace(c, '_');
        if (trimmed.Length > 31) trimmed = trimmed.Substring(0, 31);

        var existingSet = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
        if (!existingSet.Contains(trimmed)) return trimmed;

        // Append a numeric suffix until unique. Excel sheet names must be ≤31 chars total.
        for (var i = 2; i < 1000; i++)
        {
            var suffix = $" ({i})";
            var maxBase = 31 - suffix.Length;
            var baseName = trimmed.Length > maxBase ? trimmed.Substring(0, maxBase) : trimmed;
            var candidate = baseName + suffix;
            if (!existingSet.Contains(candidate)) return candidate;
        }
        return trimmed; // give up — collision is unlikely past 1000 attempts
    }

    private static string SanitizeFilename(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "sheet";
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new System.Text.StringBuilder(name.Length);
        foreach (var c in name)
        {
            sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
        }
        return sb.ToString();
    }

    private static bool MatchesSheet(string sheetName, string pattern)
    {
        if (!pattern.Contains('*'))
            return string.Equals(sheetName, pattern, StringComparison.OrdinalIgnoreCase);

        var rx = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
        return Regex.IsMatch(sheetName, rx, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }
}
