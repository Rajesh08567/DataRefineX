using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DataRefineX.Models;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;

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

public enum SplitMode
{
    /// <summary>Don't split — write everything as-is.</summary>
    None,
    /// <summary>Split each output sheet into chunks of N rows. Sheet 'Customers' (25k rows, N=10k) → 'Customers_1', 'Customers_2', 'Customers_3'.</summary>
    PerSheet,
    /// <summary>Merge all unique rows, then split into 'Unique_1', 'Unique_2', ... of N rows each.</summary>
    MergedSheets,
    /// <summary>Merge all unique rows, then split into separate files: 'Processed_..._1.xlsx', 'Processed_..._2.xlsx'.</summary>
    SeparateFiles
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

    /// <summary>How to split large output. Default None = no splitting.</summary>
    public SplitMode SplitMode { get; init; } = SplitMode.None;
    /// <summary>Max rows per output chunk when SplitMode != None. Min 1, default 10000.</summary>
    public int SplitSize { get; init; } = 10_000;

    /// <summary>When true, output xlsx sheets get header filter dropdowns (Excel Table styling). Off by default — most users don't want them.</summary>
    public bool IncludeFilters { get; init; } = false;
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
            if (options.SplitMode != SplitMode.None)
            {
                log(LogLevel.Info, $"Output split: {options.SplitMode} every {options.SplitSize:N0} rows.");
            }

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

            string outputFile = primaryOutputFile;
            var extraOutputFiles = new List<string>();

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

                // Splitting is disabled in in-place mode — overwriting source with N split files would be confusing.
                var effectiveSplit = isInPlace ? SplitMode.None : options.SplitMode;
                var splitSize = Math.Max(1, options.SplitSize);

                var writeResult = writeAsCsv
                    ? WriteCsvOutput(
                        primaryOutputFile, basePath, options, isInPlace, effectiveSplit, splitSize,
                        uniquePerSheet, duplicatesPerSheet, invalidPerSheet,
                        orderedSheetKeys, finalColumns)
                    : WriteXlsxOutput(
                        primaryOutputFile, basePath, options, isInPlace, effectiveSplit, splitSize,
                        uniquePerSheet, duplicatesPerSheet, invalidPerSheet,
                        orderedSheetKeys, finalColumns);

                outputFile = writeResult.PrimaryFile;
                extraOutputFiles.AddRange(writeResult.ExtraFiles);
            }, cancellationToken);

            if (extraOutputFiles.Count > 0)
            {
                log(LogLevel.Info, $"Split output: {1 + extraOutputFiles.Count} files written.");
            }

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

    // ---------------- Split + gap report helpers ----------------

    private readonly record struct WriteResult(string PrimaryFile, IReadOnlyList<string> ExtraFiles);

    private static IEnumerable<List<T>> Chunk<T>(IEnumerable<T> source, int size)
    {
        if (size <= 0) size = 1;
        var chunk = new List<T>(size);
        foreach (var item in source)
        {
            chunk.Add(item);
            if (chunk.Count >= size)
            {
                yield return chunk;
                chunk = new List<T>(size);
            }
        }
        if (chunk.Count > 0) yield return chunk;
    }

    // MiniExcel 1.x writes every string cell as inline `t="str"` (formula-result type) with empty
    // sharedStrings.xml AND writes [Content_Types].xml / xl/workbook.xml / xl/styles.xml with a
    // UTF-8 BOM before the XML declaration. Strict OOXML parsers (PhpSpreadsheet, libxml2) reject
    // the file outright: they parse [Content_Types].xml first and choke on EF BB BF before <?xml,
    // OR they read t="str" cells as null. Excel resave fixes both because it rewrites everything.
    //
    // This post-pass does two things:
    //   1) Strip the UTF-8 BOM from every .xml part in the zip.
    //   2) Migrate t="str" cells in worksheets to t="s" with proper sharedStrings entries.
    private static void NormalizeXlsxSharedStrings(string xlsxPath)
    {
        try
        {
            if (!File.Exists(xlsxPath)) return;

            // ---- Pass 1: read everything ----
            var strings = new List<string>();
            var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            var sheetRewrites = new Dictionary<string, string>(StringComparer.Ordinal);
            var bomFiles = new List<string>();
            var anyChanges = false;

            using (var zip = ZipFile.Open(xlsxPath, ZipArchiveMode.Read))
            {
                foreach (var entry in zip.Entries)
                {
                    // Detect BOM on every .xml / .rels part — they must NOT have one per OOXML spec.
                    if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal) ||
                        entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
                    {
                        using var bs = entry.Open();
                        var first3 = new byte[3];
                        var read = bs.Read(first3, 0, 3);
                        if (read == 3 && first3[0] == 0xEF && first3[1] == 0xBB && first3[2] == 0xBF)
                        {
                            bomFiles.Add(entry.FullName);
                            anyChanges = true;
                        }
                    }

                    // Convert t="str" cells in worksheet xml.
                    if (entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) &&
                        entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                    {
                        string xml;
                        using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                            xml = sr.ReadToEnd();

                        var rewritten = RewriteSheetInlineStrings(xml, strings, stringIndex, out var changed);
                        if (changed)
                        {
                            sheetRewrites[entry.FullName] = rewritten;
                            anyChanges = true;
                        }
                    }
                }
            }

            if (!anyChanges) return;

            // ---- Pass 2: rewrite ----
            using var zipWrite = ZipFile.Open(xlsxPath, ZipArchiveMode.Update);

            // 2a) Strip BOM from every offending entry.
            foreach (var entryName in bomFiles)
            {
                var entry = zipWrite.GetEntry(entryName);
                if (entry is null) continue;
                byte[] bytes;
                using (var es = entry.Open())
                using (var ms = new MemoryStream())
                {
                    es.CopyTo(ms);
                    bytes = ms.ToArray();
                }
                // Strip leading EF BB BF.
                if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                {
                    var stripped = new byte[bytes.Length - 3];
                    Array.Copy(bytes, 3, stripped, 0, stripped.Length);
                    bytes = stripped;
                }
                // Don't double-write if this entry is also being replaced by sheet-rewrite — skip,
                // sheet pass writes a fresh BOM-free copy below.
                if (sheetRewrites.ContainsKey(entryName)) continue;
                entry.Delete();
                var fresh = zipWrite.CreateEntry(entryName, CompressionLevel.Optimal);
                using var ws = fresh.Open();
                ws.Write(bytes, 0, bytes.Length);
            }

            // 2b) Replace worksheets with shared-string-migrated content (always BOM-free).
            foreach (var (entryName, newXml) in sheetRewrites)
            {
                var entry = zipWrite.GetEntry(entryName);
                if (entry is null) continue;
                entry.Delete();
                var fresh = zipWrite.CreateEntry(entryName, CompressionLevel.Optimal);
                using var ws = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
                ws.Write(newXml);
            }

            // 2c) Replace sharedStrings.xml only if we migrated any cells.
            if (sheetRewrites.Count > 0)
            {
                var sstEntry = zipWrite.GetEntry("xl/sharedStrings.xml");
                sstEntry?.Delete();
                var sst = zipWrite.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Optimal);
                using var sw = new StreamWriter(sst.Open(), new UTF8Encoding(false));
                sw.Write(BuildSharedStringsXml(strings));
            }
        }
        catch
        {
            // Post-processing is best-effort. If anything goes wrong (locked file, malformed xml,
            // unexpected schema), leave the MiniExcel output untouched — it still opens in Excel.
        }
    }

    private static string RewriteSheetInlineStrings(
        string sheetXml,
        List<string> strings,
        Dictionary<string, int> stringIndex,
        out bool changed)
    {
        changed = false;
        var doc = new XmlDocument { PreserveWhitespace = false };
        try { doc.LoadXml(sheetXml); }
        catch { return sheetXml; }

        var ns = new XmlNamespaceManager(doc.NameTable);
        const string mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        ns.AddNamespace("x", mainNs);

        var cells = doc.SelectNodes("//x:c[@t='str']", ns);
        if (cells is null || cells.Count == 0) return sheetXml;

        foreach (XmlNode cell in cells)
        {
            // Skip if cell has a formula — t="str" is legitimate there.
            if (cell.SelectSingleNode("x:f", ns) is not null) continue;

            var vNode = cell.SelectSingleNode("x:v", ns);
            if (vNode is null) continue;

            var value = vNode.InnerText ?? "";

            if (!stringIndex.TryGetValue(value, out var idx))
            {
                idx = strings.Count;
                strings.Add(value);
                stringIndex[value] = idx;
            }

            var attrs = cell.Attributes!;
            var tAttr = attrs["t"]!;
            tAttr.Value = "s";

            vNode.InnerText = idx.ToString(System.Globalization.CultureInfo.InvariantCulture);
            changed = true;
        }

        if (!changed) return sheetXml;

        // StringWriter is UTF-16; using a MemoryStream + UTF-8 writer keeps the declaration honest
        // so strict parsers don't trip on a utf-16 declaration in a file we'll write as UTF-8 bytes.
        using var ms = new MemoryStream();
        using (var xw = XmlWriter.Create(ms, new XmlWriterSettings
        {
            OmitXmlDeclaration = false,
            Indent = false,
            Encoding = new UTF8Encoding(false)
        }))
        {
            doc.Save(xw);
        }
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static string BuildSharedStringsXml(IReadOnlyList<string> strings)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
        sb.Append($"count=\"{strings.Count}\" uniqueCount=\"{strings.Count}\">");
        foreach (var s in strings)
        {
            sb.Append("<si><t");
            // Preserve leading/trailing whitespace, which strict readers otherwise collapse.
            if (s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[s.Length - 1])))
                sb.Append(" xml:space=\"preserve\"");
            sb.Append('>');
            sb.Append(XmlEscape(s));
            sb.Append("</t></si>");
        }
        sb.Append("</sst>");
        return sb.ToString();
    }

    private static string XmlEscape(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            switch (c)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&apos;"); break;
                default:
                    // Strip control chars that are illegal in XML 1.0 (except tab/lf/cr).
                    if (c < 0x20 && c != '\t' && c != '\n' && c != '\r') continue;
                    sb.Append(c);
                    break;
            }
        }
        return sb.ToString();
    }

    private static WriteResult WriteXlsxOutput(
        string primaryOutputFile,
        string basePath,
        ProcessingOptions options,
        bool isInPlace,
        SplitMode splitMode,
        int splitSize,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> uniquePerSheet,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> duplicatesPerSheet,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> invalidPerSheet,
        IReadOnlyList<string> orderedSheetKeys,
        List<string> finalColumns)
    {
        var extras = new List<string>();
        // MiniExcel defaults each sheet to an Excel Table (TableStyleMedium2 + AutoFilter),
        // which shows filter dropdowns on every column header. Only enable when the user asks for it.
        var xlsxConfig = new OpenXmlConfiguration
        {
            TableStyles = options.IncludeFilters ? TableStyles.Default : TableStyles.None,
            AutoFilter = options.IncludeFilters
        };

        // Build duplicate/invalid sheets once — they don't get split (audit data).
        List<Dictionary<string, object?>>? duplicateRows = null;
        List<Dictionary<string, object?>>? invalidRows = null;
        if (options.WriteDuplicatesSheet)
        {
            duplicateRows = NormalizeRowsWithSource(
                FlattenWithSource(duplicatesPerSheet, orderedSheetKeys), finalColumns);
        }
        if (options.WriteInvalidSheet)
        {
            invalidRows = NormalizeRowsWithSource(
                FlattenWithSource(invalidPerSheet, orderedSheetKeys), finalColumns, includeReason: true);
        }

        // SeparateFiles split: each chunk is its own .xlsx file containing one Unique sheet (+ duplicates/invalid copied to first file only).
        if (splitMode == SplitMode.SeparateFiles)
        {
            var wantUnique = options.WriteUniqueSheet || isInPlace;

            // If user disabled the unique sheet, splitting it makes no sense — emit a single audit-only file.
            if (!wantUnique)
            {
                var auditSheets = new Dictionary<string, object>();
                if (duplicateRows is not null) auditSheets[options.DuplicatesSheetName] = duplicateRows;
                if (invalidRows is not null) auditSheets[options.InvalidSheetName] = invalidRows;
                if (auditSheets.Count == 0) auditSheets[options.UniqueSheetName] = new List<Dictionary<string, object?>>();
                MiniExcel.SaveAs(primaryOutputFile, auditSheets, overwriteFile: true, configuration: xlsxConfig);
                NormalizeXlsxSharedStrings(primaryOutputFile);
                return new WriteResult(primaryOutputFile, extras);
            }

            var merged = orderedSheetKeys
                .Where(k => uniquePerSheet.ContainsKey(k))
                .SelectMany(k => uniquePerSheet[k]);

            var chunks = Chunk(merged, splitSize).ToList();
            if (chunks.Count == 0) chunks.Add(new List<Dictionary<string, object?>>());

            string firstFile = primaryOutputFile;
            for (var i = 0; i < chunks.Count; i++)
            {
                var path = chunks.Count == 1
                    ? primaryOutputFile
                    : $"{basePath}_part{i + 1:D2}.xlsx";

                var sheets = new Dictionary<string, object>
                {
                    [options.UniqueSheetName] = NormalizeRows(chunks[i], finalColumns).ToList()
                };
                // Duplicates / invalid only in the FIRST file — keeps audit data in one place.
                if (i == 0)
                {
                    if (duplicateRows is not null) sheets[options.DuplicatesSheetName] = duplicateRows;
                    if (invalidRows is not null) sheets[options.InvalidSheetName] = invalidRows;
                }

                MiniExcel.SaveAs(path, sheets, overwriteFile: true, configuration: xlsxConfig);
                NormalizeXlsxSharedStrings(path);
                if (i == 0) firstFile = path;
                else extras.Add(path);
            }
            return new WriteResult(firstFile, extras);
        }

        // Single-file output (None / PerSheet / MergedSheets).
        var allSheets = new Dictionary<string, object>();

        if (options.WriteUniqueSheet || isInPlace)
        {
            if (splitMode == SplitMode.MergedSheets)
            {
                var merged = orderedSheetKeys
                    .Where(k => uniquePerSheet.ContainsKey(k))
                    .SelectMany(k => uniquePerSheet[k]);
                var chunks = Chunk(merged, splitSize).ToList();
                if (chunks.Count == 0) chunks.Add(new List<Dictionary<string, object?>>());
                for (var i = 0; i < chunks.Count; i++)
                {
                    var name = chunks.Count == 1
                        ? options.UniqueSheetName
                        : $"{options.UniqueSheetName}_{i + 1}";
                    var safe = SanitizeSheetName(name, allSheets.Keys);
                    allSheets[safe] = NormalizeRows(chunks[i], finalColumns).ToList();
                }
            }
            else if (splitMode == SplitMode.PerSheet && options.PreserveSourceSheets)
            {
                foreach (var sheetName in orderedSheetKeys)
                {
                    if (!uniquePerSheet.TryGetValue(sheetName, out var bag)) continue;
                    var chunks = Chunk(bag, splitSize).ToList();
                    if (chunks.Count == 0) chunks.Add(new List<Dictionary<string, object?>>());
                    for (var i = 0; i < chunks.Count; i++)
                    {
                        var name = chunks.Count == 1 ? sheetName : $"{sheetName}_{i + 1}";
                        var safe = SanitizeSheetName(name, allSheets.Keys);
                        allSheets[safe] = NormalizeRows(chunks[i], finalColumns).ToList();
                    }
                }
            }
            else
            {
                // None (or PerSheet without PreserveSourceSheets — falls back to no split).
                if (options.PreserveSourceSheets && uniquePerSheet.Count > 0)
                {
                    foreach (var sheetName in orderedSheetKeys)
                    {
                        if (!uniquePerSheet.TryGetValue(sheetName, out var bag)) continue;
                        var safe = SanitizeSheetName(sheetName, allSheets.Keys);
                        allSheets[safe] = NormalizeRows(bag, finalColumns).ToList();
                    }
                }
                else
                {
                    var merged = orderedSheetKeys
                        .Where(k => uniquePerSheet.ContainsKey(k))
                        .SelectMany(k => uniquePerSheet[k]);
                    allSheets[options.UniqueSheetName] = NormalizeRows(merged, finalColumns).ToList();
                }
            }
        }

        if (duplicateRows is not null) allSheets[options.DuplicatesSheetName] = duplicateRows;
        if (invalidRows is not null) allSheets[options.InvalidSheetName] = invalidRows;

        if (allSheets.Count == 0)
        {
            var merged = orderedSheetKeys
                .Where(k => uniquePerSheet.ContainsKey(k))
                .SelectMany(k => uniquePerSheet[k]);
            allSheets[options.UniqueSheetName] = NormalizeRows(merged, finalColumns).ToList();
        }

        MiniExcel.SaveAs(primaryOutputFile, allSheets, overwriteFile: true, configuration: xlsxConfig);
        NormalizeXlsxSharedStrings(primaryOutputFile);
        return new WriteResult(primaryOutputFile, extras);
    }

    private static WriteResult WriteCsvOutput(
        string primaryOutputFile,
        string basePath,
        ProcessingOptions options,
        bool isInPlace,
        SplitMode splitMode,
        int splitSize,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> uniquePerSheet,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> duplicatesPerSheet,
        ConcurrentDictionary<string, ConcurrentBag<Dictionary<string, object?>>> invalidPerSheet,
        IReadOnlyList<string> orderedSheetKeys,
        List<string> finalColumns)
    {
        var extras = new List<string>();
        string firstFile = primaryOutputFile;
        var firstAssigned = false;

        void Track(string path)
        {
            if (!firstAssigned) { firstFile = path; firstAssigned = true; }
            else extras.Add(path);
        }

        // Splitting writes one CSV per chunk regardless of mode (CSV is single-table).
        if (options.WriteUniqueSheet || isInPlace)
        {
            if (isInPlace && uniquePerSheet.Count == 1 && splitMode == SplitMode.None)
            {
                var only = uniquePerSheet.Values.First();
                MiniExcel.SaveAs(primaryOutputFile,
                    NormalizeRows(only, finalColumns).ToList(),
                    excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                Track(primaryOutputFile);
            }
            else if (splitMode == SplitMode.PerSheet)
            {
                foreach (var sheetName in orderedSheetKeys)
                {
                    if (!uniquePerSheet.TryGetValue(sheetName, out var bag)) continue;
                    var chunks = Chunk(bag, splitSize).ToList();
                    if (chunks.Count == 0) chunks.Add(new List<Dictionary<string, object?>>());
                    for (var i = 0; i < chunks.Count; i++)
                    {
                        var safe = SanitizeFilename(sheetName);
                        var path = chunks.Count == 1
                            ? $"{basePath}_{safe}.csv"
                            : $"{basePath}_{safe}_{i + 1:D2}.csv";
                        MiniExcel.SaveAs(path,
                            NormalizeRows(chunks[i], finalColumns).ToList(),
                            excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                        Track(path);
                    }
                }
            }
            else if (splitMode == SplitMode.MergedSheets || splitMode == SplitMode.SeparateFiles)
            {
                var merged = orderedSheetKeys
                    .Where(k => uniquePerSheet.ContainsKey(k))
                    .SelectMany(k => uniquePerSheet[k]);
                var chunks = Chunk(merged, splitSize).ToList();
                if (chunks.Count == 0) chunks.Add(new List<Dictionary<string, object?>>());
                for (var i = 0; i < chunks.Count; i++)
                {
                    var path = chunks.Count == 1
                        ? $"{basePath}_{options.UniqueSheetName}.csv"
                        : $"{basePath}_{options.UniqueSheetName}_{i + 1:D2}.csv";
                    MiniExcel.SaveAs(path,
                        NormalizeRows(chunks[i], finalColumns).ToList(),
                        excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
                    Track(path);
                }
            }
            else
            {
                // SplitMode.None: original per-sheet CSV behavior.
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
                    Track(path);
                }
            }
        }

        if (options.WriteDuplicatesSheet)
        {
            var path = $"{basePath}_{options.DuplicatesSheetName}.csv";
            var mergedDup = FlattenWithSource(duplicatesPerSheet, orderedSheetKeys);
            MiniExcel.SaveAs(path,
                NormalizeRowsWithSource(mergedDup, finalColumns),
                excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
            Track(path);
        }
        if (options.WriteInvalidSheet)
        {
            var path = $"{basePath}_{options.InvalidSheetName}.csv";
            var mergedInv = FlattenWithSource(invalidPerSheet, orderedSheetKeys);
            MiniExcel.SaveAs(path,
                NormalizeRowsWithSource(mergedInv, finalColumns, includeReason: true),
                excelType: MiniExcelLibs.ExcelType.CSV, overwriteFile: true);
            Track(path);
        }

        if (!firstAssigned)
        {
            // Defensive fallback — should be unreachable.
            firstFile = primaryOutputFile;
        }
        return new WriteResult(firstFile, extras);
    }
}
