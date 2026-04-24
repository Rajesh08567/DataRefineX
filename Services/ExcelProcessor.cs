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

public sealed class ProcessingOptions
{
    public string EmailColumnName { get; init; } = "email";
    public string NameColumnName { get; init; } = "name";
    public string[] SheetNames { get; init; } = { "Valid+BasicCheck+DEA", "CatchAll_AcceptAll" };
    public string OutputSheetName { get; init; } = "Valid+BasicCheck+DEA";
    public string? OutputDirectory { get; init; }
    public int MaxDegreeOfParallelism { get; init; } = Math.Max(2, Environment.ProcessorCount);
}

public sealed class ExcelProcessor
{
    // Basic, fast email validation — intentionally permissive but rejects obvious garbage.
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
        var sw = Stopwatch.StartNew();

        try
        {
            if (files.Count == 0)
            {
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No files to process."
                };
            }

            log(LogLevel.Info, $"Starting processing of {files.Count} file(s) with {options.MaxDegreeOfParallelism} workers.");

            var totalRowsRead = 0L;
            var duplicatesRemoved = 0L;
            var invalidRemoved = 0L;
            var filesDoneCount = 0;

            // Preserve column order: union across files, with "email" first, then first-seen order from each file.
            var columnOrder = new List<string>();
            var columnOrderLock = new object();
            var columnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Dedupe key = normalized email (lowercase, trimmed). First-seen row wins.
            var seenEmails = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

            // Accumulator for the final valid, deduped rows.
            var acceptedRows = new ConcurrentBag<Dictionary<string, object?>>();

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = options.MaxDegreeOfParallelism,
                CancellationToken = cancellationToken
            };

            await Parallel.ForEachAsync(files, parallelOptions, async (file, ct) =>
            {
                if (ct.IsCancellationRequested) return;

                file.Status = FileStatus.Reading;
                progress.Report(BuildProgress(
                    filesDoneCount, files.Count, file.FileName,
                    totalRowsRead, duplicatesRemoved, invalidRemoved,
                    acceptedRows.Count));

                try
                {
                    if (!File.Exists(file.FullPath))
                    {
                        file.Status = FileStatus.Failed;
                        file.Message = "File not found";
                        log(LogLevel.Error, $"{file.FileName}: file not found.");
                        return;
                    }

                    var sheetsInFile = await Task.Run(() => SafeGetSheetNames(file.FullPath), ct);

                    var sheetsToRead = options.SheetNames
                        .Where(s => sheetsInFile.Contains(s, StringComparer.OrdinalIgnoreCase))
                        .ToList();

                    if (sheetsToRead.Count == 0)
                    {
                        file.Status = FileStatus.Skipped;
                        file.Message = "Target sheets not found";
                        log(LogLevel.Warning,
                            $"{file.FileName}: neither '{options.SheetNames[0]}' nor '{options.SheetNames[1]}' found — skipping.");
                        Interlocked.Increment(ref filesDoneCount);
                        return;
                    }

                    var fileRowsRead = 0;

                    foreach (var sheet in sheetsToRead)
                    {
                        ct.ThrowIfCancellationRequested();

                        await Task.Run(() =>
                        {
                            var rows = MiniExcel.Query(file.FullPath, useHeaderRow: true, sheetName: sheet);

                            foreach (var raw in rows)
                            {
                                if (ct.IsCancellationRequested) break;
                                if (raw is not IDictionary<string, object> dict) continue;

                                Interlocked.Increment(ref totalRowsRead);
                                fileRowsRead++;

                                // Update column order with first-seen columns.
                                if (dict.Count > 0)
                                {
                                    lock (columnOrderLock)
                                    {
                                        foreach (var col in dict.Keys)
                                        {
                                            if (col is null) continue;
                                            if (columnSet.Add(col))
                                            {
                                                columnOrder.Add(col);
                                            }
                                        }
                                    }
                                }

                                // Find email column (case-insensitive).
                                var email = FindEmail(dict, options.EmailColumnName);
                                var normalized = email?.Trim();

                                if (string.IsNullOrEmpty(normalized) || !IsValidEmail(normalized))
                                {
                                    Interlocked.Increment(ref invalidRemoved);
                                    continue;
                                }

                                var key = normalized.ToLowerInvariant();
                                if (!seenEmails.TryAdd(key, 0))
                                {
                                    Interlocked.Increment(ref duplicatesRemoved);
                                    continue;
                                }

                                // Snapshot row — MiniExcel reuses internal storage per sheet.
                                var snapshot = new Dictionary<string, object?>(dict.Count, StringComparer.Ordinal);
                                foreach (var kvp in dict)
                                {
                                    snapshot[kvp.Key] = kvp.Value;
                                }
                                acceptedRows.Add(snapshot);

                                if (fileRowsRead % 5000 == 0)
                                {
                                    file.RowsRead = fileRowsRead;
                                    progress.Report(BuildProgress(
                                        filesDoneCount, files.Count, file.FileName,
                                        totalRowsRead, duplicatesRemoved, invalidRemoved,
                                        acceptedRows.Count));
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
                    progress.Report(BuildProgress(
                        filesDoneCount, files.Count, file.FileName,
                        totalRowsRead, duplicatesRemoved, invalidRemoved,
                        acceptedRows.Count));
                }
            });

            cancellationToken.ThrowIfCancellationRequested();

            // Output column order: [name] -> [email] -> [everything else in first-seen order].
            List<string> finalColumns;
            lock (columnOrderLock)
            {
                finalColumns = new List<string>(columnOrder);
            }

            var nameCol = FindColumn(finalColumns, options.NameColumnName);
            var emailCol = FindColumn(finalColumns, options.EmailColumnName);

            if (nameCol is not null) finalColumns.Remove(nameCol);
            if (emailCol is not null) finalColumns.Remove(emailCol);

            var leading = new List<string>(2);
            if (nameCol is not null) leading.Add(nameCol);
            if (emailCol is not null) leading.Add(emailCol);
            finalColumns.InsertRange(0, leading);

            log(LogLevel.Info, $"Accepted {acceptedRows.Count:N0} unique valid records across {finalColumns.Count} columns. Writing output...");

            var outDir = string.IsNullOrWhiteSpace(options.OutputDirectory)
                ? Path.GetDirectoryName(files[0].FullPath) ?? Environment.CurrentDirectory
                : options.OutputDirectory!;
            Directory.CreateDirectory(outDir);

            var outputFile = Path.Combine(outDir, $"Processed_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            await Task.Run(() =>
            {
                // Normalize rows to the unified schema so MiniExcel writes consistent columns.
                IEnumerable<Dictionary<string, object?>> NormalizedRows()
                {
                    foreach (var row in acceptedRows)
                    {
                        var normalized = new Dictionary<string, object?>(finalColumns.Count, StringComparer.Ordinal);
                        foreach (var col in finalColumns)
                        {
                            normalized[col] = row.TryGetValue(col, out var v) ? v : null;
                        }
                        yield return normalized;
                    }
                }

                MiniExcel.SaveAs(
                    outputFile,
                    NormalizedRows(),
                    sheetName: options.OutputSheetName,
                    overwriteFile: true);
            }, cancellationToken);

            sw.Stop();

            log(LogLevel.Success,
                $"Wrote {acceptedRows.Count:N0} rows to {Path.GetFileName(outputFile)} in {sw.Elapsed.TotalSeconds:0.0}s.");

            return new ProcessingResult
            {
                Success = true,
                OutputPath = outputFile,
                TotalRowsRead = Interlocked.Read(ref totalRowsRead),
                DuplicatesRemoved = Interlocked.Read(ref duplicatesRemoved),
                InvalidRemoved = Interlocked.Read(ref invalidRemoved),
                ValidRecords = acceptedRows.Count,
                Elapsed = sw.Elapsed
            };
        }
        catch (OperationCanceledException)
        {
            sw.Stop();
            log(LogLevel.Warning, "Processing cancelled.");
            return new ProcessingResult
            {
                Success = false,
                ErrorMessage = "Cancelled",
                Elapsed = sw.Elapsed
            };
        }
        catch (Exception ex)
        {
            sw.Stop();
            log(LogLevel.Error, $"Fatal: {ex.Message}");
            return new ProcessingResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                Elapsed = sw.Elapsed
            };
        }
    }

    private static ProgressUpdate BuildProgress(
        int filesDone, int fileTotal, string? currentFile,
        long totalRowsRead, long duplicatesRemoved, long invalidRemoved,
        long validRecords)
    {
        var pct = fileTotal == 0 ? 0.0 : Math.Min(100.0, (double)filesDone / fileTotal * 100.0);
        return new ProgressUpdate(
            filesDone,
            fileTotal,
            currentFile,
            totalRowsRead,
            duplicatesRemoved,
            invalidRemoved,
            validRecords,
            pct);
    }

    private static string? FindColumn(IReadOnlyList<string> columns, string preferredName)
    {
        // Exact (case-insensitive) match first, else any column containing the preferred name.
        var exact = columns.FirstOrDefault(c => string.Equals(c, preferredName, StringComparison.OrdinalIgnoreCase));
        if (exact is not null) return exact;
        return columns.FirstOrDefault(c =>
            c is not null && c.IndexOf(preferredName, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static string? FindEmail(IDictionary<string, object> dict, string emailColumn)
    {
        // Exact, case-insensitive match first.
        foreach (var kvp in dict)
        {
            if (string.Equals(kvp.Key, emailColumn, StringComparison.OrdinalIgnoreCase))
            {
                return kvp.Value?.ToString();
            }
        }
        // Fallback: any key containing "email".
        foreach (var kvp in dict)
        {
            if (kvp.Key is not null &&
                kvp.Key.IndexOf("email", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return kvp.Value?.ToString();
            }
        }
        return null;
    }

    private static bool IsValidEmail(string email)
    {
        if (email.Length < 5 || email.Length > 320) return false;
        return EmailRegex.IsMatch(email);
    }

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
}
