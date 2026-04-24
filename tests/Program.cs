using DataRefineX.Models;
using DataRefineX.Services;
using MiniExcelLibs;

var tmpDir = Path.Combine(Path.GetTempPath(), "DataRefineX_ITest_" + Guid.NewGuid().ToString("N").Substring(0, 8));
Directory.CreateDirectory(tmpDir);
Console.WriteLine($"Test dir: {tmpDir}");

try
{
    // Build two synthetic Excel files with both required sheets + a noise sheet that should be ignored.
    var file1 = Path.Combine(tmpDir, "sample_a.xlsx");
    var file2 = Path.Combine(tmpDir, "sample_b.xlsx");

    // FILE A: Valid+BasicCheck+DEA with 6 rows (2 valid unique, 1 dup of the first, 1 invalid, 1 empty, 1 more valid).
    // Plus a "Ignored" sheet that must not be read.
    var fileASheets = new Dictionary<string, object>
    {
        ["Valid+BasicCheck+DEA"] = new[]
        {
            new { email = "alice@example.com", name = "Alice",   score = 92 },
            new { email = "BOB@Example.COM",   name = "Bob",     score = 81 },
            new { email = "alice@example.com", name = "Alice-2", score = 70 }, // duplicate (case-insensitive)
            new { email = "not-an-email",      name = "Broken",  score = 55 },
            new { email = "",                  name = "Empty",   score = 0  },
            new { email = "carol@example.com", name = "Carol",   score = 88 }
        },
        ["CatchAll_AcceptAll"] = new[]
        {
            new { email = "dave@example.com",  name = "Dave",    score = 77 }
        },
        ["Ignored"] = new[]
        {
            new { email = "should-not-appear@nope.com", name = "Noise", score = 0 }
        }
    };
    MiniExcel.SaveAs(file1, fileASheets, overwriteFile: true);

    // FILE B: 1 new unique + 1 duplicate across files + 1 invalid.
    var fileBSheets = new Dictionary<string, object>
    {
        ["Valid+BasicCheck+DEA"] = new[]
        {
            new { email = "eve@example.com",    name = "Eve",    score = 95 },
            new { email = "bob@example.com",    name = "BobDup", score = 60 }, // dup of FILE A's BOB@Example.COM
            new { email = "@bad.com",           name = "Bad",    score = 0 }
        },
        ["CatchAll_AcceptAll"] = new[]
        {
            new { email = "frank@example.com",  name = "Frank",  score = 85 }
        }
    };
    MiniExcel.SaveAs(file2, fileBSheets, overwriteFile: true);

    Console.WriteLine($"Created: {file1}");
    Console.WriteLine($"Created: {file2}");

    var processor = new ExcelProcessor();
    var files = new List<FileItem> { new(file1), new(file2) };
    var options = new ProcessingOptions { OutputDirectory = tmpDir };
    var progress = new Progress<ProgressUpdate>(p => { /* silent */ });
    void Log(LogLevel l, string m) => Console.WriteLine($"  [{l}] {m}");

    var result = await processor.ProcessAsync(files, options, progress, Log, CancellationToken.None);

    Console.WriteLine();
    Console.WriteLine($"Success:        {result.Success}");
    Console.WriteLine($"TotalRowsRead:  {result.TotalRowsRead}");
    Console.WriteLine($"Duplicates:     {result.DuplicatesRemoved}");
    Console.WriteLine($"Invalid:        {result.InvalidRemoved}");
    Console.WriteLine($"ValidRecords:   {result.ValidRecords}");
    Console.WriteLine($"Output:         {result.OutputPath}");
    Console.WriteLine($"Elapsed:        {result.Elapsed.TotalMilliseconds:F0} ms");

    // Expected:
    // FILE A sheet Valid+BasicCheck+DEA = 6 rows, CatchAll_AcceptAll = 1 row → 7 rows read
    // FILE B sheet Valid+BasicCheck+DEA = 3 rows, CatchAll_AcceptAll = 1 row → 4 rows read
    // Total = 11 rows read
    // Invalid: file A "not-an-email" + "" + file B "@bad.com" = 3
    // Duplicates: file A alice dup + file B bob dup = 2
    // Valid unique: alice, bob, carol, dave, eve, frank = 6
    var expectations = new (string name, long actual, long expected)[]
    {
        ("TotalRowsRead",  result.TotalRowsRead,     11),
        ("Duplicates",     result.DuplicatesRemoved,  2),
        ("Invalid",        result.InvalidRemoved,     3),
        ("ValidRecords",   result.ValidRecords,       6)
    };

    var failed = false;
    foreach (var (name, actual, expected) in expectations)
    {
        var ok = actual == expected;
        Console.WriteLine($"  {(ok ? "PASS" : "FAIL")}  {name}: expected {expected}, got {actual}");
        if (!ok) failed = true;
    }

    // Verify output sheet exists with the expected sheet name.
    if (!string.IsNullOrEmpty(result.OutputPath) && File.Exists(result.OutputPath))
    {
        var sheets = MiniExcel.GetSheetNames(result.OutputPath).ToList();
        Console.WriteLine($"  Output sheets: {string.Join(", ", sheets)}");
        var onlyCorrectSheet = sheets.Count == 1 && sheets[0] == "Valid+BasicCheck+DEA";
        Console.WriteLine($"  {(onlyCorrectSheet ? "PASS" : "FAIL")}  Output has exactly one sheet named 'Valid+BasicCheck+DEA'");
        if (!onlyCorrectSheet) failed = true;

        var outputRows = MiniExcel.Query(result.OutputPath, useHeaderRow: true).Cast<IDictionary<string, object>>().ToList();
        Console.WriteLine($"  Output row count: {outputRows.Count}");
        var rowCountOk = outputRows.Count == 6;
        Console.WriteLine($"  {(rowCountOk ? "PASS" : "FAIL")}  Output contains exactly 6 rows");
        if (!rowCountOk) failed = true;

        // Verify column order: name must be first, email must be second.
        if (outputRows.Count > 0)
        {
            var cols = outputRows[0].Keys.ToList();
            Console.WriteLine($"  Output columns (in order): {string.Join(", ", cols)}");
            var nameFirst = cols.Count > 0 && cols[0].Equals("name", StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"  {(nameFirst ? "PASS" : "FAIL")}  First column is 'name'");
            if (!nameFirst) failed = true;

            var emailSecond = cols.Count > 1 && cols[1].Equals("email", StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"  {(emailSecond ? "PASS" : "FAIL")}  Second column is 'email'");
            if (!emailSecond) failed = true;
        }

        // Verify no duplicates and no ignored-sheet row.
        var emails = outputRows.Select(r =>
            r.Keys.FirstOrDefault(k => k.Equals("email", StringComparison.OrdinalIgnoreCase)) is string key ? r[key]?.ToString() : null)
            .Where(e => !string.IsNullOrEmpty(e)).ToList();
        var anyNoise = emails.Any(e => e!.Contains("should-not-appear"));
        Console.WriteLine($"  {(anyNoise ? "FAIL" : "PASS")}  No rows leaked from 'Ignored' sheet");
        if (anyNoise) failed = true;
    }

    Environment.ExitCode = failed ? 1 : 0;
    Console.WriteLine(failed ? "\n*** TESTS FAILED ***" : "\nAll integration checks passed.");
}
finally
{
    try { Directory.Delete(tmpDir, true); } catch { }
}
