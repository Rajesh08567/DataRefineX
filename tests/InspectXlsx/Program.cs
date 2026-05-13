// Check the May 8 file that may have been the successful upload.
using System.IO;
using System.IO.Compression;
using System.Xml;

var paths = new[]
{
    @"C:\Users\lapto\Downloads\Validated Data 1\Validated Data\Processed_20260508_145447.xlsx",
    @"C:\Users\lapto\Downloads\Validated Data 1\Validated Data\Processed_20260507_175126.xlsx",
};

foreach (var path in paths)
{
    if (!File.Exists(path)) { Console.WriteLine($"MISSING: {path}"); continue; }
    Console.WriteLine($"\n========== {Path.GetFileName(path)} ==========");
    Console.WriteLine($"Size: {new FileInfo(path).Length:N0} bytes  Mod: {File.GetLastWriteTime(path)}");

    using var zip = ZipFile.OpenRead(path);
    int bom = 0;
    foreach (var e in zip.Entries.Where(x => x.FullName.EndsWith(".xml") || x.FullName.EndsWith(".rels")))
    {
        using var s = e.Open();
        var b = new byte[3]; var n = s.Read(b, 0, 3);
        if (n == 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) bom++;
    }
    Console.WriteLine($"BOMs: {bom}");

    // Cell type distribution
    var sst = new List<string>();
    var sstE = zip.GetEntry("xl/sharedStrings.xml");
    if (sstE is not null)
    {
        using var s = sstE.Open();
        var d = new XmlDocument(); d.Load(s);
        var ns = new XmlNamespaceManager(d.NameTable);
        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        foreach (XmlNode si in d.SelectNodes("//x:si", ns)!)
            sst.Add(si.SelectSingleNode("x:t", ns)?.InnerText ?? "");
        Console.WriteLine($"sharedStrings entries: {sst.Count}");
    }
    else
    {
        Console.WriteLine("NO sharedStrings.xml");
    }

    foreach (var sheetE in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet") && e.FullName.EndsWith(".xml")).OrderBy(e => e.FullName))
    {
        using var ss = sheetE.Open();
        var sd = new XmlDocument(); sd.Load(ss);
        var sns = new XmlNamespaceManager(sd.NameTable);
        sns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var dim = sd.SelectSingleNode("//x:dimension", sns)?.Attributes?["ref"]?.Value;
        var byType = new Dictionary<string, int>();
        foreach (XmlNode c in sd.SelectNodes("//x:c", sns)!)
        {
            var t = c.Attributes?["t"]?.Value ?? "(none)";
            byType[t] = byType.TryGetValue(t, out var v) ? v + 1 : 1;
        }
        Console.WriteLine($"  {sheetE.FullName}: dim={dim}");
        foreach (var kv in byType.OrderByDescending(k => k.Value))
            Console.WriteLine($"    t=\"{kv.Key}\": {kv.Value}");

        // First 3 rows
        int rn = 0;
        foreach (XmlNode r in sd.SelectNodes("//x:row", sns)!)
        {
            if (rn++ >= 3) break;
            var cells = new List<string>();
            foreach (XmlNode c in r.SelectNodes("x:c", sns)!)
            {
                var cref = c.Attributes?["r"]?.Value;
                var t = c.Attributes?["t"]?.Value ?? "";
                var v = c.SelectSingleNode("x:v", sns)?.InnerText;
                var val = (t == "s" && int.TryParse(v, out var i) && i < sst.Count) ? sst[i] : v;
                cells.Add($"{cref}({t})=\"{val}\"");
            }
            Console.WriteLine($"    [{string.Join(", ", cells)}]");
        }
    }
}
