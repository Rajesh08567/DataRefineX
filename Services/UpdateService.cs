using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using DataRefineX.Models;

namespace DataRefineX.Services;

public sealed class UpdateService : IDisposable
{
    private readonly HttpClient _http;
    private readonly string? _manifestUrl;

    public UpdateService(string? manifestUrl, TimeSpan httpTimeout)
    {
        _manifestUrl = manifestUrl;
        _http = new HttpClient { Timeout = httpTimeout };
        _http.DefaultRequestHeaders.Add("User-Agent",
            $"DataRefineX-Updater/{CurrentVersion}");
        _http.DefaultRequestHeaders.Add("Accept", "application/json, */*");
    }

    public static Version CurrentVersion =>
        Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0);

    public async Task<UpdateInfo?> CheckForUpdateAsync(CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(_manifestUrl)) return null;

        try
        {
            // Cache-busting query parameter so CDNs/GitHub don't serve a stale manifest.
            var bustedUrl = _manifestUrl + (_manifestUrl.Contains('?') ? "&" : "?") + "_=" + DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            var json = await _http.GetStringAsync(bustedUrl, ct);

            var info = JsonSerializer.Deserialize<UpdateInfo>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (info is null || string.IsNullOrWhiteSpace(info.Version) || string.IsNullOrWhiteSpace(info.Url))
                return null;

            if (!Version.TryParse(info.Version, out var latest)) return null;

            // Only return if strictly newer than the running version.
            return latest > CurrentVersion ? info : null;
        }
        catch
        {
            // Offline, manifest malformed, host unreachable — treat as "no update".
            return null;
        }
    }

    /// <summary>
    /// Downloads the update EXE to a temp file. Reports progress 0..100.
    /// Verifies SHA-256 if provided by the manifest. Returns the local path.
    /// </summary>
    public async Task<string> DownloadAsync(
        UpdateInfo info,
        IProgress<double>? progress,
        CancellationToken ct)
    {
        // Download into a per-session temp subfolder so we don't collide with prior attempts.
        var tempDir = Path.Combine(Path.GetTempPath(), "DataRefineX_update");
        Directory.CreateDirectory(tempDir);

        var downloadName = SafeFileNameFromUrl(info.Url) ?? $"DataRefineX-v{info.Version}.exe";
        var destPath = Path.Combine(tempDir, downloadName);

        // Clean prior partial download if any.
        if (File.Exists(destPath))
        {
            try { File.Delete(destPath); } catch { /* best effort */ }
        }

        using var response = await _http.GetAsync(info.Url, HttpCompletionOption.ResponseHeadersRead, ct);
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength ?? -1L;

        await using (var netStream = await response.Content.ReadAsStreamAsync(ct))
        await using (var fileStream = File.Create(destPath))
        {
            var buffer = new byte[81920];
            long totalRead = 0;
            int read;
            while ((read = await netStream.ReadAsync(buffer, ct)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, read), ct);
                totalRead += read;
                if (totalBytes > 0 && progress is not null)
                {
                    progress.Report((double)totalRead / totalBytes * 100.0);
                }
            }
        }

        // Verify SHA-256 if the manifest provided one.
        if (!string.IsNullOrWhiteSpace(info.Sha256))
        {
            var actual = await ComputeSha256Async(destPath, ct);
            if (!string.Equals(actual, info.Sha256.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                try { File.Delete(destPath); } catch { }
                throw new InvalidDataException(
                    $"Update integrity check failed (SHA-256 mismatch). Expected {info.Sha256}, got {actual}.");
            }
        }

        progress?.Report(100.0);
        return destPath;
    }

    private static string? SafeFileNameFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var name = Path.GetFileName(uri.LocalPath);
            if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                return null;
            if (!name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) return null;
            return name;
        }
        catch
        {
            return null;
        }
    }

    private static async Task<string> ComputeSha256Async(string path, CancellationToken ct)
    {
        await using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        var hash = await sha.ComputeHashAsync(stream, ct);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public void Dispose() => _http.Dispose();
}
