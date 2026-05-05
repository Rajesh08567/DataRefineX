using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace DataRefineX.Services;

public sealed class GoogleSheetFetchResult
{
    public bool Success { get; init; }
    public string? LocalPath { get; init; }
    public string? SuggestedName { get; init; }
    public string? ErrorMessage { get; init; }
}

/// <summary>
/// Downloads publicly-shared Google Sheets via the export URL.
/// Only works for sheets shared with "Anyone with the link" — private sheets return HTML and we surface a helpful error.
/// </summary>
public sealed class GoogleSheetsService
{
    private static readonly Regex IdFromPath = new(
        @"/spreadsheets/d/(?<id>[A-Za-z0-9_\-]{20,})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex IdFromQuery = new(
        @"[?&]id=(?<id>[A-Za-z0-9_\-]{20,})",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex GidRegex = new(
        @"[#&?]gid=(?<gid>\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly HttpClient _http;

    public GoogleSheetsService(HttpClient? http = null)
    {
        _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
    }

    public static bool TryExtractId(string url, out string id)
    {
        id = "";
        if (string.IsNullOrWhiteSpace(url)) return false;

        var m = IdFromPath.Match(url);
        if (!m.Success) m = IdFromQuery.Match(url);
        if (!m.Success) return false;

        id = m.Groups["id"].Value;
        return true;
    }

    public async Task<GoogleSheetFetchResult> DownloadAsync(string url, CancellationToken ct = default)
    {
        if (!TryExtractId(url, out var id))
        {
            return new GoogleSheetFetchResult
            {
                Success = false,
                ErrorMessage = "That doesn't look like a Google Sheets URL. Expected something like https://docs.google.com/spreadsheets/d/<ID>/edit"
            };
        }

        // Always export the entire workbook as .xlsx so MultiSheet handling works.
        var exportUrl = $"https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx";

        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, exportUrl);
            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);

            if (!resp.IsSuccessStatusCode)
            {
                return new GoogleSheetFetchResult
                {
                    Success = false,
                    ErrorMessage = $"Google rejected the request ({(int)resp.StatusCode}). The sheet may be private — share it as 'Anyone with the link'."
                };
            }

            var contentType = resp.Content.Headers.ContentType?.MediaType ?? "";
            // Public xlsx export returns a spreadsheet MIME; private/login-redirect returns text/html.
            if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                return new GoogleSheetFetchResult
                {
                    Success = false,
                    ErrorMessage = "The sheet isn't publicly accessible. In Google Sheets: Share → 'Anyone with the link' → Viewer, then try again."
                };
            }

            var tempDir = Path.Combine(Path.GetTempPath(), "DataRefineX", "GoogleSheets");
            Directory.CreateDirectory(tempDir);

            var fileName = TryGetFileName(resp) ?? $"GoogleSheet_{id}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                fileName += ".xlsx";

            var localPath = Path.Combine(tempDir, SanitizeFilename(fileName));

            await using var stream = await resp.Content.ReadAsStreamAsync(ct);
            await using var file = File.Create(localPath);
            await stream.CopyToAsync(file, ct);

            return new GoogleSheetFetchResult
            {
                Success = true,
                LocalPath = localPath,
                SuggestedName = fileName
            };
        }
        catch (TaskCanceledException)
        {
            return new GoogleSheetFetchResult { Success = false, ErrorMessage = "Download timed out. Check your connection and try again." };
        }
        catch (HttpRequestException ex)
        {
            return new GoogleSheetFetchResult { Success = false, ErrorMessage = $"Network error: {ex.Message}" };
        }
        catch (Exception ex)
        {
            return new GoogleSheetFetchResult { Success = false, ErrorMessage = $"Could not download: {ex.Message}" };
        }
    }

    private static string? TryGetFileName(HttpResponseMessage resp)
    {
        var disp = resp.Content.Headers.ContentDisposition;
        var raw = disp?.FileNameStar ?? disp?.FileName;
        if (string.IsNullOrWhiteSpace(raw)) return null;
        return raw.Trim('"');
    }

    private static string SanitizeFilename(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new System.Text.StringBuilder(name.Length);
        foreach (var c in name)
        {
            sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
        }
        return sb.ToString();
    }
}
