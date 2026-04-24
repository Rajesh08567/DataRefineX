using System.Text.Json.Serialization;

namespace DataRefineX.Models;

public sealed class UpdateInfo
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = string.Empty;

    [JsonPropertyName("url")]
    public string Url { get; set; } = string.Empty;

    [JsonPropertyName("releaseNotes")]
    public string? ReleaseNotes { get; set; }

    [JsonPropertyName("sha256")]
    public string? Sha256 { get; set; }
}
