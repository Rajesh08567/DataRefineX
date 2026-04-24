namespace DataRefineX;

internal static class AppConfig
{
    // =====================================================================
    //  UPDATE MANIFEST URL
    //
    //  Host a JSON file at this URL with the following shape:
    //
    //    {
    //      "version":      "1.2.0",
    //      "url":          "https://.../DataRefineX-v1.2.0.exe",
    //      "releaseNotes": "Optional — shown in the update banner",
    //      "sha256":       "Optional — hex checksum; if present it's verified"
    //    }
    //
    //  When the running app's version is lower than "version", the user sees
    //  an "Update available" banner. Clicking Update downloads the file at
    //  "url", verifies sha256 if provided, then replaces the running EXE.
    //
    //  Simplest hosting: put latest.json in a GitHub repo and link the raw
    //  URL here. Upload each DataRefineX-vX.Y.Z.exe as a GitHub Release
    //  asset and paste its download URL into "url".
    //
    //  Set to null (or empty) to disable auto-update entirely.
    // =====================================================================
    public const string? UpdateManifestUrl =
        "https://raw.githubusercontent.com/YOUR_USER/DataRefineX/main/latest.json";

    // Delay before the first update check (gives the UI time to settle).
    public static readonly TimeSpan UpdateCheckDelay = TimeSpan.FromSeconds(3);

    // HTTP timeout for the manifest fetch and the download.
    public static readonly TimeSpan UpdateHttpTimeout = TimeSpan.FromSeconds(30);
}
