using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;

namespace DataRefineX.Services;

/// <summary>
/// Applies a downloaded update by spawning a tiny batch script that waits for
/// the current process to exit, moves the new EXE next to the current one,
/// deletes the old EXE, then launches the new EXE. The app shuts down
/// immediately after spawning the helper.
/// </summary>
public static class UpdateInstaller
{
    public static void ApplyAndRestart(string downloadedExePath)
    {
        if (!File.Exists(downloadedExePath))
            throw new FileNotFoundException("Downloaded update file is missing.", downloadedExePath);

        var currentExePath = GetCurrentExecutablePath();
        var targetDir = Path.GetDirectoryName(currentExePath)
            ?? throw new InvalidOperationException("Cannot determine current EXE directory.");
        var targetPath = Path.Combine(targetDir, Path.GetFileName(downloadedExePath));

        var batchPath = WriteHelperBatch(
            srcPath: downloadedExePath,
            targetPath: targetPath,
            oldPath: currentExePath);

        // Spawn the helper detached so it survives our exit.
        Process.Start(new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c \"\"{batchPath}\"\"",
            CreateNoWindow = true,
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetTempPath()
        });

        // Exit cleanly so the helper can replace the EXE.
        Application.Current.Shutdown();
    }

    private static string GetCurrentExecutablePath()
    {
        // .NET 6+ single-file-friendly way to get the host EXE path.
        var path = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(path) && File.Exists(path)) return path;

        // Fallback: scan the app base directory for any DataRefineX*.exe.
        var baseDir = AppContext.BaseDirectory;
        if (!string.IsNullOrEmpty(baseDir) && Directory.Exists(baseDir))
        {
            var match = Directory.EnumerateFiles(baseDir, "DataRefineX*.exe").FirstOrDefault();
            if (!string.IsNullOrEmpty(match)) return match;
        }

        throw new InvalidOperationException("Cannot locate the running executable.");
    }

    private static string WriteHelperBatch(string srcPath, string targetPath, string oldPath)
    {
        // The batch:
        //  1) waits long enough for the current process to fully exit
        //  2) moves the downloaded EXE into the target folder (overwrites if same name)
        //  3) deletes the old EXE if its path differs from the new one
        //  4) relaunches the new EXE
        //  5) deletes itself
        // NOTE: srcPath may equal targetPath if already placed correctly; handle gracefully.
        var batch = $@"@echo off
setlocal
set ""SRC={srcPath}""
set ""DST={targetPath}""
set ""OLD={oldPath}""

REM Wait up to ~10s for the current process to release file locks.
for /L %%i in (1,1,10) do (
    >nul 2>&1 (ren ""%OLD%"" ""%~nxOLD%.lock"" && ren ""%OLD%.lock"" ""%~nxOLD%"")
    if not errorlevel 1 goto :ready
    timeout /t 1 /nobreak > NUL
)
:ready

REM Move the downloaded EXE into place (no-op if SRC==DST).
if /i not ""%SRC%""==""%DST%"" (
    move /Y ""%SRC%"" ""%DST%"" > NUL 2>&1
)

REM Delete the old EXE if it isn't the same path as the new one.
if /i not ""%OLD%""==""%DST%"" (
    if exist ""%OLD%"" del /Q ""%OLD%"" > NUL 2>&1
)

REM Launch the new version.
start """" ""%DST%""

REM Self-delete.
del /Q ""%~f0"" > NUL 2>&1
endlocal
";

        var batchPath = Path.Combine(Path.GetTempPath(), $"drx_update_{Guid.NewGuid():N}.bat");
        File.WriteAllText(batchPath, batch);
        return batchPath;
    }
}
