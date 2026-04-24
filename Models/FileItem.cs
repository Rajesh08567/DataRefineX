using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace DataRefineX.Models;

public enum FileStatus
{
    Queued,
    Reading,
    Processed,
    Failed,
    Skipped
}

public sealed class FileItem : INotifyPropertyChanged
{
    public FileItem(string fullPath)
    {
        FullPath = fullPath;
        FileName = Path.GetFileName(fullPath);
        try
        {
            var info = new FileInfo(fullPath);
            SizeBytes = info.Length;
        }
        catch
        {
            SizeBytes = 0;
        }
    }

    public string FullPath { get; }
    public string FileName { get; }
    public long SizeBytes { get; }

    public string SizeDisplay => FormatSize(SizeBytes);

    private FileStatus _status = FileStatus.Queued;
    public FileStatus Status
    {
        get => _status;
        set
        {
            if (_status == value) return;
            _status = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(IsActive));
        }
    }

    public string StatusText => _status switch
    {
        FileStatus.Queued => "Queued",
        FileStatus.Reading => "Reading",
        FileStatus.Processed => "Done",
        FileStatus.Failed => "Failed",
        FileStatus.Skipped => "Skipped",
        _ => "Unknown"
    };

    public bool IsActive => _status == FileStatus.Reading;

    private int _rowsRead;
    public int RowsRead
    {
        get => _rowsRead;
        set { if (_rowsRead != value) { _rowsRead = value; OnPropertyChanged(); OnPropertyChanged(nameof(RowsDisplay)); } }
    }

    public string RowsDisplay => _rowsRead > 0 ? $"{_rowsRead:N0} rows" : string.Empty;

    private string? _message;
    public string? Message
    {
        get => _message;
        set { if (_message != value) { _message = value; OnPropertyChanged(); } }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    private static string FormatSize(long bytes)
    {
        if (bytes <= 0) return "—";
        string[] suffix = { "B", "KB", "MB", "GB", "TB" };
        double value = bytes;
        int i = 0;
        while (value >= 1024 && i < suffix.Length - 1) { value /= 1024; i++; }
        return $"{value:0.#} {suffix[i]}";
    }
}
