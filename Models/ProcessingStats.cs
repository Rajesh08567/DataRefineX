using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DataRefineX.Models;

public sealed class ProcessingStats : INotifyPropertyChanged
{
    private long _totalRowsRead;
    public long TotalRowsRead
    {
        get => _totalRowsRead;
        set { if (_totalRowsRead != value) { _totalRowsRead = value; OnPropertyChanged(); } }
    }

    private long _duplicatesRemoved;
    public long DuplicatesRemoved
    {
        get => _duplicatesRemoved;
        set { if (_duplicatesRemoved != value) { _duplicatesRemoved = value; OnPropertyChanged(); } }
    }

    private long _invalidRemoved;
    public long InvalidRemoved
    {
        get => _invalidRemoved;
        set { if (_invalidRemoved != value) { _invalidRemoved = value; OnPropertyChanged(); } }
    }

    private long _validRecords;
    public long ValidRecords
    {
        get => _validRecords;
        set { if (_validRecords != value) { _validRecords = value; OnPropertyChanged(); } }
    }

    private int _filesProcessed;
    public int FilesProcessed
    {
        get => _filesProcessed;
        set { if (_filesProcessed != value) { _filesProcessed = value; OnPropertyChanged(); } }
    }

    private int _filesTotal;
    public int FilesTotal
    {
        get => _filesTotal;
        set { if (_filesTotal != value) { _filesTotal = value; OnPropertyChanged(); } }
    }

    private string? _outputPath;
    public string? OutputPath
    {
        get => _outputPath;
        set { if (_outputPath != value) { _outputPath = value; OnPropertyChanged(); } }
    }

    private TimeSpan _elapsed;
    public TimeSpan Elapsed
    {
        get => _elapsed;
        set { if (_elapsed != value) { _elapsed = value; OnPropertyChanged(); OnPropertyChanged(nameof(ElapsedDisplay)); } }
    }

    public string ElapsedDisplay => _elapsed.TotalMilliseconds < 1
        ? "0s"
        : _elapsed.TotalSeconds < 60
            ? $"{_elapsed.TotalSeconds:0.0}s"
            : $"{(int)_elapsed.TotalMinutes}m {_elapsed.Seconds}s";

    public void Reset()
    {
        TotalRowsRead = 0;
        DuplicatesRemoved = 0;
        InvalidRemoved = 0;
        ValidRecords = 0;
        FilesProcessed = 0;
        FilesTotal = 0;
        OutputPath = null;
        Elapsed = TimeSpan.Zero;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
