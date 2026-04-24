using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using DataRefineX.Models;
using DataRefineX.Services;

namespace DataRefineX.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly ExcelProcessor _processor = new();
    private readonly Dispatcher _dispatcher;
    private readonly DispatcherTimer _elapsedTimer;
    private DateTime _startedAt;

    private CancellationTokenSource? _cts;

    public MainViewModel()
    {
        _dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;

        Files = new ObservableCollection<FileItem>();
        Logs = new ObservableCollection<LogEntry>();
        Stats = new ProcessingStats();

        StartCommand = new RelayCommand(
            async _ => await StartAsync(),
            _ => !IsProcessing && Files.Any(f => f.Status == FileStatus.Queued));
        CancelCommand = new RelayCommand(_ => Cancel(), _ => IsProcessing);
        ClearCommand = new RelayCommand(_ => ClearAll(), _ => !IsProcessing && (Files.Count > 0 || Logs.Count > 0));
        BrowseCommand = new RelayCommand(_ => { /* handled in view */ });
        OpenOutputCommand = new RelayCommand(_ => OpenOutput(), _ => !string.IsNullOrEmpty(Stats.OutputPath));
        OpenOutputFolderCommand = new RelayCommand(_ => OpenOutputFolder(), _ => !string.IsNullOrEmpty(Stats.OutputPath));
        RemoveFileCommand = new RelayCommand(param =>
        {
            if (param is FileItem f && !IsProcessing) Files.Remove(f);
        }, param => param is FileItem && !IsProcessing);

        Files.CollectionChanged += (_, __) =>
        {
            StartCommand.RaiseCanExecuteChanged();
            ClearCommand.RaiseCanExecuteChanged();
            OnPropertyChanged(nameof(HasFiles));
            OnPropertyChanged(nameof(QueueEmptyHint));
        };

        _elapsedTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _elapsedTimer.Tick += (_, _) =>
        {
            if (IsProcessing)
                Stats.Elapsed = DateTime.Now - _startedAt;
        };

        AddLog(LogLevel.Info, "Ready. Drag Excel files into the window or click Browse to begin.");
    }

    public ObservableCollection<FileItem> Files { get; }
    public ObservableCollection<LogEntry> Logs { get; }
    public ProcessingStats Stats { get; }

    public RelayCommand StartCommand { get; }
    public RelayCommand CancelCommand { get; }
    public RelayCommand ClearCommand { get; }
    public RelayCommand BrowseCommand { get; }
    public RelayCommand OpenOutputCommand { get; }
    public RelayCommand OpenOutputFolderCommand { get; }
    public RelayCommand RemoveFileCommand { get; }

    public bool HasFiles => Files.Count > 0;
    public bool QueueEmptyHint => Files.Count == 0;

    private bool _isProcessing;
    public bool IsProcessing
    {
        get => _isProcessing;
        private set
        {
            if (_isProcessing == value) return;
            _isProcessing = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(NotProcessing));
            StartCommand.RaiseCanExecuteChanged();
            CancelCommand.RaiseCanExecuteChanged();
            ClearCommand.RaiseCanExecuteChanged();
        }
    }

    public bool NotProcessing => !_isProcessing;

    private double _overallProgress;
    public double OverallProgress
    {
        get => _overallProgress;
        private set { if (Math.Abs(_overallProgress - value) > 0.001) { _overallProgress = value; OnPropertyChanged(); } }
    }

    private string? _currentFile;
    public string? CurrentFile
    {
        get => _currentFile;
        private set { if (_currentFile != value) { _currentFile = value; OnPropertyChanged(); } }
    }

    private string _statusMessage = "Idle";
    public string StatusMessage
    {
        get => _statusMessage;
        private set { if (_statusMessage != value) { _statusMessage = value; OnPropertyChanged(); } }
    }

    private bool _isCompleted;
    public bool IsCompleted
    {
        get => _isCompleted;
        private set { if (_isCompleted != value) { _isCompleted = value; OnPropertyChanged(); } }
    }

    public void AddFiles(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            if (string.IsNullOrWhiteSpace(path)) continue;
            if (!File.Exists(path)) continue;

            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm")
            {
                AddLog(LogLevel.Warning, $"Skipping '{Path.GetFileName(path)}' — only .xlsx/.xlsm supported.");
                continue;
            }

            if (Files.Any(f => string.Equals(f.FullPath, path, StringComparison.OrdinalIgnoreCase)))
                continue;

            Files.Add(new FileItem(path));
        }
    }

    private void ClearAll()
    {
        Files.Clear();
        Logs.Clear();
        Stats.Reset();
        OverallProgress = 0;
        CurrentFile = null;
        StatusMessage = "Idle";
        IsCompleted = false;
        AddLog(LogLevel.Info, "Cleared.");
    }

    private async Task StartAsync()
    {
        if (IsProcessing) return;

        // Only process files that haven't been run yet. This way, once a run
        // finishes the button correctly disables, and it only re-enables when
        // the user queues new files (which start in Queued state).
        var filesToProcess = Files.Where(f => f.Status == FileStatus.Queued).ToList();
        if (filesToProcess.Count == 0) return;

        IsProcessing = true;
        IsCompleted = false;
        Stats.Reset();
        Stats.FilesTotal = filesToProcess.Count;
        OverallProgress = 0;
        CurrentFile = null;
        StatusMessage = "Starting...";

        _startedAt = DateTime.Now;
        _elapsedTimer.Start();

        foreach (var f in filesToProcess)
        {
            f.RowsRead = 0;
            f.Message = null;
        }

        _cts = new CancellationTokenSource();

        var progress = new Progress<ProgressUpdate>(OnProgress);

        var options = new ProcessingOptions
        {
            OutputDirectory = Path.GetDirectoryName(filesToProcess[0].FullPath)
        };

        try
        {
            var result = await _processor.ProcessAsync(
                filesToProcess,
                options,
                progress,
                AddLog,
                _cts.Token);

            if (result.Success)
            {
                Stats.OutputPath = result.OutputPath;
                Stats.TotalRowsRead = result.TotalRowsRead;
                Stats.DuplicatesRemoved = result.DuplicatesRemoved;
                Stats.InvalidRemoved = result.InvalidRemoved;
                Stats.ValidRecords = result.ValidRecords;
                Stats.Elapsed = result.Elapsed;
                OverallProgress = 100;
                StatusMessage = "Completed";
                IsCompleted = true;
                CurrentFile = Path.GetFileName(result.OutputPath);
                AddLog(LogLevel.Success, $"Completed in {result.Elapsed.TotalSeconds:0.0}s. Output: {result.OutputPath}");
            }
            else
            {
                StatusMessage = result.ErrorMessage ?? "Failed";
                AddLog(LogLevel.Error, $"Processing failed: {result.ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelled";
            AddLog(LogLevel.Warning, "Cancelled by user.");
        }
        catch (Exception ex)
        {
            StatusMessage = "Error";
            AddLog(LogLevel.Error, ex.Message);
        }
        finally
        {
            _elapsedTimer.Stop();
            IsProcessing = false;
            _cts?.Dispose();
            _cts = null;
            OpenOutputCommand.RaiseCanExecuteChanged();
            OpenOutputFolderCommand.RaiseCanExecuteChanged();
        }
    }

    private void Cancel()
    {
        _cts?.Cancel();
        StatusMessage = "Cancelling...";
    }

    private void OnProgress(ProgressUpdate p)
    {
        OverallProgress = p.OverallPercent;
        CurrentFile = p.CurrentFile;
        Stats.TotalRowsRead = p.TotalRowsRead;
        Stats.DuplicatesRemoved = p.DuplicatesRemoved;
        Stats.InvalidRemoved = p.InvalidRemoved;
        Stats.ValidRecords = p.ValidRecords;
        Stats.FilesProcessed = p.FileIndex;
        StatusMessage = p.FileIndex >= p.FileTotal
            ? "Writing output..."
            : $"Processing {p.FileIndex + 1} of {p.FileTotal}";
    }

    private void AddLog(LogLevel level, string message)
    {
        if (_dispatcher.CheckAccess())
        {
            AppendLog(level, message);
        }
        else
        {
            _dispatcher.BeginInvoke(new Action(() => AppendLog(level, message)));
        }
    }

    private void AppendLog(LogLevel level, string message)
    {
        Logs.Add(new LogEntry(level, message));
        while (Logs.Count > 2000)
        {
            Logs.RemoveAt(0);
        }
        ClearCommand.RaiseCanExecuteChanged();
    }

    private void OpenOutput()
    {
        if (string.IsNullOrEmpty(Stats.OutputPath) || !File.Exists(Stats.OutputPath)) return;
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Stats.OutputPath)
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AddLog(LogLevel.Error, $"Could not open file: {ex.Message}");
        }
    }

    private void OpenOutputFolder()
    {
        if (string.IsNullOrEmpty(Stats.OutputPath)) return;
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"/select,\"{Stats.OutputPath}\"")
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AddLog(LogLevel.Error, $"Could not open folder: {ex.Message}");
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
