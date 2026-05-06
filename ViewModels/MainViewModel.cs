using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using DataRefineX.Models;
using DataRefineX.Services;
using AppConfig = DataRefineX.AppConfig;

namespace DataRefineX.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly ExcelProcessor _processor = new();
    private readonly GoogleSheetsService _googleSheets = new();
    private readonly UpdateService _updateService;
    private readonly Dispatcher _dispatcher;
    private readonly DispatcherTimer _elapsedTimer;
    private DateTime _startedAt;

    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _updateCts;

    public MainViewModel()
    {
        _dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
        _updateService = new UpdateService(AppConfig.UpdateManifestUrl, AppConfig.UpdateHttpTimeout);

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

        AddGoogleSheetCommand = new RelayCommand(
            async _ => await AddGoogleSheetAsync(),
            _ => !IsProcessing && !IsFetchingGoogleSheet && !string.IsNullOrWhiteSpace(GoogleSheetUrl));

        UpdateNowCommand = new RelayCommand(
            async _ => await DownloadAndApplyUpdateAsync(),
            _ => UpdateInfo is not null && !IsDownloadingUpdate);
        DismissUpdateCommand = new RelayCommand(_ =>
        {
            UpdateDismissed = true;
        }, _ => UpdateInfo is not null && !IsDownloadingUpdate);

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

        // Fire-and-forget update check. Failures are silent.
        _ = CheckForUpdatesAsync();
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
    public RelayCommand UpdateNowCommand { get; }
    public RelayCommand DismissUpdateCommand { get; }
    public RelayCommand AddGoogleSheetCommand { get; }

    private string _googleSheetUrl = "";
    public string GoogleSheetUrl
    {
        get => _googleSheetUrl;
        set
        {
            if (_googleSheetUrl == value) return;
            _googleSheetUrl = value;
            OnPropertyChanged();
            AddGoogleSheetCommand.RaiseCanExecuteChanged();
        }
    }

    private bool _isFetchingGoogleSheet;
    public bool IsFetchingGoogleSheet
    {
        get => _isFetchingGoogleSheet;
        private set
        {
            if (_isFetchingGoogleSheet == value) return;
            _isFetchingGoogleSheet = value;
            OnPropertyChanged();
            AddGoogleSheetCommand.RaiseCanExecuteChanged();
        }
    }

    private async Task AddGoogleSheetAsync()
    {
        var url = GoogleSheetUrl?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(url)) return;

        if (!GoogleSheetsService.TryExtractId(url, out _))
        {
            AddLog(LogLevel.Warning, "Not a recognized Google Sheets URL. Expected: https://docs.google.com/spreadsheets/d/<ID>/edit");
            return;
        }

        IsFetchingGoogleSheet = true;
        AddLog(LogLevel.Info, "Downloading Google Sheet (must be shared as 'Anyone with the link')...");

        try
        {
            var result = await _googleSheets.DownloadAsync(url);
            if (!result.Success || result.LocalPath is null)
            {
                AddLog(LogLevel.Error, result.ErrorMessage ?? "Download failed.");
                return;
            }

            AddFiles(new[] { result.LocalPath });
            AddLog(LogLevel.Success, $"Added '{result.SuggestedName ?? Path.GetFileName(result.LocalPath)}' from Google Sheets.");
            GoogleSheetUrl = "";
        }
        catch (Exception ex)
        {
            AddLog(LogLevel.Error, $"Google Sheets fetch failed: {ex.Message}");
        }
        finally
        {
            IsFetchingGoogleSheet = false;
        }
    }

    private UpdateInfo? _updateInfo;
    public UpdateInfo? UpdateInfo
    {
        get => _updateInfo;
        private set
        {
            if (_updateInfo == value) return;
            _updateInfo = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsUpdateBannerVisible));
            OnPropertyChanged(nameof(UpdateVersionDisplay));
            UpdateNowCommand.RaiseCanExecuteChanged();
            DismissUpdateCommand.RaiseCanExecuteChanged();
        }
    }

    private bool _updateDismissed;
    public bool UpdateDismissed
    {
        get => _updateDismissed;
        private set
        {
            if (_updateDismissed == value) return;
            _updateDismissed = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsUpdateBannerVisible));
        }
    }

    private bool _isDownloadingUpdate;
    public bool IsDownloadingUpdate
    {
        get => _isDownloadingUpdate;
        private set
        {
            if (_isDownloadingUpdate == value) return;
            _isDownloadingUpdate = value;
            OnPropertyChanged();
            UpdateNowCommand.RaiseCanExecuteChanged();
            DismissUpdateCommand.RaiseCanExecuteChanged();
        }
    }

    private double _updateProgress;
    public double UpdateProgress
    {
        get => _updateProgress;
        private set
        {
            if (Math.Abs(_updateProgress - value) < 0.01) return;
            _updateProgress = value;
            OnPropertyChanged();
        }
    }

    public bool IsUpdateBannerVisible => _updateInfo is not null && !_updateDismissed;
    public string UpdateVersionDisplay => _updateInfo is null
        ? string.Empty
        : $"v{_updateInfo.Version}";

    public bool HasFiles => Files.Count > 0;
    public bool QueueEmptyHint => Files.Count == 0;

    public string AppVersionDisplay
    {
        get
        {
            // Read attributes directly off the assembly — works in single-file published apps,
            // unlike FileVersionInfo which depends on Assembly.Location (empty when bundled).
            var asm = typeof(MainViewModel).Assembly;
            var v = System.Reflection.CustomAttributeExtensions
                        .GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>(asm)?.InformationalVersion
                 ?? System.Reflection.CustomAttributeExtensions
                        .GetCustomAttribute<System.Reflection.AssemblyFileVersionAttribute>(asm)?.Version
                 ?? asm.GetName().Version?.ToString()
                 ?? "1.0";
            // Strip SourceLink build metadata like '+abc1234'.
            var plus = v.IndexOf('+');
            if (plus >= 0) v = v.Substring(0, plus);
            // Trim a trailing '.0' so 1.1.0.0 → 1.1.0 and 1.1.0 → 1.1.
            while (v.EndsWith(".0") && v.Count(c => c == '.') > 1) v = v.Substring(0, v.Length - 2);
            return $"v{v}";
        }
    }

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

    // ---------- Sheet selection ----------

    private SheetSelectionMode _sheetMode = SheetSelectionMode.All;
    public SheetSelectionMode SheetMode
    {
        get => _sheetMode;
        set
        {
            if (_sheetMode == value) return;
            _sheetMode = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsByNameMode));
            OnPropertyChanged(nameof(IsFirstNMode));
            _ = RefreshDetectedColumnsAsync();
        }
    }

    public bool IsByNameMode => _sheetMode == SheetSelectionMode.ByName;
    public bool IsFirstNMode => _sheetMode == SheetSelectionMode.FirstN;

    private string _sheetNamesText = "";
    public string SheetNamesText
    {
        get => _sheetNamesText;
        set { if (_sheetNamesText != value) { _sheetNamesText = value; OnPropertyChanged(); } }
    }

    private int _firstNSheets = 1;
    public int FirstNSheets
    {
        get => _firstNSheets;
        set
        {
            var clamped = Math.Max(1, value);
            if (_firstNSheets == clamped) return;
            _firstNSheets = clamped;
            OnPropertyChanged();
        }
    }

    // ---------- Dedup ----------

    private DedupKeyMode _dedupMode = DedupKeyMode.SingleColumn;
    public DedupKeyMode DedupMode
    {
        get => _dedupMode;
        set
        {
            if (_dedupMode == value) return;
            _dedupMode = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsSingleColumnDedup));
            OnPropertyChanged(nameof(IsMultiColumnDedup));
        }
    }

    public bool IsSingleColumnDedup => _dedupMode == DedupKeyMode.SingleColumn;
    public bool IsMultiColumnDedup => _dedupMode == DedupKeyMode.MultipleColumns;

    public ObservableCollection<string> DetectedColumns { get; } = new();

    private string? _dedupColumn;
    public string? DedupColumn
    {
        get => _dedupColumn;
        set { if (_dedupColumn != value) { _dedupColumn = value; OnPropertyChanged(); } }
    }

    private string _dedupColumnsText = "";
    public string DedupColumnsText
    {
        get => _dedupColumnsText;
        set { if (_dedupColumnsText != value) { _dedupColumnsText = value; OnPropertyChanged(); } }
    }

    private bool _caseSensitive;
    public bool CaseSensitive
    {
        get => _caseSensitive;
        set { if (_caseSensitive != value) { _caseSensitive = value; OnPropertyChanged(); } }
    }

    private bool _splitMultiValueCells;
    public bool SplitMultiValueCells
    {
        get => _splitMultiValueCells;
        set { if (_splitMultiValueCells != value) { _splitMultiValueCells = value; OnPropertyChanged(); } }
    }

    private string _multiValueDelimiters = "; , |";
    public string MultiValueDelimiters
    {
        get => _multiValueDelimiters;
        set { if (_multiValueDelimiters != value) { _multiValueDelimiters = value; OnPropertyChanged(); } }
    }

    // ---------- Validation ----------

    private bool _validateEmail;
    public bool ValidateEmail
    {
        get => _validateEmail;
        set { if (_validateEmail != value) { _validateEmail = value; OnPropertyChanged(); OnPropertyChanged(nameof(Validation)); } }
    }

    private bool _validateNotEmpty;
    public bool ValidateNotEmpty
    {
        get => _validateNotEmpty;
        set { if (_validateNotEmpty != value) { _validateNotEmpty = value; OnPropertyChanged(); OnPropertyChanged(nameof(Validation)); } }
    }

    public ValidationMode Validation
    {
        get
        {
            var v = ValidationMode.None;
            if (ValidateEmail) v |= ValidationMode.Email;
            if (ValidateNotEmpty) v |= ValidationMode.NotEmpty;
            return v;
        }
    }

    private bool _preserveSourceSheets = true;
    public bool PreserveSourceSheets
    {
        get => _preserveSourceSheets;
        set { if (_preserveSourceSheets != value) { _preserveSourceSheets = value; OnPropertyChanged(); } }
    }

    // ---------- Output sheet toggles ----------

    private bool _writeUniqueSheet = true;
    public bool WriteUniqueSheet
    {
        get => _writeUniqueSheet;
        set { if (_writeUniqueSheet != value) { _writeUniqueSheet = value; OnPropertyChanged(); } }
    }

    private bool _writeDuplicatesSheet = true;
    public bool WriteDuplicatesSheet
    {
        get => _writeDuplicatesSheet;
        set { if (_writeDuplicatesSheet != value) { _writeDuplicatesSheet = value; OnPropertyChanged(); } }
    }

    private bool _writeInvalidSheet = true;
    public bool WriteInvalidSheet
    {
        get => _writeInvalidSheet;
        set { if (_writeInvalidSheet != value) { _writeInvalidSheet = value; OnPropertyChanged(); } }
    }

    // ---------- Output format ----------

    private OutputFormat _outputFormat = OutputFormat.Xlsx;
    public OutputFormat OutputFormat
    {
        get => _outputFormat;
        set { if (_outputFormat != value) { _outputFormat = value; OnPropertyChanged(); } }
    }

    private OutputDestination _outputDestination = OutputDestination.NewFile;
    public OutputDestination OutputDestination
    {
        get => _outputDestination;
        set
        {
            if (_outputDestination == value) return;
            _outputDestination = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsInPlace));
            OnPropertyChanged(nameof(IsNewFile));
        }
    }

    public bool IsInPlace => _outputDestination == OutputDestination.InPlace;
    public bool IsNewFile => _outputDestination == OutputDestination.NewFile;

    private async Task RefreshDetectedColumnsAsync()
    {
        if (Files.Count == 0)
        {
            DetectedColumns.Clear();
            return;
        }

        try
        {
            var snapshotFiles = Files.ToList();
            var scanOptions = new ProcessingOptions
            {
                SheetMode = SheetMode,
                SheetNames = ParseSheetNames(),
                FirstNSheets = FirstNSheets
            };
            var cols = await _processor.ScanHeadersAsync(snapshotFiles, scanOptions);

            await _dispatcher.InvokeAsync(() =>
            {
                DetectedColumns.Clear();
                foreach (var c in cols) DetectedColumns.Add(c);
                // Auto-fill ONLY when user hasn't picked anything yet — never overwrite a value the user typed.
                if (string.IsNullOrWhiteSpace(DedupColumn))
                {
                    DedupColumn = DetectedColumns.FirstOrDefault();
                }
            });
        }
        catch
        {
            // Header scanning is best-effort — failures shouldn't block the user.
        }
    }

    private string[] ParseSheetNames()
    {
        if (SheetMode == SheetSelectionMode.ByName)
        {
            return SheetNamesText.Split(new[] { ',', ';', '\n', '\r' },
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
        return Array.Empty<string>();
    }

    public void AddFiles(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            if (string.IsNullOrWhiteSpace(path)) continue;
            if (!File.Exists(path)) continue;

            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm" && ext != ".csv")
            {
                AddLog(LogLevel.Warning, $"Skipping '{Path.GetFileName(path)}' — only .xlsx, .xlsm, .csv supported.");
                continue;
            }

            if (Files.Any(f => string.Equals(f.FullPath, path, StringComparison.OrdinalIgnoreCase)))
                continue;

            Files.Add(new FileItem(path));
        }

        _ = RefreshDetectedColumnsAsync();
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

        // Resolve dedup columns and validate config BEFORE flipping into processing state.
        var dedupColumns = DedupMode switch
        {
            DedupKeyMode.SingleColumn => string.IsNullOrWhiteSpace(DedupColumn)
                ? (DetectedColumns.FirstOrDefault() is { } first ? new[] { first } : Array.Empty<string>())
                : new[] { DedupColumn! },
            DedupKeyMode.MultipleColumns => DedupColumnsText.Split(
                new[] { ',', ';', '\n', '\r' },
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
            _ => Array.Empty<string>()
        };

        if (DedupMode == DedupKeyMode.SingleColumn && dedupColumns.Length == 0)
        {
            AddLog(LogLevel.Warning, "Pick a unique column in the Read panel before processing.");
            return;
        }
        if (DedupMode == DedupKeyMode.MultipleColumns && dedupColumns.Length == 0)
        {
            AddLog(LogLevel.Warning, "Multi-column dedup needs at least one column listed.");
            return;
        }

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
            OutputDirectory = Path.GetDirectoryName(filesToProcess[0].FullPath),
            SheetMode = SheetMode,
            SheetNames = ParseSheetNames(),
            FirstNSheets = FirstNSheets,
            DedupMode = DedupMode,
            DedupColumns = dedupColumns,
            CaseSensitive = CaseSensitive,
            SplitMultiValueCells = SplitMultiValueCells,
            MultiValueDelimiters = string.Concat((MultiValueDelimiters ?? "").Where(c => !char.IsWhiteSpace(c))),
            Validation = Validation,
            ValidationColumn = dedupColumns.FirstOrDefault() ?? "",
            WriteUniqueSheet = WriteUniqueSheet,
            WriteDuplicatesSheet = WriteDuplicatesSheet,
            WriteInvalidSheet = WriteInvalidSheet,
            OutputFormat = OutputFormat,
            Destination = OutputDestination,
            PreserveSourceSheets = PreserveSourceSheets
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

    // ------------------ Update flow ------------------

    private async Task CheckForUpdatesAsync()
    {
        try
        {
            await Task.Delay(AppConfig.UpdateCheckDelay);
            var info = await _updateService.CheckForUpdateAsync();
            if (info is null) return;

            await _dispatcher.InvokeAsync(() =>
            {
                UpdateInfo = info;
                var notes = string.IsNullOrWhiteSpace(info.ReleaseNotes)
                    ? ""
                    : $" — {info.ReleaseNotes}";
                AddLog(LogLevel.Info, $"Update available: v{info.Version}{notes}");
            });
        }
        catch
        {
            // Silent — offline or other issue. User isn't blocked.
        }
    }

    private async Task DownloadAndApplyUpdateAsync()
    {
        if (UpdateInfo is null || IsDownloadingUpdate) return;
        var info = UpdateInfo;

        IsDownloadingUpdate = true;
        UpdateProgress = 0;
        _updateCts?.Dispose();
        _updateCts = new CancellationTokenSource();

        try
        {
            AddLog(LogLevel.Info, $"Downloading update v{info.Version}...");
            var progress = new Progress<double>(p => UpdateProgress = p);
            var downloadedPath = await _updateService.DownloadAsync(info, progress, _updateCts.Token);

            AddLog(LogLevel.Success, "Update downloaded. Restarting to apply...");
            // Slight delay so the user sees the final state.
            await Task.Delay(400);

            UpdateInstaller.ApplyAndRestart(downloadedPath);
            // Application.Current.Shutdown() is called inside ApplyAndRestart.
        }
        catch (OperationCanceledException)
        {
            AddLog(LogLevel.Warning, "Update cancelled.");
        }
        catch (Exception ex)
        {
            AddLog(LogLevel.Error, $"Update failed: {ex.Message}");
            IsDownloadingUpdate = false;
        }
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
