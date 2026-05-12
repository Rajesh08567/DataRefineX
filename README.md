# DataRefineX

Drag-and-drop Excel & CSV deduplication for Windows. Pick any column as the unique key, validate it, choose what to keep, and write a clean output file in seconds.

Built with WPF on .NET 8 and [MiniExcel](https://github.com/mini-software/MiniExcel) for fast, low-memory streaming.

## Features

- **Bulk drag & drop** — `.xlsx`, `.xlsm`, `.csv`. Drop multiple files at once.
- **Google Sheets** — paste a share link (Anyone with the link), DataRefineX downloads it as `.xlsx` and queues it.
- **Flexible sheet selection** — All sheets / By name (with `*` wildcards) / First N sheets.
- **Per-column dedup** — pick a single column (e.g. `email`), combine multiple columns, or match the whole row. Case-sensitive matching is optional.
- **Multi-value cell splitting** — turn `a@x.com; b@x.com; c@x.com` in one cell into three rows that share the rest of the columns.
- **Validation** — drop rows where the unique column is empty or isn't a valid email.
- **Three output buckets** — Unique, Duplicates, and Invalid. Toggle each independently.
- **Preserve original sheet structure** — output keeps the source sheet names and tab order, dedup'd within each sheet.
- **Output splitting** — split very large outputs four ways: don't split, split each sheet at N rows, merge all unique rows then split into sheets of N, or merge and split into separate files of N.
- **Output as XLSX or CSV** — multi-sheet workbook, or one CSV per bucket.
- **In-place mode** — overwrite each source file with cleaned data (per-file dedup, originals replaced).
- **Live progress + activity log** — rows read, duplicates removed, invalid removed, valid records, elapsed time.
- **Self-update** — checks for new versions on launch and downloads/applies them in one click.

## Quick start

1. Download the latest `DataRefineX-vX.Y.Z.exe` from [publish/](publish/) (self-contained — no .NET install needed).
2. Run it. Drag your `.xlsx`/`.csv` files onto the drop zone (or click to browse).
3. Configure the Read panel (which sheets, which column is unique, validation rules).
4. Configure the Write panel (where to save, which buckets to include, file format).
5. Click **Start processing**. When done, click **Open file** or **Show in folder**.

## Settings reference

### Read panel

| Setting | Options | Notes |
|---|---|---|
| Which sheets | `All sheets`, `By name`, `First N` | `By name` accepts `*` wildcards (e.g. `Data*`, `*2024`). |
| How to detect duplicates | `Single column`, `Multiple columns`, `Whole row`, `None` | `None` skips dedup entirely — useful for validation-only passes. |
| Unique column | dropdown | Auto-populated from detected headers. Type a name if it isn't listed. |
| Case-sensitive matching | checkbox | Off by default — `Foo` and `foo` are duplicates. Enable for hashes / IDs. |
| Split multi-value cells | checkbox + delimiters | Each character in the delimiter box is a separator. Spaces ignored. |
| Validate the unique column | `Email format`, `Not empty` | Failed rows go to the Invalid bucket. |

### Write panel

| Setting | Options | Notes |
|---|---|---|
| Destination | `New file`, `Same file (overwrite source)` | In-place is destructive — backup first. |
| Keep original sheet names | checkbox | On: each input sheet → its own output sheet. Off: merge into one `Unique` sheet. |
| What to include | `Unique`, `Duplicates removed`, `Invalid / empty`, `Header filter dropdowns` | First three become their own sheet (or CSV file). `Header filter dropdowns` is off by default — turn on to get Excel's filter arrows on each output sheet's header row. |
| File format | `XLSX`, `CSV` | Only when `Destination = New file`. In-place follows the source extension. |
| Split output | `Don't split`, `Per sheet`, `Merged sheets`, `Separate files` | Combine with **Rows per chunk** (default 10000). Splitting is disabled in in-place mode. |

## Output structure

**XLSX** — one workbook with up to three buckets as separate sheets:

- `Customers`, `Vendors`, ... — one per source sheet (when *Keep original sheet names* is on), each holding only that sheet's unique rows.
- `Duplicates` — every row that was dropped as a duplicate, with a `_Source` column showing which sheet it came from.
- `Invalid` — every row that failed validation, with `_Source` and `_Reason` columns.

**CSV** — one file per bucket: `Processed_<timestamp>_Unique.csv`, `..._Duplicates.csv`, `..._Invalid.csv`.

Column order in every output sheet matches the original header order from the source files. Sheet order in the output matches the order in which sheets were first encountered.

## Build from source

Requires the [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

```powershell
# Debug build
dotnet build DataRefineX.csproj

# Release single-file self-contained EXE (Windows x64)
dotnet publish DataRefineX.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The published EXE lands at `bin\Release\net8.0-windows\win-x64\publish\DataRefineX-v<Version>.exe` and is renamed automatically (see the `VersionedPublishName` target in [DataRefineX.csproj](DataRefineX.csproj)).

## Project layout

```
DataRefineX/
├── App.xaml(.cs)              WPF entry point
├── MainWindow.xaml(.cs)       UI — drag/drop, settings, queue, progress, log
├── AppConfig.cs               Update check URL & timeouts
├── Models/                    FileItem, LogEntry, ProcessingStats, ProgressUpdate, UpdateInfo
├── Services/
│   ├── ExcelProcessor.cs      Core dedup/validate/write engine
│   ├── GoogleSheetsService.cs Google Sheets URL → .xlsx download
│   ├── UpdateService.cs       Manifest check + download
│   └── UpdateInstaller.cs     Apply downloaded update + restart
├── ViewModels/
│   ├── MainViewModel.cs       MVVM glue, all bindings
│   └── RelayCommand.cs
├── Converters/                XAML value converters (bool ↔ visibility, enum match, etc.)
├── Themes/                    Colors.xaml, Styles.xaml
├── Assets/                    Icons + logos
└── tests/                     Standalone integration test (Program.cs)
```

## Tech

- **.NET 8** + **WPF** (`net8.0-windows`)
- **MiniExcel 1.34.2** — streaming xlsx/csv read/write
- Self-contained single-file publish with compression and partial trimming
- MVVM with `INotifyPropertyChanged` and `RelayCommand`
- Concurrent file processing with `Parallel.ForEachAsync`
- `ConcurrentDictionary` + `ConcurrentBag` per-sheet buckets, with an ordered key list for deterministic output sheet order

## License

Internal / unreleased.
