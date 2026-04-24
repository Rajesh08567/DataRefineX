using System.Collections.Specialized;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;
using DataRefineX.Models;
using DataRefineX.ViewModels;

namespace DataRefineX;

public partial class MainWindow : Window
{
    private MainViewModel Vm => (MainViewModel)DataContext;

    public MainWindow()
    {
        InitializeComponent();

        Loaded += (_, _) =>
        {
            if (Vm.Logs is INotifyCollectionChanged ncc)
            {
                ncc.CollectionChanged += (_, _) => Dispatcher.BeginInvoke(new Action(ScrollLogToEnd));
            }
        };

        DragEnter += Window_DragEnter;
        DragOver += Window_DragEnter;
        Drop += Window_Drop;

        PreviewDragOver += (_, e) => { e.Effects = HasFiles(e) ? DragDropEffects.Copy : DragDropEffects.None; e.Handled = true; };

        SourceInitialized += OnSourceInitialized;
    }

    // ----- Proper work-area clamping for a borderless Window -----

    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        // Clamp the initial size to the current monitor's work area so the
        // title bar (our custom header + window controls) can never start
        // off-screen on smaller displays (1366x768, 1280x720, etc).
        var workArea = SystemParameters.WorkArea;
        if (Width > workArea.Width)  { Width  = workArea.Width;  }
        if (Height > workArea.Height) { Height = workArea.Height; }

        // Re-center on the work area after clamping.
        Left = workArea.Left + (workArea.Width  - Width)  / 2;
        Top  = workArea.Top  + (workArea.Height - Height) / 2;

        // Hook WM_GETMINMAXINFO so Maximize respects the work area instead
        // of extending 8px past each edge (WPF + WindowChrome quirk).
        var hwnd = new WindowInteropHelper(this).Handle;
        HwndSource.FromHwnd(hwnd)?.AddHook(WndProc);
    }

    private const int WM_GETMINMAXINFO = 0x0024;
    private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != WM_GETMINMAXINFO) return IntPtr.Zero;

        var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
        if (monitor == IntPtr.Zero) return IntPtr.Zero;

        var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
        if (!GetMonitorInfo(monitor, ref info)) return IntPtr.Zero;

        var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
        var work = info.rcWork;
        var mon  = info.rcMonitor;

        mmi.ptMaxPosition.X = Math.Abs(work.Left - mon.Left);
        mmi.ptMaxPosition.Y = Math.Abs(work.Top  - mon.Top);
        mmi.ptMaxSize.X     = Math.Abs(work.Right - work.Left);
        mmi.ptMaxSize.Y     = Math.Abs(work.Bottom - work.Top);

        Marshal.StructureToPtr(mmi, lParam, true);
        handled = true;
        return IntPtr.Zero;
    }

    // ----- Title bar buttons -----

    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState.Minimized;

    private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        => WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    // ----- Drag & drop -----

    private static bool HasFiles(DragEventArgs e) => e.Data.GetDataPresent(DataFormats.FileDrop);

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        e.Effects = HasFiles(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (!HasFiles(e)) return;
        var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
        Vm.AddFiles(paths);
        e.Handled = true;
    }

    private void DropZone_DragEnter(object sender, DragEventArgs e)
    {
        if (sender is Border b && HasFiles(e))
        {
            b.Background = (Brush)FindResource("AccentSoftBrush");
            b.BorderBrush = (Brush)FindResource("AccentBrush");
        }
        e.Effects = HasFiles(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void DropZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = HasFiles(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void DropZone_DragLeave(object sender, DragEventArgs e)
    {
        if (sender is Border b)
        {
            b.Background = (Brush)FindResource("BgSurfaceBrush");
            b.BorderBrush = (Brush)FindResource("BorderBrush");
        }
    }

    private void DropZone_Drop(object sender, DragEventArgs e)
    {
        if (sender is Border b)
        {
            b.Background = (Brush)FindResource("BgSurfaceBrush");
            b.BorderBrush = (Brush)FindResource("BorderBrush");
        }
        if (!HasFiles(e)) return;
        var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
        Vm.AddFiles(paths);
        e.Handled = true;
    }

    private void DropZone_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) => OpenFileDialog();

    private void BrowseButton_Click(object sender, RoutedEventArgs e) => OpenFileDialog();

    private void OpenFileDialog()
    {
        if (Vm.IsProcessing) return;
        var dlg = new OpenFileDialog
        {
            Filter = "Excel Workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Multiselect = true,
            Title = "Select Excel files"
        };
        if (dlg.ShowDialog(this) == true)
        {
            Vm.AddFiles(dlg.FileNames);
        }
    }

    private void ScrollLogToEnd()
    {
        if (LogList.Items.Count == 0) return;
        var last = LogList.Items[LogList.Items.Count - 1];
        LogList.ScrollIntoView(last);
    }
}
