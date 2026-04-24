using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using DataRefineX.Models;

namespace DataRefineX.Converters;

public sealed class BoolToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }
    public bool Collapse { get; set; } = true;

    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        var b = value is bool bv && bv;
        if (Invert) b = !b;
        return b ? Visibility.Visible : (Collapse ? Visibility.Collapsed : Visibility.Hidden);
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b ? !b : true;

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b ? !b : true;
}

public sealed class CountToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }

    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        var hasAny = value switch
        {
            int i => i > 0,
            long l => l > 0,
            System.Collections.ICollection c => c.Count > 0,
            _ => false
        };
        if (Invert) hasAny = !hasAny;
        return hasAny ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class FileStatusToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not FileStatus status) return DependencyProperty.UnsetValue;
        var key = status switch
        {
            FileStatus.Queued => "TextMutedBrush",
            FileStatus.Reading => "InfoBrush",
            FileStatus.Processed => "SuccessBrush",
            FileStatus.Failed => "DangerBrush",
            FileStatus.Skipped => "WarningBrush",
            _ => "TextMutedBrush"
        };
        return Application.Current.TryFindResource(key) as Brush ?? Brushes.Gray;
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class LogLevelToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not LogLevel level) return DependencyProperty.UnsetValue;
        var key = level switch
        {
            LogLevel.Info => "TextSecondaryBrush",
            LogLevel.Success => "SuccessBrush",
            LogLevel.Warning => "WarningBrush",
            LogLevel.Error => "DangerBrush",
            _ => "TextSecondaryBrush"
        };
        return Application.Current.TryFindResource(key) as Brush ?? Brushes.Gray;
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class LongToNumberConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value switch
        {
            long l => l.ToString("N0", culture),
            int i => i.ToString("N0", culture),
            _ => value?.ToString() ?? string.Empty
        };
    }

    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class NullOrEmptyToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }
    public object Convert(object value, Type targetType, object? parameter, CultureInfo culture)
    {
        var isEmpty = value is null || (value is string s && string.IsNullOrWhiteSpace(s));
        var visible = Invert ? isEmpty : !isEmpty;
        return visible ? Visibility.Visible : Visibility.Collapsed;
    }
    public object ConvertBack(object value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
