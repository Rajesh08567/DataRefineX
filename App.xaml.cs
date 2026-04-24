using System.Windows;
using System.Windows.Threading;

namespace DataRefineX;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        DispatcherUnhandledException += (_, args) =>
        {
            MessageBox.Show(
                $"An unexpected error occurred:\n\n{args.Exception.Message}",
                "DataRefineX",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            args.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                MessageBox.Show(
                    $"A fatal error occurred:\n\n{ex.Message}",
                    "DataRefineX",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        };
    }
}
