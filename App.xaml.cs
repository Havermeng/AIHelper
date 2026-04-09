using System.Windows;
using System.Windows.Threading;
using LaptopSessionViewer.Services;

namespace LaptopSessionViewer;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private readonly AppLogService _logService = new();
    private bool _fatalErrorShown;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        RegisterGlobalExceptionHandlers();
        LogStartupContext();

        try
        {
            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Show();
        }
        catch (Exception exception)
        {
            ShowFatalStartupError("AIHelper failed to start.", exception);
            Shutdown(-1);
        }
    }

    private void RegisterGlobalExceptionHandlers()
    {
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        ShowFatalStartupError("AIHelper crashed because of an unhandled UI exception.", e.Exception);
        e.Handled = true;
        Shutdown(-1);
    }

    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        var exception = e.ExceptionObject as Exception
                        ?? new Exception(e.ExceptionObject?.ToString() ?? "Unknown fatal exception.");
        ShowFatalStartupError("AIHelper crashed because of an unhandled exception.", exception);
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        ShowFatalStartupError("AIHelper crashed because of an unobserved task exception.", e.Exception);
        e.SetObserved();
    }

    private void LogStartupContext()
    {
        _logService.Info(
            nameof(App),
            $"Starting AIHelper. OS={Environment.OSVersion}; Is64BitOS={Environment.Is64BitOperatingSystem}; " +
            $"Is64BitProcess={Environment.Is64BitProcess}; Runtime={Environment.Version}; BaseDir={AppContext.BaseDirectory}");
    }

    private void ShowFatalStartupError(string message, Exception exception)
    {
        _logService.Error(nameof(App), message, exception);

        if (_fatalErrorShown)
        {
            return;
        }

        _fatalErrorShown = true;

        try
        {
            MessageBox.Show(
                $"{message}\n\n{exception.Message}\n\nLog: {_logService.LogPath}",
                "AIHelper",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        catch
        {
        }
    }
}
