using System.Windows;
using System.Windows.Threading;
using System.Windows.Interop;
using LaptopSessionViewer.Services;

namespace LaptopSessionViewer;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private readonly AppLogService _logService = new();
    private bool _fatalErrorShown;
    private SingleInstanceService? _singleInstanceService;

    protected override void OnStartup(StartupEventArgs e)
    {
        RegisterGlobalExceptionHandlers();
        LogStartupContext();

        _singleInstanceService = new SingleInstanceService("AIHelper", _logService);

        if (!_singleInstanceService.TryAcquirePrimaryInstance())
        {
            TryActivateExistingInstance();
            Shutdown(0);
            return;
        }

        _singleInstanceService.StartActivationListener(ActivatePrimaryWindow);
        base.OnStartup(e);

        try
        {
            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Show();
            ActivatePrimaryWindow();
        }
        catch (Exception exception)
        {
            ShowFatalStartupError("AIHelper failed to start.", exception);
            Shutdown(-1);
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _singleInstanceService?.Dispose();
        _singleInstanceService = null;
        base.OnExit(e);
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

    private void TryActivateExistingInstance()
    {
        try
        {
            _singleInstanceService?
                .SignalPrimaryInstanceAsync(TimeSpan.FromSeconds(3))
                .GetAwaiter()
                .GetResult();
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(App), "Failed to activate the already running AIHelper instance.", exception);
        }
    }

    private void ActivatePrimaryWindow()
    {
        Dispatcher.BeginInvoke(
            () =>
            {
                if (MainWindow is null)
                {
                    return;
                }

                BringWindowToFront(MainWindow);
            });
    }

    private static void BringWindowToFront(Window window)
    {
        if (window.WindowState == WindowState.Minimized)
        {
            window.WindowState = WindowState.Normal;
        }

        if (!window.IsVisible)
        {
            window.Show();
        }

        window.Activate();
        window.Topmost = true;
        window.Topmost = false;
        window.Focus();

        var handle = new WindowInteropHelper(window).Handle;

        if (handle != IntPtr.Zero)
        {
            NativeMethods.ShowWindow(handle, NativeMethods.SW_RESTORE);
            NativeMethods.SetForegroundWindow(handle);
        }
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

    private static class NativeMethods
    {
        public const int SW_RESTORE = 9;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
