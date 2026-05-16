using System.Windows;
using System.Windows.Threading;
using System.Windows.Interop;
using System.IO;
using System.Runtime;
using LaptopSessionViewer.Services;

namespace LaptopSessionViewer;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private const string StartMinimizedArgument = "--start-minimized";
    private readonly AppLogService _logService = new();
    private bool _fatalErrorShown;
    private SingleInstanceService? _singleInstanceService;

    protected override void OnStartup(StartupEventArgs e)
    {
        RegisterGlobalExceptionHandlers();
        LogStartupContext();
        ConfigureStartupOptimization();

        var startMinimized = HasStartMinimizedArgument(e.Args);

        _singleInstanceService = new SingleInstanceService("AIHelper", _logService);

        if (!_singleInstanceService.TryAcquirePrimaryInstance())
        {
            if (!startMinimized)
            {
                TryActivateExistingInstance();
            }

            Shutdown(0);
            return;
        }

        _singleInstanceService.StartActivationListener(ActivatePrimaryWindow);
        base.OnStartup(e);

        StartupWindow? startupWindow = null;

        try
        {
            if (!startMinimized)
            {
                startupWindow = new StartupWindow();
                startupWindow.Show();
                startupWindow.Dispatcher.Invoke(() => { }, DispatcherPriority.Render);
            }

            var mainWindow = new MainWindow();
            if (startMinimized)
            {
                mainWindow.ShowActivated = false;
                mainWindow.WindowState = WindowState.Minimized;
            }

            MainWindow = mainWindow;
            mainWindow.ContentRendered += (_, _) =>
            {
                try
                {
                    startupWindow?.Close();
                }
                catch
                {
                }
            };
            mainWindow.Show();

            if (!startMinimized)
            {
                ActivatePrimaryWindow();
            }
            else
            {
                _logService.Info(nameof(App), "Started minimized after silent update.");
            }
        }
        catch (Exception exception)
        {
            try
            {
                startupWindow?.Close();
            }
            catch
            {
            }

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

    private void ConfigureStartupOptimization()
    {
        try
        {
            var profileDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AIHelper",
                "jit-profiles");
            Directory.CreateDirectory(profileDirectory);
            ProfileOptimization.SetProfileRoot(profileDirectory);
            ProfileOptimization.StartProfile("startup.profile");
            _logService.Info(nameof(App), $"Startup profile optimization enabled. Root={profileDirectory}");
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(App), "Failed to enable startup profile optimization.", exception);
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

    private static bool HasStartMinimizedArgument(IEnumerable<string> args)
    {
        return args.Any(argument =>
            string.Equals(argument, StartMinimizedArgument, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(argument, "/start-minimized", StringComparison.OrdinalIgnoreCase));
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
