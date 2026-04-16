using System.IO;
using System.IO.Pipes;
using System.Security.Principal;
using System.Text;

namespace LaptopSessionViewer.Services;

public sealed class SingleInstanceService : IDisposable
{
    private readonly AppLogService _logService;
    private readonly string _mutexName;
    private readonly string _pipeName;
    private Mutex? _mutex;
    private CancellationTokenSource? _listenerCancellation;
    private Task? _listenerTask;

    public SingleInstanceService(string applicationId, AppLogService logService)
    {
        _logService = logService;

        var identity = WindowsIdentity.GetCurrent();
        var instanceScope = identity.User?.Value ?? Environment.UserName;
        var safeName = NormalizeName($"{applicationId}-{instanceScope}");

        _mutexName = $@"Local\{safeName}";
        _pipeName = safeName;
    }

    public bool TryAcquirePrimaryInstance()
    {
        _mutex = new Mutex(initiallyOwned: true, _mutexName, out var createdNew);

        if (createdNew)
        {
            return true;
        }

        _mutex.Dispose();
        _mutex = null;
        return false;
    }

    public void StartActivationListener(Action onActivationRequested)
    {
        if (_listenerTask is not null)
        {
            return;
        }

        _listenerCancellation = new CancellationTokenSource();
        var cancellationToken = _listenerCancellation.Token;

        _listenerTask = Task.Run(
            async () =>
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        await using var server = new NamedPipeServerStream(
                            _pipeName,
                            PipeDirection.In,
                            1,
                            PipeTransmissionMode.Byte,
                            PipeOptions.Asynchronous);

                        await server.WaitForConnectionAsync(cancellationToken).ConfigureAwait(false);

                        using var reader = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: false);
                        var command = (await reader.ReadLineAsync().ConfigureAwait(false))?.Trim();

                        if (string.Equals(command, "ACTIVATE", StringComparison.OrdinalIgnoreCase))
                        {
                            onActivationRequested();
                        }
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        break;
                    }
                    catch (Exception exception)
                    {
                        _logService.Error(nameof(SingleInstanceService), "Single-instance listener failed.", exception);

                        try
                        {
                            await Task.Delay(300, cancellationToken).ConfigureAwait(false);
                        }
                        catch (OperationCanceledException)
                        {
                            break;
                        }
                    }
                }
            },
            cancellationToken);
    }

    public async Task<bool> SignalPrimaryInstanceAsync(TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var remaining = deadline - DateTime.UtcNow;
                var timeoutMilliseconds = (int)Math.Max(250, Math.Min(remaining.TotalMilliseconds, 1000));

                await using var client = new NamedPipeClientStream(
                    ".",
                    _pipeName,
                    PipeDirection.Out,
                    PipeOptions.Asynchronous);

                await client.ConnectAsync(timeoutMilliseconds).ConfigureAwait(false);

                await using var writer = new StreamWriter(client, new UTF8Encoding(false))
                {
                    AutoFlush = true
                };

                await writer.WriteLineAsync("ACTIVATE").ConfigureAwait(false);
                return true;
            }
            catch (TimeoutException)
            {
            }
            catch (IOException)
            {
            }
            catch (Exception exception)
            {
                _logService.Error(nameof(SingleInstanceService), "Failed to signal the primary AIHelper instance.", exception);
                break;
            }

            await Task.Delay(150).ConfigureAwait(false);
        }

        return false;
    }

    public void Dispose()
    {
        if (_listenerCancellation is not null)
        {
            _listenerCancellation.Cancel();
        }

        try
        {
            _listenerTask?.Wait(500);
        }
        catch
        {
        }

        _listenerCancellation?.Dispose();
        _listenerCancellation = null;
        _listenerTask = null;

        if (_mutex is not null)
        {
            try
            {
                _mutex.ReleaseMutex();
            }
            catch (ApplicationException)
            {
            }

            _mutex.Dispose();
            _mutex = null;
        }
    }

    private static string NormalizeName(string value)
    {
        var builder = new StringBuilder(value.Length);

        foreach (var character in value)
        {
            builder.Append(char.IsLetterOrDigit(character) ? character : '-');
        }

        return builder.ToString();
    }
}
