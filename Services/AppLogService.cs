using System.Diagnostics;
using System.IO;
using System.Text;

namespace LaptopSessionViewer.Services;

public sealed class AppLogService
{
    private readonly string _logPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AIHelper",
            "aihelper.log");

    public void Info(string source, string message)
    {
        Write("INFO", source, message);
    }

    public void Error(string source, string message, Exception? exception = null)
    {
        Write("ERROR", source, message, exception);
    }

    private void Write(string level, string source, string message, Exception? exception = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(_logPath);

            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var builder = new StringBuilder()
                .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                .Append(" [")
                .Append(level)
                .Append("] ")
                .Append(source)
                .Append(": ")
                .AppendLine(message);

            if (exception is not null)
            {
                builder.AppendLine(exception.ToString());
            }

            File.AppendAllText(_logPath, builder.ToString(), Encoding.UTF8);
            Trace.WriteLine(builder.ToString());
        }
        catch
        {
        }
    }
}
