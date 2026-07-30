using System.IO;
using System.Text;
using LaptopSessionViewer.Models;

namespace LaptopSessionViewer.Services;

public sealed class SessionCheckpointService
{
    private const int MaximumContextCharacters = 16000;

    public string CheckpointDirectory =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "AIHelper Checkpoints");

    public string CreateCheckpoint(
        SessionRecord session,
        string transcript,
        AppLanguage language)
    {
        Directory.CreateDirectory(CheckpointDirectory);
        var title = SanitizeFileName(session.DisplayTitle);
        var baseName = string.IsNullOrWhiteSpace(title) ? session.ShortSessionId : title;
        var fileName = $"{DateTime.Now:yyyyMMdd-HHmmss}-{TrimFileName(baseName, 64)}.md";
        var path = CreateUniquePath(Path.Combine(CheckpointDirectory, fileName));
        var content = BuildContent(session, transcript, language);
        File.WriteAllText(path, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return path;
    }

    private static string BuildContent(
        SessionRecord session,
        string transcript,
        AppLanguage language)
    {
        var russian = language == AppLanguage.Russian;
        var context = PrepareRecentContext(transcript);
        var builder = new StringBuilder();

        builder.AppendLine(russian ? "# Контрольная точка AIHelper" : "# AIHelper checkpoint");
        builder.AppendLine();
        builder.AppendLine($"{(russian ? "Создано" : "Created")}: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        builder.AppendLine($"{(russian ? "Сессия" : "Session")}: {session.SessionId}");
        builder.AppendLine($"{(russian ? "Источник" : "Source")}: {session.Source}");
        builder.AppendLine($"{(russian ? "Модель" : "Model")}: {session.ModelProvider}");
        builder.AppendLine($"{(russian ? "Рабочая папка" : "Working directory")}: {session.WorkingDirectory}");
        builder.AppendLine(
            $"{(russian ? "Сообщения и инструменты" : "Messages and tools")}: {session.TotalMessageCount} / {session.ToolCallCount}");
        builder.AppendLine();

        AppendSection(
            builder,
            russian ? "Исходная задача" : "Original task",
            session.Preview);
        AppendSection(
            builder,
            russian ? "Последний видимый результат" : "Last visible result",
            session.LastMessagePreview);
        AppendSection(
            builder,
            russian ? "Заметка пользователя" : "User note",
            session.Note);
        AppendSection(
            builder,
            russian ? "Недавний контекст" : "Recent context",
            context);

        builder.AppendLine(russian ? "## Как продолжить" : "## How to continue");
        builder.AppendLine();
        builder.AppendLine(
            russian
                ? "Сначала прочитай эту контрольную точку, проверь текущее состояние файлов в рабочей папке и продолжи незавершённую задачу. Не считай старые предположения подтверждёнными без проверки."
                : "Read this checkpoint first, inspect the current files in the working directory, and continue the unfinished task. Do not treat old assumptions as confirmed without checking them.");

        return builder.ToString();
    }

    private static void AppendSection(StringBuilder builder, string title, string text)
    {
        builder.AppendLine($"## {title}");
        builder.AppendLine();
        builder.AppendLine(string.IsNullOrWhiteSpace(text) ? "—" : text.Trim());
        builder.AppendLine();
    }

    private static string PrepareRecentContext(string transcript)
    {
        if (string.IsNullOrWhiteSpace(transcript))
        {
            return "—";
        }

        var normalized = transcript.Trim();
        return normalized.Length <= MaximumContextCharacters
            ? normalized
            : $"[…]\n{normalized[^MaximumContextCharacters..]}";
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars().ToHashSet();
        return new string(
                value
                    .Where(character => !invalid.Contains(character) && !char.IsControl(character))
                    .ToArray())
            .Trim()
            .TrimEnd('.');
    }

    private static string TrimFileName(string value, int maxLength)
    {
        return value.Length <= maxLength ? value : value[..maxLength].TrimEnd();
    }

    private static string CreateUniquePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);

        for (var index = 2; index < 1000; index++)
        {
            var candidate = Path.Combine(directory, $"{name}-{index}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(directory, $"{name}-{Guid.NewGuid():N}{extension}");
    }
}
