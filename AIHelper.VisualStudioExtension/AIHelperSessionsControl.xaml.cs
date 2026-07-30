using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.VisualStudio.Shell;

namespace AIHelper.VisualStudioExtension;

public partial class AIHelperSessionsControl : UserControl, INotifyPropertyChanged
{
    private readonly ObservableCollection<AIHelperSessionItem> _sessions = new();
    private readonly ObservableCollection<string> _resumeTools = new()
    {
        "Codex",
        "OpenCode",
        "Qwen",
        "Claude",
        "Gemini",
        "Kilo Code"
    };
    private string _searchText = string.Empty;
    private AIHelperSessionItem? _selectedSession;
    private string _selectedResumeTool = "Codex";
    private string _statusText = string.Empty;

    public AIHelperSessionsControl()
    {
        InitializeComponent();
        FilteredSessions = CollectionViewSource.GetDefaultView(_sessions);
        FilteredSessions.Filter = FilterSession;
        DataContext = this;
        LoadSessions();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ICollectionView FilteredSessions { get; }

    public ObservableCollection<string> ResumeTools => _resumeTools;

    public string FeedPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "sessions-feed.json");

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetField(ref _searchText, value))
            {
                FilteredSessions.Refresh();
            }
        }
    }

    public AIHelperSessionItem? SelectedSession
    {
        get => _selectedSession;
        set
        {
            if (SetField(ref _selectedSession, value) && value is not null)
            {
                SelectedResumeTool = GetDefaultResumeTool(value);
            }
        }
    }

    public string SelectedResumeTool
    {
        get => _selectedResumeTool;
        set => SetField(ref _selectedResumeTool, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetField(ref _statusText, value);
    }

    private void LoadSessions()
    {
        _sessions.Clear();

        if (!File.Exists(FeedPath))
        {
            StatusText = "Фид сессий не найден. Откройте AIHelper и обновите список сессий.";
            return;
        }

        try
        {
            var json = File.ReadAllText(FeedPath);
            var serializer = new JavaScriptSerializer();
            var feed = serializer.Deserialize<AIHelperSessionFeed>(json) ?? new AIHelperSessionFeed();

            foreach (var session in feed.Sessions ?? [])
            {
                _sessions.Add(session);
            }

            SelectedSession = _sessions.FirstOrDefault();
            StatusText = $"Загружено сессий: {_sessions.Count}. Фид обновлен: {feed.GeneratedAtUtc}";
            FilteredSessions.Refresh();
        }
        catch (Exception exception)
        {
            StatusText = $"Не удалось прочитать сессии AIHelper: {exception.Message}";
        }
    }

    private bool FilterSession(object item)
    {
        if (item is not AIHelperSessionItem session)
        {
            return false;
        }

        var filter = SearchText.Trim();
        if (string.IsNullOrWhiteSpace(filter))
        {
            return true;
        }

        return Contains(session.Title, filter) ||
               Contains(session.OriginalTitle, filter) ||
               Contains(session.Preview, filter) ||
               Contains(session.LastMessagePreview, filter) ||
               Contains(session.Source, filter) ||
               Contains(session.ModelProvider, filter) ||
               Contains(session.SessionId, filter) ||
               Contains(session.WorkingDirectory, filter) ||
               Contains(session.FilePath, filter);
    }

    private static bool Contains(string? value, string filter)
    {
        return value?.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        LoadSessions();
    }

    private void OpenFileButton_Click(object sender, RoutedEventArgs e)
    {
        ThreadHelper.ThrowIfNotOnUIThread();
        var path = SelectedSession?.FilePath;

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            StatusText = "Файл сессии не найден.";
            return;
        }

        VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var directory = SelectedSession?.WorkingDirectory;

        if (string.IsNullOrWhiteSpace(directory) ||
            directory == "-" ||
            !Directory.Exists(directory))
        {
            directory = Path.GetDirectoryName(SelectedSession?.FilePath ?? string.Empty);
        }

        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            StatusText = "Папка не найдена.";
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"\"{directory}\"",
            UseShellExecute = true
        });
    }

    private void CopyIdButton_Click(object sender, RoutedEventArgs e)
    {
        var sessionId = SelectedSession?.SessionId;

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        Clipboard.SetText(sessionId);
        StatusText = "ID сессии скопирован.";
    }

    private void OpenFeedButton_Click(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(FeedPath))
        {
            StatusText = "Файл фида не найден.";
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{FeedPath}\"",
            UseShellExecute = true
        });
    }

    private void ResumeButton_Click(object sender, RoutedEventArgs e)
    {
        var session = SelectedSession;

        if (session is null)
        {
            StatusText = "Select a session first.";
            return;
        }

        try
        {
            StatusText = AIHelperSessionLauncher.Launch(session, SelectedResumeTool);
        }
        catch (Exception exception)
        {
            StatusText = $"Failed to launch {SelectedResumeTool}: {exception.Message}";
        }
    }

    private void CopyResumeCommandButton_Click(object sender, RoutedEventArgs e)
    {
        var session = SelectedSession;

        if (session is null)
        {
            StatusText = "Select a session first.";
            return;
        }

        try
        {
            var plan = AIHelperSessionLauncher.BuildPlan(session, SelectedResumeTool);
            var text = string.IsNullOrWhiteSpace(plan.ClipboardPrompt)
                ? plan.Target
                : plan.ClipboardPrompt;

            Clipboard.SetText(text);
            StatusText = string.IsNullOrWhiteSpace(plan.ClipboardPrompt)
                ? $"{SelectedResumeTool} command copied."
                : $"{SelectedResumeTool} handoff prompt copied.";
        }
        catch (Exception exception)
        {
            StatusText = $"Failed to build {SelectedResumeTool} command: {exception.Message}";
        }
    }

    private static string GetDefaultResumeTool(AIHelperSessionItem session)
    {
        return session.Source switch
        {
            "OpenCode" => "OpenCode",
            "Qwen" => "Qwen",
            "Claude" => "Claude",
            "Gemini" => "Gemini",
            _ => "Codex"
        };
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}
