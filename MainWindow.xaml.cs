using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using LaptopSessionViewer.Models;
using LaptopSessionViewer.Services;
using Microsoft.Win32;

namespace LaptopSessionViewer;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly AppLogService _logService = new();
    private readonly CodexEnvironmentService _environmentService = new();
    private readonly AppUpdateService _updateService = new();
    private readonly DnsManagementService _dnsManagementService = new();
    private readonly DnsPresetSettingsService _dnsPresetSettingsService = new();
    private readonly SessionFavoritesService _favoritesService = new();
    private readonly SessionNotesService _notesService = new();
    private readonly SessionService _sessionService = new();
    private readonly SessionViewerSettingsService _settingsService = new();
    private readonly DispatcherTimer _refreshTimer;
    private List<SessionRecord> _allSessions = [];
    private HashSet<string> _favoriteSessionIds = [];
    private Dictionary<string, string> _sessionNotes = new(StringComparer.OrdinalIgnoreCase);
    private bool _autoRefreshEnabled = true;
    private bool _isLoading;
    private bool _isRefreshing;
    private bool _isDnsBusy;
    private bool _isApplyingDangerousAccessDefaults;
    private bool _isSetupBusy;
    private bool _isSetupCodexSectionExpanded;
    private bool _isSetupCoreSectionExpanded;
    private bool _isSetupDnsSectionExpanded;
    private bool _isSetupLocalAiSectionExpanded;
    private bool _isUpdateBusy;
    private AppUpdateSnapshot? _lastAppUpdateSnapshot;
    private CodexEnvironmentSnapshot? _lastEnvironmentSnapshot;
    private string _configuredCodexModel = string.Empty;
    private string _dnsDohTemplate = string.Empty;
    private string _dnsStatusForeground = "#F8E7D6";
    private string _dnsStatusText = string.Empty;
    private bool _dnsUseDoh;
    private DateTime? _lastUpdatedAtLocal;
    private string _lastUpdatedText = string.Empty;
    private string _newSessionModel = string.Empty;
    private string _newSessionProfile = string.Empty;
    private string _newSessionPrompt = string.Empty;
    private string _newSessionStatusForeground = "#F8E7D6";
    private string _newSessionStatusText = string.Empty;
    private bool _newSessionUseFullAuto;
    private bool _newSessionUseOss;
    private bool _newSessionUseSearch;
    private string _newSessionWorkingDirectory =
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    private string _primaryDnsServer = string.Empty;
    private DnsAdapterRecord? _selectedDnsAdapter;
    private DnsPreset? _selectedDnsPreset;
    private string _selectedApprovalPolicy = "on-request";
    private AppSection _selectedAppSection = AppSection.Sessions;
    private string _searchText = string.Empty;
    private string _secondaryDnsServer = string.Empty;
    private string _selectedLocalProvider = string.Empty;
    private string _selectedSandboxMode = "workspace-write";
    private SessionListTab _selectedSessionListTab = SessionListTab.Sessions;
    private LanguageOption? _selectedLanguageOption;
    private bool _settingsDangerousFullAccess;
    private string _settingsStatusForeground = "#F8E7D6";
    private string _settingsStatusKey = "SettingsStatusReady";
    private object[] _settingsStatusArgs = [];
    private string _settingsStatusText = string.Empty;
    private string _selectedSessionNote = string.Empty;
    private SessionRecord? _selectedSession;
    private string _setupStatusForeground = "#F8E7D6";
    private string _setupStatusText = string.Empty;
    private string _updateStatusForeground = "#F8E7D6";
    private string _updateStatusKey = "UpdateStatusReady";
    private object[] _updateStatusArgs = [];
    private string _updateStatusText = string.Empty;
    private string _statusForeground = "#F8E7D6";
    private string _statusKey = "StatusReady";
    private object[] _statusArgs = [];
    private string _statusText = string.Empty;
    private int _totalMessages;
    private int _totalSessions;
    private int _totalToolCalls;
    private int _updatedTodaySessions;

    public MainWindow()
    {
        var initialSettings = _settingsService.LoadSettings();
        Strings.SetLanguage(initialSettings.Language);
        _settingsDangerousFullAccess = initialSettings.DefaultDangerousFullAccess;
        _selectedLanguageOption = LanguageOptions.First(option => option.Language == initialSettings.Language);

        InitializeComponent();
        DataContext = this;

        RefreshLaunchOptionCollections();
        RefreshLocalAiModelOptions();
        RefreshCreativeAiToolOptions();
        RefreshAiAgentToolOptions();
        LoadNewSessionConfigurationInfoSafe();
        ApplyDangerousAccessDefaultsToNewSession();
        LoadDnsPresetsSafe();
        RefreshLocalizedChromeText();
        RefreshSectionChromeText();

        _refreshTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(15)
        };

        SourceInitialized += MainWindow_SourceInitialized;
        _refreshTimer.Tick += RefreshTimer_Tick;
        Loaded += MainWindow_Loaded;
        Closed += MainWindow_Closed;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public LocalizationService Strings { get; } = new();

    public ObservableCollection<DnsAdapterRecord> DnsAdapters { get; } = [];

    public ObservableCollection<DnsPreset> DnsPresets { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupCoreChecks { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupCodexChecks { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupLocalAiChecks { get; } = [];

    public ObservableCollection<LocalAiModelOption> LocalAiModelOptions { get; } = [];

    public ObservableCollection<CreativeAiToolOption> CreativeAiToolOptions { get; } = [];

    public ObservableCollection<CreativeAiToolOption> AiAgentToolOptions { get; } = [];

    public ObservableCollection<LaunchOption> SandboxModeOptions { get; } = [];

    public ObservableCollection<LaunchOption> ApprovalPolicyOptions { get; } = [];

    public ObservableCollection<LaunchOption> LocalProviderOptions { get; } = [];

    public ObservableCollection<string> ModelSuggestions { get; } = [];

    public ObservableCollection<string> ProfileSuggestions { get; } = [];

    public string PrimaryDnsServer
    {
        get => _primaryDnsServer;
        set
        {
            if (SetField(ref _primaryDnsServer, value))
            {
                OnPropertyChanged(nameof(CanApplyDnsPreset));
            }
        }
    }

    public string SecondaryDnsServer
    {
        get => _secondaryDnsServer;
        set
        {
            if (SetField(ref _secondaryDnsServer, value))
            {
                OnPropertyChanged(nameof(CanApplyDnsPreset));
            }
        }
    }

    public bool DnsUseDoh
    {
        get => _dnsUseDoh;
        set
        {
            if (SetField(ref _dnsUseDoh, value))
            {
                if (!value)
                {
                    DnsDohTemplate = string.Empty;
                }

                OnPropertyChanged(nameof(CanApplyDnsPreset));
                OnPropertyChanged(nameof(DnsDohTemplateVisibility));
            }
        }
    }

    public string DnsDohTemplate
    {
        get => _dnsDohTemplate;
        set
        {
            if (SetField(ref _dnsDohTemplate, value))
            {
                OnPropertyChanged(nameof(CanApplyDnsPreset));
            }
        }
    }

    public IReadOnlyList<LanguageOption> LanguageOptions { get; } =
    [
        new LanguageOption
        {
            Language = AppLanguage.English,
            DisplayName = "English"
        },
        new LanguageOption
        {
            Language = AppLanguage.Russian,
            DisplayName = "\u0420\u0443\u0441\u0441\u043a\u0438\u0439"
        }
    ];

    public ObservableCollection<SessionRecord> Sessions { get; } = [];

    public AppSection SelectedAppSection
    {
        get => _selectedAppSection;
        set
        {
            if (SetField(ref _selectedAppSection, value))
            {
                OnPropertyChanged(nameof(SessionsSectionButtonBackground));
                OnPropertyChanged(nameof(SessionsSectionButtonForeground));
                OnPropertyChanged(nameof(NewSessionSectionButtonBackground));
                OnPropertyChanged(nameof(NewSessionSectionButtonForeground));
                OnPropertyChanged(nameof(SetupSectionButtonBackground));
                OnPropertyChanged(nameof(SetupSectionButtonForeground));
                OnPropertyChanged(nameof(SettingsSectionButtonBackground));
                OnPropertyChanged(nameof(SettingsSectionButtonForeground));
                OnPropertyChanged(nameof(SessionsSectionVisibility));
                OnPropertyChanged(nameof(NewSessionSectionVisibility));
                OnPropertyChanged(nameof(SetupSectionVisibility));
                OnPropertyChanged(nameof(SettingsSectionVisibility));

                if (value == AppSection.Setup && IsLoaded && _lastEnvironmentSnapshot is null)
                {
                    _ = RefreshSetupStatusAsync();
                }

                if (value == AppSection.Settings && IsLoaded && _lastAppUpdateSnapshot is null)
                {
                    _ = RefreshSettingsSectionAsync();
                }
            }
        }
    }

    public string SessionsSectionButtonBackground =>
        SelectedAppSection == AppSection.Sessions ? "#D97732" : "#1D3545";

    public string SessionsSectionButtonForeground => "#FFFDF9";

    public string NewSessionSectionButtonBackground =>
        SelectedAppSection == AppSection.NewSession ? "#D97732" : "#1D3545";

    public string NewSessionSectionButtonForeground => "#FFFDF9";

    public string SetupSectionButtonBackground =>
        SelectedAppSection == AppSection.Setup ? "#D97732" : "#1D3545";

    public string SetupSectionButtonForeground => "#FFFDF9";

    public string SettingsSectionButtonBackground =>
        SelectedAppSection == AppSection.Settings ? "#D97732" : "#1D3545";

    public string SettingsSectionButtonForeground => "#FFFDF9";

    public Visibility SessionsSectionVisibility =>
        SelectedAppSection == AppSection.Sessions ? Visibility.Visible : Visibility.Collapsed;

    public Visibility NewSessionSectionVisibility =>
        SelectedAppSection == AppSection.NewSession ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SetupSectionVisibility =>
        SelectedAppSection == AppSection.Setup ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SettingsSectionVisibility =>
        SelectedAppSection == AppSection.Settings ? Visibility.Visible : Visibility.Collapsed;

    public string SessionsTabText => $"{Strings["SessionsTab"]} ({RegularSessions})";

    public string FavoritesTabText => $"{Strings["FavoritesTab"]} ({FavoriteSessions})";

    public bool HasVisibleSessions => Sessions.Count > 0;

    public string EmptySessionsText =>
        SelectedSessionListTab == SessionListTab.Favorites
            ? Strings["EmptyFavoritesTab"]
            : Strings["EmptySessionsTab"];

    public SessionListTab SelectedSessionListTab
    {
        get => _selectedSessionListTab;
        set
        {
            if (SetField(ref _selectedSessionListTab, value))
            {
                OnPropertyChanged(nameof(SessionsTabBackground));
                OnPropertyChanged(nameof(SessionsTabForeground));
                OnPropertyChanged(nameof(FavoritesTabBackground));
                OnPropertyChanged(nameof(FavoritesTabForeground));
                OnPropertyChanged(nameof(EmptySessionsText));
                ApplyFilter();
            }
        }
    }

    public string SessionsTabBackground =>
        SelectedSessionListTab == SessionListTab.Sessions ? "#16212B" : "#E7D8CA";

    public string SessionsTabForeground =>
        SelectedSessionListTab == SessionListTab.Sessions ? "#FFFDF9" : "#16212B";

    public string FavoritesTabBackground =>
        SelectedSessionListTab == SessionListTab.Favorites ? "#16212B" : "#E7D8CA";

    public string FavoritesTabForeground =>
        SelectedSessionListTab == SessionListTab.Favorites ? "#FFFDF9" : "#16212B";

    public bool AutoRefreshEnabled
    {
        get => _autoRefreshEnabled;
        set
        {
            if (SetField(ref _autoRefreshEnabled, value))
            {
                UpdateRefreshTimer();
            }
        }
    }

    public bool CanOpenSelectedFile =>
        SelectedSession is not null && File.Exists(SelectedSession.FilePath);

    public bool CanDeleteSelectedSession =>
        SelectedSession is not null && File.Exists(SelectedSession.FilePath);

    public bool CanResumeSelectedSession =>
        SelectedSession is not null &&
        !string.IsNullOrWhiteSpace(SelectedSession.SessionId) &&
        File.Exists(_environmentService.CodexCommandPath);

    public bool CanEditSelectedSessionNote => SelectedSession is not null;

    public bool CanSaveSelectedSessionNote =>
        SelectedSession is not null &&
        !string.Equals(
            NormalizeNote(SelectedSessionNote),
            NormalizeNote(SelectedSession.Note),
            StringComparison.Ordinal);

    public bool CanClearSelectedSessionNote =>
        SelectedSession is not null && !string.IsNullOrWhiteSpace(SelectedSessionNote);

    public bool CanUseSelectedSessionDirectory =>
        SelectedSession is not null &&
        Directory.Exists(SelectedSession.WorkingDirectory);

    public bool CanLaunchNewSession =>
        File.Exists(_environmentService.CodexCommandPath) &&
        Directory.Exists(GetNormalizedNewSessionWorkingDirectory());

    public bool CanLaunchCodexLogin => File.Exists(_environmentService.CodexCommandPath);

    public bool CanInstallLocalAiTools => !IsSetupBusy;

    public bool CanInstallLocalAiModels =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAvailable == true;

    public bool CanManageCreativeAiTools => !IsSetupBusy;

    public bool CanManageAiAgents => !IsSetupBusy;

    public bool CanUninstallOllama =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAvailable == true;

    public bool CanUninstallLmStudio =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.LmStudioAvailable == true;

    public bool IsDnsBusy
    {
        get => _isDnsBusy;
        private set
        {
            if (SetField(ref _isDnsBusy, value))
            {
                RefreshDnsCommandStates();
            }
        }
    }

    public DnsAdapterRecord? SelectedDnsAdapter
    {
        get => _selectedDnsAdapter;
        set
        {
            if (SetField(ref _selectedDnsAdapter, value))
            {
                OnPropertyChanged(nameof(SelectedDnsAdapterDescriptionText));
                OnPropertyChanged(nameof(SelectedDnsAdapterServersText));
                OnPropertyChanged(nameof(CanApplyDnsPreset));
                OnPropertyChanged(nameof(CanResetDnsAutomatic));
                OnPropertyChanged(nameof(CanRestorePreviousDns));
            }
        }
    }

    public DnsPreset? SelectedDnsPreset
    {
        get => _selectedDnsPreset;
        set
        {
            if (SetField(ref _selectedDnsPreset, value))
            {
                ApplyDnsPresetToEditors(value);
                OnPropertyChanged(nameof(CanApplyDnsPreset));
                OnPropertyChanged(nameof(CanEditSelectedDnsPreset));
                OnPropertyChanged(nameof(CanDeleteSelectedDnsPreset));
                OnPropertyChanged(nameof(CanEditDnsFields));
                OnPropertyChanged(nameof(SelectedDnsPresetDescriptionText));
                OnPropertyChanged(nameof(DnsDohTemplateVisibility));
            }
        }
    }

    public bool CanApplyDnsPreset =>
        SelectedDnsAdapter is not null &&
        (SelectedDnsPreset?.IsAutomaticPreset == true ||
         (!string.IsNullOrWhiteSpace(PrimaryDnsServer) &&
          (!DnsUseDoh || !string.IsNullOrWhiteSpace(DnsDohTemplate)))) &&
        !IsDnsBusy;

    public bool CanEditSelectedDnsPreset =>
        SelectedDnsPreset?.IsCustom == true && !IsDnsBusy;

    public bool CanDeleteSelectedDnsPreset =>
        SelectedDnsPreset?.IsCustom == true && !IsDnsBusy;

    public bool CanEditDnsFields => !IsDnsBusy;

    public string SelectedDnsPresetDescriptionText =>
        SelectedDnsPreset is null ? string.Empty : SelectedDnsPreset.Description;

    public Visibility DnsDohTemplateVisibility =>
        DnsUseDoh ? Visibility.Visible : Visibility.Collapsed;

    public bool CanResetDnsAutomatic => SelectedDnsAdapter is not null && !IsDnsBusy;

    public bool CanRestorePreviousDns =>
        SelectedDnsAdapter?.HasSavedBackup == true && !IsDnsBusy;

    public bool CanRefreshDnsAdapters => !IsDnsBusy;

    public string SelectedDnsAdapterDescriptionText =>
        SelectedDnsAdapter is null
            ? Strings["DnsNoAdapterSelected"]
            : string.IsNullOrWhiteSpace(SelectedDnsAdapter.Description)
                ? SelectedDnsAdapter.DisplayName
                : $"{SelectedDnsAdapter.DisplayName}{Environment.NewLine}{SelectedDnsAdapter.Description}";

    public string SelectedDnsAdapterServersText =>
        SelectedDnsAdapter is null
            ? Strings["DnsCurrentServersNone"]
            : SelectedDnsAdapter.IsAutomatic
                ? SelectedDnsAdapter.DnsServers.Count == 0
                    ? Strings["DnsAutomaticMode"]
                    : $"{Strings["DnsAutomaticMode"]}: {SelectedDnsAdapter.DnsServersText}"
                : SelectedDnsAdapter.DnsServersText;

    public bool IsManualLaunchMode => !NewSessionUseFullAuto;

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetField(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(IsRefreshAvailable));
            }
        }
    }

    public bool IsRefreshAvailable => !IsLoading;

    public string LastUpdatedText
    {
        get => _lastUpdatedText;
        private set => SetField(ref _lastUpdatedText, value);
    }

    public string FavoriteButtonText =>
        SelectedSession?.IsFavorite == true ? Strings["RemoveFavorite"] : Strings["AddFavorite"];

    public bool IsSetupBusy
    {
        get => _isSetupBusy;
        private set
        {
            if (SetField(ref _isSetupBusy, value))
            {
                OnPropertyChanged(nameof(CanInstallLocalAiTools));
                OnPropertyChanged(nameof(CanInstallLocalAiModels));
                OnPropertyChanged(nameof(CanManageCreativeAiTools));
                OnPropertyChanged(nameof(CanManageAiAgents));
                OnPropertyChanged(nameof(CanUninstallOllama));
                OnPropertyChanged(nameof(CanUninstallLmStudio));
            }
        }
    }

    public bool IsSetupCoreSectionExpanded
    {
        get => _isSetupCoreSectionExpanded;
        set
        {
            if (SetField(ref _isSetupCoreSectionExpanded, value))
            {
                OnPropertyChanged(nameof(SetupCoreSectionContentVisibility));
                OnPropertyChanged(nameof(SetupCoreSectionToggleGlyph));
            }
        }
    }

    public bool IsSetupCodexSectionExpanded
    {
        get => _isSetupCodexSectionExpanded;
        set
        {
            if (SetField(ref _isSetupCodexSectionExpanded, value))
            {
                OnPropertyChanged(nameof(SetupCodexSectionContentVisibility));
                OnPropertyChanged(nameof(SetupCodexSectionToggleGlyph));
            }
        }
    }

    public bool IsSetupLocalAiSectionExpanded
    {
        get => _isSetupLocalAiSectionExpanded;
        set
        {
            if (SetField(ref _isSetupLocalAiSectionExpanded, value))
            {
                OnPropertyChanged(nameof(SetupLocalAiSectionContentVisibility));
                OnPropertyChanged(nameof(SetupLocalAiSectionToggleGlyph));
            }
        }
    }

    public bool IsSetupDnsSectionExpanded
    {
        get => _isSetupDnsSectionExpanded;
        set
        {
            if (SetField(ref _isSetupDnsSectionExpanded, value))
            {
                OnPropertyChanged(nameof(SetupDnsSectionContentVisibility));
                OnPropertyChanged(nameof(SetupDnsSectionToggleGlyph));
            }
        }
    }

    public Visibility SetupCoreSectionContentVisibility =>
        IsSetupCoreSectionExpanded ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SetupCodexSectionContentVisibility =>
        IsSetupCodexSectionExpanded ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SetupLocalAiSectionContentVisibility =>
        IsSetupLocalAiSectionExpanded ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SetupDnsSectionContentVisibility =>
        IsSetupDnsSectionExpanded ? Visibility.Visible : Visibility.Collapsed;

    public string SetupCoreSectionToggleGlyph => IsSetupCoreSectionExpanded ? "−" : "+";

    public string SetupCodexSectionToggleGlyph => IsSetupCodexSectionExpanded ? "−" : "+";

    public string SetupLocalAiSectionToggleGlyph => IsSetupLocalAiSectionExpanded ? "−" : "+";

    public string SetupDnsSectionToggleGlyph => IsSetupDnsSectionExpanded ? "−" : "+";

    public bool IsUpdateBusy
    {
        get => _isUpdateBusy;
        private set
        {
            if (SetField(ref _isUpdateBusy, value))
            {
                RefreshUpdateCommandStates();
            }
        }
    }

    public string CurrentAppVersionText =>
        _lastAppUpdateSnapshot?.CurrentVersionDisplay ?? _updateService.CurrentVersionDisplay;

    public string LatestAppVersionText =>
        _lastAppUpdateSnapshot?.LatestVersionDisplay ?? Strings["UpdateVersionUnknown"];

    public string UpdateReleaseTitleText =>
        string.IsNullOrWhiteSpace(_lastAppUpdateSnapshot?.ReleaseTitle)
            ? Strings["UpdateReleaseUnknown"]
            : _lastAppUpdateSnapshot!.ReleaseTitle;

    public string UpdatePublishedText =>
        _lastAppUpdateSnapshot?.PublishedAtUtc is { } publishedAtUtc
            ? publishedAtUtc.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss")
            : Strings["UpdatePublishedUnknown"];

    public bool CanCheckForUpdates => !IsUpdateBusy;

    public bool CanDownloadUpdate =>
        _lastAppUpdateSnapshot?.IsUpdateAvailable == true &&
        _lastAppUpdateSnapshot.HasInstallerAsset &&
        !IsUpdateBusy;

    public bool CanOpenReleasePage =>
        !IsUpdateBusy &&
        (!string.IsNullOrWhiteSpace(_lastAppUpdateSnapshot?.ReleasePageUrl) ||
         !string.IsNullOrWhiteSpace(_updateService.ReleasePageUrl));

    public string UpdateStatusText
    {
        get => _updateStatusText;
        private set => SetField(ref _updateStatusText, value);
    }

    public string UpdateStatusForeground
    {
        get => _updateStatusForeground;
        private set => SetField(ref _updateStatusForeground, value);
    }

    public bool SettingsDangerousFullAccess
    {
        get => _settingsDangerousFullAccess;
        set => SetField(ref _settingsDangerousFullAccess, value);
    }

    public string SettingsStatusText
    {
        get => _settingsStatusText;
        private set => SetField(ref _settingsStatusText, value);
    }

    public string SettingsStatusForeground
    {
        get => _settingsStatusForeground;
        private set => SetField(ref _settingsStatusForeground, value);
    }

    public string NewSessionPrompt
    {
        get => _newSessionPrompt;
        set
        {
            if (SetField(ref _newSessionPrompt, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
            }
        }
    }

    public string NewSessionWorkingDirectory
    {
        get => _newSessionWorkingDirectory;
        set
        {
            if (SetField(ref _newSessionWorkingDirectory, value))
            {
                OnPropertyChanged(nameof(CanLaunchNewSession));
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
            }
        }
    }

    public string NewSessionModel
    {
        get => _newSessionModel;
        set
        {
            if (SetField(ref _newSessionModel, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
                OnPropertyChanged(nameof(NewSessionModelHelpText));
            }
        }
    }

    public string NewSessionProfile
    {
        get => _newSessionProfile;
        set
        {
            if (SetField(ref _newSessionProfile, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
                OnPropertyChanged(nameof(NewSessionProfileHelpText));
            }
        }
    }

    public string SelectedSandboxMode
    {
        get => _selectedSandboxMode;
        set
        {
            if (SetField(ref _selectedSandboxMode, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
                OnPropertyChanged(nameof(NewSessionSandboxHelpText));
            }
        }
    }

    public string SelectedApprovalPolicy
    {
        get => _selectedApprovalPolicy;
        set
        {
            if (SetField(ref _selectedApprovalPolicy, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
                OnPropertyChanged(nameof(NewSessionApprovalHelpText));
            }
        }
    }

    public string SelectedLocalProvider
    {
        get => _selectedLocalProvider;
        set
        {
            if (SetField(ref _selectedLocalProvider, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
                OnPropertyChanged(nameof(NewSessionLocalProviderHelpText));
            }
        }
    }

    public bool NewSessionUseSearch
    {
        get => _newSessionUseSearch;
        set
        {
            if (SetField(ref _newSessionUseSearch, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
            }
        }
    }

    public bool NewSessionUseOss
    {
        get => _newSessionUseOss;
        set
        {
            if (SetField(ref _newSessionUseOss, value))
            {
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
            }
        }
    }

    public bool NewSessionUseFullAuto
    {
        get => _newSessionUseFullAuto;
        set
        {
            if (SetField(ref _newSessionUseFullAuto, value))
            {
                OnPropertyChanged(nameof(IsManualLaunchMode));
                OnPropertyChanged(nameof(NewSessionPreviewCommandText));
            }
        }
    }

    public string NewSessionPreviewCommandText =>
        _environmentService.BuildInteractiveCommandPreview(BuildNewSessionLaunchOptions());

    public string NewSessionPromptHelpText => Strings["NewSessionPromptHelp"];

    public string NewSessionWorkingDirectoryHelpText => Strings["NewSessionWorkingDirectoryHelp"];

    public string NewSessionModelHelpText =>
        string.IsNullOrWhiteSpace(_configuredCodexModel)
            ? Strings["NewSessionModelHelp"]
            : Strings.Format("NewSessionModelHelpConfigured", _configuredCodexModel);

    public string NewSessionProfileHelpText =>
        ProfileSuggestions.Count == 0
            ? Strings["NewSessionProfileHelp"]
            : Strings.Format("NewSessionProfileHelpConfigured", string.Join(", ", ProfileSuggestions));

    public string NewSessionSandboxHelpText =>
        SandboxModeOptions.FirstOrDefault(option => option.Value == SelectedSandboxMode)?.Description ??
        Strings["NewSessionSandboxHelp"];

    public string NewSessionApprovalHelpText =>
        ApprovalPolicyOptions.FirstOrDefault(option => option.Value == SelectedApprovalPolicy)?.Description ??
        Strings["NewSessionApprovalHelp"];

    public string NewSessionLocalProviderHelpText =>
        LocalProviderOptions.FirstOrDefault(option => option.Value == SelectedLocalProvider)?.Description ??
        Strings["NewSessionLocalProviderHelp"];

    public string NewSessionFlagsHelpText => Strings["NewSessionFlagsHelp"];

    public string NewSessionPreviewHelpText => Strings["NewSessionPreviewHelp"];

    public string NewSessionStatusText
    {
        get => _newSessionStatusText;
        private set => SetField(ref _newSessionStatusText, value);
    }

    public string NewSessionStatusForeground
    {
        get => _newSessionStatusForeground;
        private set => SetField(ref _newSessionStatusForeground, value);
    }

    public string SetupStatusText
    {
        get => _setupStatusText;
        private set => SetField(ref _setupStatusText, value);
    }

    public string SetupStatusForeground
    {
        get => _setupStatusForeground;
        private set => SetField(ref _setupStatusForeground, value);
    }

    public string DnsStatusText
    {
        get => _dnsStatusText;
        private set => SetField(ref _dnsStatusText, value);
    }

    public string DnsStatusForeground
    {
        get => _dnsStatusForeground;
        private set => SetField(ref _dnsStatusForeground, value);
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetField(ref _searchText, value))
            {
                ApplyFilter();
            }
        }
    }

    public LanguageOption? SelectedLanguageOption
    {
        get => _selectedLanguageOption;
        set
        {
            if (value is null)
            {
                return;
            }

            if (SetField(ref _selectedLanguageOption, value))
            {
                ApplyLanguageChange(value.Language);
            }
        }
    }

    public string SelectedSessionNote
    {
        get => _selectedSessionNote;
        set
        {
            if (SetField(ref _selectedSessionNote, value))
            {
                OnPropertyChanged(nameof(CanSaveSelectedSessionNote));
                OnPropertyChanged(nameof(CanClearSelectedSessionNote));
            }
        }
    }

    public SessionRecord? SelectedSession
    {
        get => _selectedSession;
        set
        {
            if (!ReferenceEquals(_selectedSession, value))
            {
                PersistSelectedSessionNote(showStatus: false, refreshFilter: false);
            }

            if (SetField(ref _selectedSession, value))
            {
                SelectedSessionNote = value?.Note ?? string.Empty;
                OnPropertyChanged(nameof(CanOpenSelectedFile));
                OnPropertyChanged(nameof(CanDeleteSelectedSession));
                OnPropertyChanged(nameof(CanResumeSelectedSession));
                OnPropertyChanged(nameof(CanUseSelectedSessionDirectory));
                OnPropertyChanged(nameof(CanEditSelectedSessionNote));
                OnPropertyChanged(nameof(CanSaveSelectedSessionNote));
                OnPropertyChanged(nameof(CanClearSelectedSessionNote));
                OnPropertyChanged(nameof(FavoriteButtonText));
                OnPropertyChanged(nameof(SelectedSessionTitleText));
                OnPropertyChanged(nameof(SelectedSessionPreviewText));
                OnPropertyChanged(nameof(SelectedSessionTranscriptText));
                OnPropertyChanged(nameof(SelectedSessionFavoriteText));
            }
        }
    }

    public string SelectedSessionTitleText =>
        SelectedSession?.Title ?? Strings["NoSessionSelected"];

    public string SelectedSessionPreviewText =>
        SelectedSession?.Preview ?? Strings["SelectSessionHint"];

    public string SelectedSessionTranscriptText =>
        SelectedSession?.TranscriptText ?? Strings["NoTranscriptLoaded"];

    public string SelectedSessionFavoriteText =>
        SelectedSession is null ? "-" : SelectedSession.IsFavorite ? Strings["Yes"] : Strings["No"];

    public string StatusForeground
    {
        get => _statusForeground;
        private set => SetField(ref _statusForeground, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetField(ref _statusText, value);
    }

    public int TotalMessages
    {
        get => _totalMessages;
        private set => SetField(ref _totalMessages, value);
    }

    public int TotalSessions
    {
        get => _totalSessions;
        private set => SetField(ref _totalSessions, value);
    }

    public int FavoriteSessions => _allSessions.Count(session => session.IsFavorite);

    public int RegularSessions => _allSessions.Count(session => !session.IsFavorite);

    public int TotalToolCalls
    {
        get => _totalToolCalls;
        private set => SetField(ref _totalToolCalls, value);
    }

    public int UpdatedTodaySessions
    {
        get => _updatedTodaySessions;
        private set => SetField(ref _updatedTodaySessions, value);
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        UpdateRefreshTimer();
        await RefreshSessionsAsync(isAutomaticRefresh: false);
        await RefreshSetupStatusAsync();
        await RefreshDnsAdaptersAsync();
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        FitToWorkArea();
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        PersistSelectedSessionNote(showStatus: false, refreshFilter: false);
        _refreshTimer.Stop();
    }

    private void OpenSelectedFileButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null || !File.Exists(selectedSession.FilePath))
        {
            return;
        }

        OpenExplorerSelect(selectedSession.FilePath);
    }

    private async void DeleteSelectedSessionButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null || !File.Exists(selectedSession.FilePath))
        {
            return;
        }

        var result = MessageBox.Show(
            Strings.Format("DeleteDialogMessage", selectedSession.Title, selectedSession.SessionId),
            Strings["DeleteDialogTitle"],
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            _sessionService.DeleteSession(selectedSession);
            _favoriteSessionIds.Remove(selectedSession.SessionId);
            _sessionNotes.Remove(selectedSession.SessionId);
            _favoritesService.SaveFavorites(_favoriteSessionIds);
            _notesService.SaveNotes(_sessionNotes);
            await RefreshSessionsAsync(isAutomaticRefresh: false);
        }
        catch (Exception exception)
        {
            SetStatus("#FFD6D6", "StatusDeleteFailed", exception.Message);
        }
    }

    private void OpenSessionsFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (!Directory.Exists(_environmentService.SessionsFolder))
        {
            SetStatus("#FFD6D6", "StatusFolderNotFound", _environmentService.SessionsFolder);
            return;
        }

        CodexEnvironmentService.OpenFolder(_environmentService.SessionsFolder);
    }

    private void ResumeSelectedSessionButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null)
        {
            return;
        }

        if (!File.Exists(_environmentService.CodexCommandPath))
        {
            SetStatus("#FFD6D6", "StatusCodexCmdMissing", _environmentService.CodexCommandPath);
            return;
        }

        var workingDirectory = Directory.Exists(selectedSession.WorkingDirectory)
            ? selectedSession.WorkingDirectory
            : Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        Process.Start(
            new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments =
                    $"/k cd /d {QuoteForCommandLine(workingDirectory)} && call {QuoteForCommandLine(_environmentService.CodexCommandPath)} resume {QuoteForCommandLine(selectedSession.SessionId)}",
                WorkingDirectory = workingDirectory,
                UseShellExecute = true
            });
    }

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshSessionsAsync(isAutomaticRefresh: false);
    }

    private async void RefreshTimer_Tick(object? sender, EventArgs e)
    {
        await RefreshSessionsAsync(isAutomaticRefresh: true);
    }

    private void ToggleFavoriteButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null)
        {
            return;
        }

        if (_favoriteSessionIds.Contains(selectedSession.SessionId))
        {
            _favoriteSessionIds.Remove(selectedSession.SessionId);
            selectedSession.IsFavorite = false;
        }
        else
        {
            _favoriteSessionIds.Add(selectedSession.SessionId);
            selectedSession.IsFavorite = true;
        }

        _favoritesService.SaveFavorites(_favoriteSessionIds);
        OnPropertyChanged(nameof(FavoriteButtonText));
        OnPropertyChanged(nameof(FavoriteSessions));
        OnPropertyChanged(nameof(RegularSessions));
        OnPropertyChanged(nameof(SessionsTabText));
        OnPropertyChanged(nameof(FavoritesTabText));
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        ApplyFilter(selectedSession.SessionId);
    }

    private void SessionsTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedSessionListTab = SessionListTab.Sessions;
    }

    private void FavoritesTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedSessionListTab = SessionListTab.Favorites;
    }

    private void SessionsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Sessions;
    }

    private void NewSessionSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.NewSession;
    }

    private void SetupSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Setup;
    }

    private void ToggleSetupCoreSectionButton_Click(object sender, RoutedEventArgs e)
    {
        IsSetupCoreSectionExpanded = !IsSetupCoreSectionExpanded;
    }

    private void ToggleSetupCodexSectionButton_Click(object sender, RoutedEventArgs e)
    {
        IsSetupCodexSectionExpanded = !IsSetupCodexSectionExpanded;
    }

    private void ToggleSetupLocalAiSectionButton_Click(object sender, RoutedEventArgs e)
    {
        IsSetupLocalAiSectionExpanded = !IsSetupLocalAiSectionExpanded;
    }

    private void ToggleSetupDnsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        IsSetupDnsSectionExpanded = !IsSetupDnsSectionExpanded;
    }

    private void SettingsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Settings;
    }

    private void UseSelectedSessionDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanUseSelectedSessionDirectory || SelectedSession is null)
        {
            return;
        }

        NewSessionWorkingDirectory = SelectedSession.WorkingDirectory;
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionStatusDirectoryCopied"]);
    }

    private void LaunchNewSessionButton_Click(object sender, RoutedEventArgs e)
    {
        var workingDirectory = GetNormalizedNewSessionWorkingDirectory();

        if (!Directory.Exists(workingDirectory))
        {
            SetNewSessionStatus("#FFD6D6", Strings.Format("NewSessionStatusDirectoryMissing", workingDirectory));
            return;
        }

        if (!File.Exists(_environmentService.CodexCommandPath))
        {
            SetNewSessionStatus("#FFD6D6", Strings.Format("StatusCodexCmdMissing", _environmentService.CodexCommandPath));
            return;
        }

        try
        {
            _environmentService.LaunchInteractiveSession(BuildNewSessionLaunchOptions());
            SetNewSessionStatus("#F8E7D6", Strings["NewSessionStatusStarted"]);
        }
        catch (Exception exception)
        {
            SetNewSessionStatus("#FFD6D6", Strings.Format("NewSessionStatusLaunchFailed", exception.Message));
        }
    }

    private async void RefreshSetupStatusButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshSetupStatusAsync();
        await RefreshDnsAdaptersAsync(preserveStatus: false);
    }

    private async void CheckForUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshUpdateStatusAsync();
    }

    private async void DownloadUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        var snapshot = _lastAppUpdateSnapshot;

        if (snapshot is null)
        {
            await RefreshUpdateStatusAsync();
            snapshot = _lastAppUpdateSnapshot;
        }

        if (snapshot is null || !snapshot.IsUpdateAvailable || !snapshot.HasInstallerAsset)
        {
            return;
        }

        var latestVersion = snapshot.LatestVersionDisplay.TrimStart('v', 'V');
        var initialDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Downloads");

        if (!Directory.Exists(initialDirectory))
        {
            initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*",
            DefaultExt = ".exe",
            AddExtension = true,
            FileName = string.IsNullOrWhiteSpace(latestVersion)
                ? "AIHelper-Setup.exe"
                : $"AIHelper-Setup-{latestVersion}.exe",
            InitialDirectory = initialDirectory,
            OverwritePrompt = true
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            IsUpdateBusy = true;
            SetUpdateStatus("#F8E7D6", "UpdateStatusDownloading");
            await _updateService.DownloadInstallerAsync(snapshot.InstallerDownloadUrl, dialog.FileName);
            SetUpdateStatus("#F8E7D6", "UpdateStatusDownloaded", dialog.FileName);

            var launchResult = MessageBox.Show(
                Strings.Format("UpdateLaunchInstallerMessage", dialog.FileName),
                Strings["UpdateLaunchInstallerTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (launchResult != MessageBoxResult.Yes)
            {
                return;
            }

            Process.Start(
                new ProcessStartInfo
                {
                    FileName = dialog.FileName,
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(dialog.FileName) ?? string.Empty
                });
            SetUpdateStatus("#F8E7D6", "UpdateStatusInstallerStarted");
        }
        catch (Exception exception)
        {
            SetUpdateStatus("#FFD6D6", "UpdateStatusDownloadFailed", exception.Message);
        }
        finally
        {
            IsUpdateBusy = false;
            RefreshUpdateCommandStates();
        }
    }

    private void OpenReleasePageButton_Click(object sender, RoutedEventArgs e)
    {
        var releasePageUrl = _lastAppUpdateSnapshot?.ReleasePageUrl;

        if (string.IsNullOrWhiteSpace(releasePageUrl))
        {
            releasePageUrl = _updateService.ReleasePageUrl;
        }

        if (string.IsNullOrWhiteSpace(releasePageUrl))
        {
            return;
        }

        try
        {
            Process.Start(
                new ProcessStartInfo
                {
                    FileName = releasePageUrl,
                    UseShellExecute = true
                });
            SetUpdateStatus("#F8E7D6", "UpdateStatusReleaseOpened");
        }
        catch (Exception exception)
        {
            SetUpdateStatus("#FFD6D6", "UpdateStatusOpenFailed", exception.Message);
        }
    }

    private void ApplyDangerousAccessSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        if (SettingsDangerousFullAccess)
        {
            var confirmation = MessageBox.Show(
                Strings["SettingsDangerousAccessWarningMessage"],
                Strings["SettingsDangerousAccessWarningTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                SettingsDangerousFullAccess = false;
                return;
            }
        }

        _settingsService.SaveDefaultDangerousFullAccess(SettingsDangerousFullAccess);
        ApplyDangerousAccessDefaultsToNewSession();
        SetSettingsStatus(
            "#F8E7D6",
            SettingsDangerousFullAccess ? "SettingsStatusDangerousEnabled" : "SettingsStatusDangerousDisabled");
    }

    private async void RefreshDnsAdaptersButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshDnsAdaptersAsync(preserveStatus: false);
    }

    private void AddDnsPresetButton_Click(object sender, RoutedEventArgs e)
    {
        var editor = new DnsPresetEditorWindow(Strings)
        {
            Owner = this
        };

        if (editor.ShowDialog() != true || editor.ResultPreset is null)
        {
            return;
        }

        if (HasDuplicateDnsPresetName(editor.ResultPreset.Name))
        {
            SetDnsStatus("#FFD6D6", Strings["DnsPresetDuplicateName"]);
            return;
        }

        DnsPresets.Add(editor.ResultPreset);
        SaveCustomDnsPresets();
        SelectedDnsPreset = DnsPresets.FirstOrDefault(
            preset => preset.IsCustom &&
                      string.Equals(preset.Name, editor.ResultPreset.Name, StringComparison.OrdinalIgnoreCase));
        SetDnsStatus("#F8E7D6", Strings["DnsPresetSaved"]);
    }

    private void DuplicateDnsPresetButton_Click(object sender, RoutedEventArgs e)
    {
        var preset = SelectedDnsPreset;

        if (preset is null)
        {
            return;
        }

        var duplicate = preset.Clone();
        duplicate.IsCustom = true;
        duplicate.IsAutomaticPreset = false;
        duplicate.Name = BuildUniqueDnsPresetName(preset.Name);

        var editor = new DnsPresetEditorWindow(Strings, duplicate)
        {
            Owner = this
        };

        if (editor.ShowDialog() != true || editor.ResultPreset is null)
        {
            return;
        }

        if (HasDuplicateDnsPresetName(editor.ResultPreset.Name))
        {
            SetDnsStatus("#FFD6D6", Strings["DnsPresetDuplicateName"]);
            return;
        }

        DnsPresets.Add(editor.ResultPreset);
        SaveCustomDnsPresets();
        SelectedDnsPreset = DnsPresets.FirstOrDefault(
            item => item.IsCustom &&
                    string.Equals(item.Name, editor.ResultPreset.Name, StringComparison.OrdinalIgnoreCase));
        SetDnsStatus("#F8E7D6", Strings["DnsPresetSaved"]);
    }

    private void EditDnsPresetButton_Click(object sender, RoutedEventArgs e)
    {
        var preset = SelectedDnsPreset;

        if (preset is null || !preset.IsCustom)
        {
            return;
        }

        var editor = new DnsPresetEditorWindow(Strings, preset.Clone())
        {
            Owner = this
        };

        if (editor.ShowDialog() != true || editor.ResultPreset is null)
        {
            return;
        }

        if (HasDuplicateDnsPresetName(editor.ResultPreset.Name, preset))
        {
            SetDnsStatus("#FFD6D6", Strings["DnsPresetDuplicateName"]);
            return;
        }

        preset.Name = editor.ResultPreset.Name;
        preset.PrimaryDns = editor.ResultPreset.PrimaryDns;
        preset.SecondaryDns = editor.ResultPreset.SecondaryDns;
        preset.Description = editor.ResultPreset.Description;
        preset.EnableDoh = editor.ResultPreset.EnableDoh;
        preset.DohTemplate = editor.ResultPreset.DohTemplate;

        SaveCustomDnsPresets();
        ReplaceDnsPresetCollection();
        SelectedDnsPreset = DnsPresets.FirstOrDefault(
            item => item.IsCustom &&
                    string.Equals(item.Name, preset.Name, StringComparison.OrdinalIgnoreCase));
        SetDnsStatus("#F8E7D6", Strings["DnsPresetUpdated"]);
    }

    private void DeleteDnsPresetButton_Click(object sender, RoutedEventArgs e)
    {
        var preset = SelectedDnsPreset;

        if (preset is null || !preset.IsCustom)
        {
            return;
        }

        var result = MessageBox.Show(
            Strings.Format("DnsDeletePresetMessage", preset.Name),
            Strings["DnsDeletePresetTitle"],
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        DnsPresets.Remove(preset);
        SaveCustomDnsPresets();
        ReplaceDnsPresetCollection();
        SelectedDnsPreset = DnsPresets.FirstOrDefault();
        SetDnsStatus("#F8E7D6", Strings["DnsPresetDeleted"]);
    }

    private void ImportDnsPresetsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON (*.json)|*.json|All files (*.*)|*.*",
                DefaultExt = ".json",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            var imported = _dnsPresetSettingsService.ImportCustomPresets(dialog.FileName);
            var mergedPresets = MergeImportedDnsPresets(imported);
            ReplaceDnsPresetCollection(
                DnsPresetCatalog.CreateDefaultPresets(Strings)
                    .Concat(mergedPresets)
                    .ToList());
            SaveCustomDnsPresets();
            SelectedDnsPreset = imported.FirstOrDefault() is { } firstImported
                ? DnsPresets.FirstOrDefault(
                    item => item.IsCustom &&
                            string.Equals(item.Name, firstImported.Name, StringComparison.OrdinalIgnoreCase))
                : SelectedDnsPreset;

            SetDnsStatus("#F8E7D6", Strings.Format("DnsPresetImported", imported.Count));
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
    }

    private void ExportDnsPresetsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON (*.json)|*.json|All files (*.*)|*.*",
                DefaultExt = ".json",
                FileName = "dns-presets.json",
                AddExtension = true
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            var customPresets = DnsPresets.Where(preset => preset.IsCustom).ToList();
            _dnsPresetSettingsService.ExportCustomPresets(dialog.FileName, customPresets);
            SetDnsStatus("#F8E7D6", Strings.Format("DnsPresetExported", customPresets.Count));
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
    }

    private async void ApplySelectedDnsPresetButton_Click(object sender, RoutedEventArgs e)
    {
        var adapter = SelectedDnsAdapter;

        if (adapter is null)
        {
            return;
        }

        if (!ConfirmDnsWarning(
                Strings["DnsApplyWarningTitle"],
                Strings["DnsApplyWarningMessage"]))
        {
            return;
        }

        try
        {
            IsDnsBusy = true;
            SetDnsStatus("#F8E7D6", Strings["DnsStatusApplying"]);
            await Task.Run(() => _dnsManagementService.ApplyPreset(adapter, BuildDnsPresetForApply()));
            await RefreshDnsAdaptersAsync(preserveStatus: false);
            SetDnsStatus("#F8E7D6", Strings["DnsStatusApplied"]);
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
        finally
        {
            IsDnsBusy = false;
            RefreshDnsCommandStates();
        }
    }

    private async void ResetAutomaticDnsButton_Click(object sender, RoutedEventArgs e)
    {
        var adapter = SelectedDnsAdapter;

        if (adapter is null)
        {
            return;
        }

        if (!ConfirmDnsWarning(
                Strings["DnsResetWarningTitle"],
                Strings["DnsResetWarningMessage"]))
        {
            return;
        }

        try
        {
            IsDnsBusy = true;
            SetDnsStatus("#F8E7D6", Strings["DnsStatusResetting"]);
            await Task.Run(() => _dnsManagementService.ResetToAutomatic(adapter));
            await RefreshDnsAdaptersAsync(preserveStatus: false);
            SetDnsStatus("#F8E7D6", Strings["DnsStatusReset"]);
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
        finally
        {
            IsDnsBusy = false;
            RefreshDnsCommandStates();
        }
    }

    private async void RestoreDnsBackupButton_Click(object sender, RoutedEventArgs e)
    {
        var adapter = SelectedDnsAdapter;

        if (adapter is null)
        {
            return;
        }

        if (!ConfirmDnsWarning(
                Strings["DnsRestoreWarningTitle"],
                Strings["DnsRestoreWarningMessage"]))
        {
            return;
        }

        try
        {
            IsDnsBusy = true;
            SetDnsStatus("#F8E7D6", Strings["DnsStatusRestoring"]);
            await Task.Run(() => _dnsManagementService.RestoreBackup(adapter));
            await RefreshDnsAdaptersAsync(preserveStatus: false);
            SetDnsStatus("#F8E7D6", Strings["DnsStatusRestored"]);
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
        finally
        {
            IsDnsBusy = false;
            RefreshDnsCommandStates();
        }
    }

    private void InstallCodexStackButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchCodexInstallRepairTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusInstallerStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallOllamaButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOllamaInstallTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOllamaInstallStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallLmStudioButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchLmStudioInstallTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusLmStudioInstallStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void UninstallOllamaButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveRuntimeWarningTitle"],
                Strings.Format("SetupRemoveRuntimeWarningMessage", "Ollama")))
        {
            return;
        }

        try
        {
            _environmentService.LaunchOllamaUninstallTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOllamaUninstallStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void UninstallLmStudioButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveRuntimeWarningTitle"],
                Strings.Format("SetupRemoveRuntimeWarningMessage", "LM Studio")))
        {
            return;
        }

        try
        {
            _environmentService.LaunchLmStudioUninstallTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusLmStudioUninstallStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallLocalAiModelButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: LocalAiModelOption option })
        {
            return;
        }

        if (!_environmentService.IsOllamaInstalled())
        {
            SetSetupStatus("#FFD6D6", Strings["SetupStatusOllamaMissing"]);
            return;
        }

        try
        {
            _environmentService.LaunchOllamaModelInstallTerminal(option.ModelTag);
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusModelInstallStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void RemoveLocalAiModelButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: LocalAiModelOption option })
        {
            return;
        }

        if (!option.IsInstalled)
        {
            return;
        }

        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveModelWarningTitle"],
                Strings.Format("SetupRemoveModelWarningMessage", option.Name, option.ModelTag)))
        {
            return;
        }

        try
        {
            _environmentService.LaunchOllamaModelRemoveTerminal(option.ModelTag);
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusModelRemoveStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallCreativeAiToolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: CreativeAiToolOption option })
        {
            return;
        }

        try
        {
            _environmentService.LaunchCreativeToolInstallTerminal(option.PackageId, option.Name);
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusCreativeToolInstallStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void RemoveCreativeAiToolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: CreativeAiToolOption option })
        {
            return;
        }

        if (!option.IsInstalled || IsSetupBusy)
        {
            return;
        }

        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveCreativeToolWarningTitle"],
                Strings.Format("SetupRemoveCreativeToolWarningMessage", option.Name)))
        {
            return;
        }

        try
        {
            _environmentService.LaunchCreativeToolUninstallTerminal(option.PackageId, option.Name);
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusCreativeToolRemoveStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallAiAgentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: CreativeAiToolOption option })
        {
            return;
        }

        try
        {
            _environmentService.LaunchOpenClawInstallTerminal();
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusAiAgentInstallStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void RemoveAiAgentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: CreativeAiToolOption option })
        {
            return;
        }

        if (!option.IsInstalled || IsSetupBusy)
        {
            return;
        }

        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveAiAgentWarningTitle"],
                Strings.Format("SetupRemoveAiAgentWarningMessage", option.Name)))
        {
            return;
        }

        try
        {
            _environmentService.LaunchOpenClawUninstallTerminal();
            SetSetupStatus("#F8E7D6", Strings.Format("SetupStatusAiAgentRemoveStarted", option.Name));
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void LaunchCodexLoginButton_Click(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(_environmentService.CodexCommandPath))
        {
            SetSetupStatus("#FFD6D6", Strings.Format("StatusCodexCmdMissing", _environmentService.CodexCommandPath));
            return;
        }

        try
        {
            _environmentService.LaunchCodexLoginTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusLoginStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenCodexHomeButton_Click(object sender, RoutedEventArgs e)
    {
        if (!Directory.Exists(_environmentService.CodexHomeFolder))
        {
            SetSetupStatus("#FFD6D6", Strings.Format("StatusFolderNotFound", _environmentService.CodexHomeFolder));
            return;
        }

        CodexEnvironmentService.OpenFolder(_environmentService.CodexHomeFolder);
    }

    private void SaveSelectedNoteButton_Click(object sender, RoutedEventArgs e)
    {
        PersistSelectedSessionNote(showStatus: true, refreshFilter: true);
    }

    private void ClearSelectedNoteButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedSession is null)
        {
            return;
        }

        SelectedSessionNote = string.Empty;
        PersistSelectedSessionNote(showStatus: true, refreshFilter: true);
    }

    private void ApplyFilter(string? preferredSessionId = null)
    {
        var currentId = preferredSessionId ?? SelectedSession?.SessionId;
        var filter = SearchText.Trim();
        var sourceSessions = SelectedSessionListTab == SessionListTab.Favorites
            ? _allSessions.Where(session => session.IsFavorite).ToList()
            : _allSessions.Where(session => !session.IsFavorite).ToList();

        var filtered = string.IsNullOrWhiteSpace(filter)
            ? sourceSessions
            : sourceSessions
                .Where(session => session.SearchBlob.Contains(filter, StringComparison.OrdinalIgnoreCase))
                .ToList();

        filtered = filtered
            .OrderByDescending(session => session.UpdatedAtUtc)
            .ToList();

        Sessions.Clear();

        foreach (var session in filtered)
        {
            Sessions.Add(session);
        }

        if (Sessions.Count == 0)
        {
            SelectedSession = null;
            OnPropertyChanged(nameof(HasVisibleSessions));
            return;
        }

        SelectedSession = Sessions.FirstOrDefault(session => session.SessionId == currentId) ?? Sessions[0];
        OnPropertyChanged(nameof(HasVisibleSessions));
    }

    private void ApplySessions(IReadOnlyList<SessionRecord> refreshedSessions)
    {
        _favoriteSessionIds = _favoritesService.LoadFavorites();
        _sessionNotes = _notesService.LoadNotes();
        _allSessions = refreshedSessions.ToList();

        foreach (var session in _allSessions)
        {
            session.IsFavorite = _favoriteSessionIds.Contains(session.SessionId);
            session.Note = _sessionNotes.TryGetValue(session.SessionId, out var note) ? note : string.Empty;
            UpdateSessionSearchBlob(session);
        }

        TotalSessions = _allSessions.Count;
        UpdatedTodaySessions = _allSessions.Count(
            session => session.UpdatedAtUtc.ToLocalTime().Date == DateTime.Today);
        TotalMessages = _allSessions.Sum(session => session.TotalMessageCount);
        TotalToolCalls = _allSessions.Sum(session => session.ToolCallCount);
        OnPropertyChanged(nameof(FavoriteSessions));
        OnPropertyChanged(nameof(RegularSessions));
        OnPropertyChanged(nameof(SessionsTabText));
        OnPropertyChanged(nameof(FavoritesTabText));
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        ApplyFilter();
    }

    private void ApplyLanguageChange(AppLanguage language)
    {
        if (Strings.CurrentLanguage == language)
        {
            return;
        }

        Strings.SetLanguage(language);
        _settingsService.SaveLanguage(language);

        OnPropertyChanged(nameof(Strings));
        OnPropertyChanged(nameof(FavoriteButtonText));
        OnPropertyChanged(nameof(SessionsTabText));
        OnPropertyChanged(nameof(FavoritesTabText));
        OnPropertyChanged(nameof(EmptySessionsText));
        OnPropertyChanged(nameof(SelectedSessionTitleText));
        OnPropertyChanged(nameof(SelectedSessionPreviewText));
        OnPropertyChanged(nameof(SelectedSessionTranscriptText));
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        OnPropertyChanged(nameof(NewSessionPreviewCommandText));
        OnPropertyChanged(nameof(NewSessionPromptHelpText));
        OnPropertyChanged(nameof(NewSessionWorkingDirectoryHelpText));
        OnPropertyChanged(nameof(NewSessionModelHelpText));
        OnPropertyChanged(nameof(NewSessionProfileHelpText));
        OnPropertyChanged(nameof(NewSessionSandboxHelpText));
        OnPropertyChanged(nameof(NewSessionApprovalHelpText));
        OnPropertyChanged(nameof(NewSessionLocalProviderHelpText));
        OnPropertyChanged(nameof(NewSessionFlagsHelpText));
        OnPropertyChanged(nameof(NewSessionPreviewHelpText));
        OnPropertyChanged(nameof(SelectedDnsAdapterDescriptionText));
        OnPropertyChanged(nameof(SelectedDnsAdapterServersText));
        OnPropertyChanged(nameof(SelectedDnsPresetDescriptionText));
        OnPropertyChanged(nameof(CanApplyDnsPreset));
        OnPropertyChanged(nameof(CanEditSelectedDnsPreset));
        OnPropertyChanged(nameof(CanDeleteSelectedDnsPreset));
        OnPropertyChanged(nameof(CanEditDnsFields));
        OnPropertyChanged(nameof(CurrentAppVersionText));
        OnPropertyChanged(nameof(LatestAppVersionText));
        OnPropertyChanged(nameof(UpdateReleaseTitleText));
        OnPropertyChanged(nameof(UpdatePublishedText));
        OnPropertyChanged(nameof(CanCheckForUpdates));
        OnPropertyChanged(nameof(CanDownloadUpdate));
        OnPropertyChanged(nameof(CanOpenReleasePage));
        OnPropertyChanged(nameof(CanInstallLocalAiTools));
        OnPropertyChanged(nameof(CanInstallLocalAiModels));
        OnPropertyChanged(nameof(CanManageCreativeAiTools));
        OnPropertyChanged(nameof(CanManageAiAgents));

        RefreshLaunchOptionCollections();
        RefreshLocalAiModelOptions();
        RefreshCreativeAiToolOptions(_lastEnvironmentSnapshot);
        RefreshAiAgentToolOptions(_lastEnvironmentSnapshot);
        LoadDnsPresets(SelectedDnsPreset);
        RefreshLocalizedChromeText();
        RefreshSectionChromeText();

        if (_lastEnvironmentSnapshot is not null)
        {
            ApplySetupSnapshot(_lastEnvironmentSnapshot);
        }

        if (_lastAppUpdateSnapshot is not null)
        {
            ApplyUpdateSnapshot(_lastAppUpdateSnapshot);
        }

        if (IsLoaded)
        {
            _ = RefreshSessionsAsync(isAutomaticRefresh: false);
        }
    }

    private void RefreshLocalizedChromeText()
    {
        StatusText = FormatLocalizedText(_statusKey, _statusArgs);
        SettingsStatusText = FormatLocalizedText(_settingsStatusKey, _settingsStatusArgs);
        UpdateStatusText = FormatLocalizedText(_updateStatusKey, _updateStatusArgs);
        LastUpdatedText = _lastUpdatedAtLocal is null
            ? Strings["NoRefreshYet"]
            : Strings.Format("LastUpdated", _lastUpdatedAtLocal.Value.ToString("dd.MM.yyyy HH:mm:ss"));
    }

    private void RefreshSectionChromeText()
    {
        NewSessionStatusText = Strings["NewSessionStatusReady"];
        NewSessionStatusForeground = "#F8E7D6";
        SetupStatusText = Strings["SetupStatusReady"];
        SetupStatusForeground = "#F8E7D6";
        _settingsStatusKey = "SettingsStatusReady";
        _settingsStatusArgs = [];
        SettingsStatusText = Strings["SettingsStatusReady"];
        SettingsStatusForeground = "#F8E7D6";
        _updateStatusKey = "UpdateStatusReady";
        _updateStatusArgs = [];
        UpdateStatusText = Strings["UpdateStatusReady"];
        UpdateStatusForeground = "#F8E7D6";
        DnsStatusText = Strings["DnsStatusReady"];
        DnsStatusForeground = "#F8E7D6";
    }

    private void RefreshLaunchOptionCollections()
    {
        ReplaceLaunchOptions(
            SandboxModeOptions,
            [
                new LaunchOption
                {
                    Value = "workspace-write",
                    DisplayName = Strings["NewSessionSandboxWorkspace"],
                    Description = Strings["NewSessionSandboxWorkspaceHelp"]
                },
                new LaunchOption
                {
                    Value = "read-only",
                    DisplayName = Strings["NewSessionSandboxReadonly"],
                    Description = Strings["NewSessionSandboxReadonlyHelp"]
                },
                new LaunchOption
                {
                    Value = "danger-full-access",
                    DisplayName = Strings["NewSessionSandboxDanger"],
                    Description = Strings["NewSessionSandboxDangerHelp"]
                }
            ]);

        ReplaceLaunchOptions(
            ApprovalPolicyOptions,
            [
                new LaunchOption
                {
                    Value = "on-request",
                    DisplayName = Strings["NewSessionApprovalOnRequest"],
                    Description = Strings["NewSessionApprovalOnRequestHelp"]
                },
                new LaunchOption
                {
                    Value = "never",
                    DisplayName = Strings["NewSessionApprovalNever"],
                    Description = Strings["NewSessionApprovalNeverHelp"]
                },
                new LaunchOption
                {
                    Value = "untrusted",
                    DisplayName = Strings["NewSessionApprovalUntrusted"],
                    Description = Strings["NewSessionApprovalUntrustedHelp"]
                }
            ]);

        ReplaceLaunchOptions(
            LocalProviderOptions,
            [
                new LaunchOption
                {
                    Value = string.Empty,
                    DisplayName = Strings["NewSessionLocalProviderNone"],
                    Description = Strings["NewSessionLocalProviderNoneHelp"]
                },
                new LaunchOption
                {
                    Value = "lmstudio",
                    DisplayName = "LM Studio",
                    Description = Strings["NewSessionLocalProviderLmStudioHelp"]
                },
                new LaunchOption
                {
                    Value = "ollama",
                    DisplayName = "Ollama",
                    Description = Strings["NewSessionLocalProviderOllamaHelp"]
                }
            ]);
    }

    private void RefreshLocalAiModelOptions(IReadOnlyDictionary<string, string>? installedModels = null)
    {
        var localModels = installedModels ??
                          _lastEnvironmentSnapshot?.InstalledOllamaModels ??
                          new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        LocalAiModelOptions.Clear();

        LocalAiModelOptions.Add(
            new LocalAiModelOption
            {
                Name = "Qwen2.5 Coder 7B",
                ModelTag = "qwen2.5-coder:7b",
                SizeLabel = "7B",
                Description = Strings["SetupLocalAiModelCoderDescription"],
                IsInstalled = localModels.TryGetValue("qwen2.5-coder:7b", out var qwenSize),
                InstalledStatusText = localModels.ContainsKey("qwen2.5-coder:7b")
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = localModels.ContainsKey("qwen2.5-coder:7b") ? "#1F7A52" : "#5E6C76",
                InstalledSizeText = localModels.TryGetValue("qwen2.5-coder:7b", out qwenSize)
                    ? Strings.Format("SetupLocalAiInstalledSize", qwenSize)
                    : Strings["SetupLocalAiMissingSize"]
            });
        LocalAiModelOptions.Add(
            new LocalAiModelOption
            {
                Name = "Phi-4 Mini",
                ModelTag = "phi4-mini",
                SizeLabel = "3.8B",
                Description = Strings["SetupLocalAiModelReasoningDescription"],
                IsInstalled = localModels.TryGetValue("phi4-mini", out var phiSize),
                InstalledStatusText = localModels.ContainsKey("phi4-mini")
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = localModels.ContainsKey("phi4-mini") ? "#1F7A52" : "#5E6C76",
                InstalledSizeText = localModels.TryGetValue("phi4-mini", out phiSize)
                    ? Strings.Format("SetupLocalAiInstalledSize", phiSize)
                    : Strings["SetupLocalAiMissingSize"]
            });
        LocalAiModelOptions.Add(
            new LocalAiModelOption
            {
                Name = "Gemma 3 4B",
                ModelTag = "gemma3:4b",
                SizeLabel = "4B",
                Description = Strings["SetupLocalAiModelGeneralDescription"],
                IsInstalled = localModels.TryGetValue("gemma3:4b", out var gemmaSize),
                InstalledStatusText = localModels.ContainsKey("gemma3:4b")
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = localModels.ContainsKey("gemma3:4b") ? "#1F7A52" : "#5E6C76",
                InstalledSizeText = localModels.TryGetValue("gemma3:4b", out gemmaSize)
                    ? Strings.Format("SetupLocalAiInstalledSize", gemmaSize)
                    : Strings["SetupLocalAiMissingSize"]
            });
    }

    private void RefreshCreativeAiToolOptions(CodexEnvironmentSnapshot? snapshot = null)
    {
        var environment = snapshot ?? _lastEnvironmentSnapshot;
        var comfyInstalled = environment?.ComfyUiDesktopAvailable == true;
        var pinokioInstalled = environment?.PinokioAvailable == true;

        CreativeAiToolOptions.Clear();
        CreativeAiToolOptions.Add(
            new CreativeAiToolOption
            {
                Name = "ComfyUI Desktop",
                PackageId = "Comfy.ComfyUI-Desktop",
                CoverageLabel = Strings["SetupCreativeAiCoverageUniversal"],
                Description = Strings["SetupCreativeAiToolComfyDescription"],
                IsInstalled = comfyInstalled,
                InstalledStatusText = comfyInstalled
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = comfyInstalled ? "#1F7A52" : "#5E6C76",
                InstalledDetailText = comfyInstalled
                    ? environment?.ComfyUiDesktopDetail ?? "Comfy.ComfyUI-Desktop"
                    : Strings["SetupCreativeAiMissingDetail"]
            });
        CreativeAiToolOptions.Add(
            new CreativeAiToolOption
            {
                Name = "Pinokio",
                PackageId = "pinokiocomputer.pinokio",
                CoverageLabel = Strings["SetupCreativeAiCoverageLauncher"],
                Description = Strings["SetupCreativeAiToolPinokioDescription"],
                IsInstalled = pinokioInstalled,
                InstalledStatusText = pinokioInstalled
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = pinokioInstalled ? "#1F7A52" : "#5E6C76",
                InstalledDetailText = pinokioInstalled
                    ? environment?.PinokioDetail ?? "pinokiocomputer.pinokio"
                    : Strings["SetupCreativeAiMissingDetail"]
            });
    }

    private void RefreshAiAgentToolOptions(CodexEnvironmentSnapshot? snapshot = null)
    {
        var environment = snapshot ?? _lastEnvironmentSnapshot;
        var openClawInstalled = environment?.OpenClawAvailable == true;

        AiAgentToolOptions.Clear();
        AiAgentToolOptions.Add(
            new CreativeAiToolOption
            {
                Name = "OpenClaw",
                PackageId = "openclaw",
                CoverageLabel = Strings["SetupAiAgentCoverageLocal"],
                Description = Strings["SetupAiAgentOpenClawDescription"],
                IsInstalled = openClawInstalled,
                InstalledStatusText = openClawInstalled
                    ? Strings["SetupLocalAiInstalled"]
                    : Strings["SetupLocalAiNotInstalled"],
                InstalledStatusBrush = openClawInstalled ? "#1F7A52" : "#5E6C76",
                InstalledDetailText = openClawInstalled
                    ? environment?.OpenClawDetail ?? "openclaw"
                    : Strings["SetupAiAgentMissingDetail"]
            });
    }

    private void LoadNewSessionConfigurationInfo()
    {
        var configInfo = _environmentService.GetCodexConfigInfo();
        var previousConfiguredModel = _configuredCodexModel;
        _configuredCodexModel = configInfo.DefaultModel;

        ReplaceStringCollection(ModelSuggestions, configInfo.AvailableModels);
        ReplaceStringCollection(ProfileSuggestions, configInfo.Profiles);

        if (string.IsNullOrWhiteSpace(NewSessionModel) ||
            string.Equals(NewSessionModel, previousConfiguredModel, StringComparison.OrdinalIgnoreCase))
        {
            NewSessionModel = configInfo.DefaultModel;
        }

        if (string.IsNullOrWhiteSpace(NewSessionProfile) && ProfileSuggestions.Count == 1)
        {
            NewSessionProfile = ProfileSuggestions[0];
        }

        OnPropertyChanged(nameof(NewSessionModelHelpText));
        OnPropertyChanged(nameof(NewSessionProfileHelpText));
    }

    private void LoadNewSessionConfigurationInfoSafe()
    {
        try
        {
            LoadNewSessionConfigurationInfo();
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to load Codex config for the New Session page.", exception);
            LoadFallbackNewSessionConfigurationInfo();
        }
    }

    private void LoadFallbackNewSessionConfigurationInfo()
    {
        _configuredCodexModel = string.Empty;

        ReplaceStringCollection(
            ModelSuggestions,
            [
                "gpt-5.4",
                "gpt-5.4-mini",
                "gpt-5.3-codex",
                "gpt-5.3-codex-spark",
                "gpt-5.2"
            ]);
        ReplaceStringCollection(ProfileSuggestions, []);

        if (string.IsNullOrWhiteSpace(NewSessionModel))
        {
            NewSessionModel = string.Empty;
        }

        OnPropertyChanged(nameof(NewSessionModelHelpText));
        OnPropertyChanged(nameof(NewSessionProfileHelpText));
    }

    private void ApplyDangerousAccessDefaultsToNewSession()
    {
        if (_isApplyingDangerousAccessDefaults)
        {
            return;
        }

        _isApplyingDangerousAccessDefaults = true;

        try
        {
            if (SettingsDangerousFullAccess)
            {
                NewSessionUseFullAuto = false;
                SelectedSandboxMode = "danger-full-access";
                SelectedApprovalPolicy = "never";
            }
            else
            {
                if (string.Equals(SelectedSandboxMode, "danger-full-access", StringComparison.OrdinalIgnoreCase))
                {
                    SelectedSandboxMode = "workspace-write";
                }

                if (string.Equals(SelectedApprovalPolicy, "never", StringComparison.OrdinalIgnoreCase))
                {
                    SelectedApprovalPolicy = "on-request";
                }
            }
        }
        finally
        {
            _isApplyingDangerousAccessDefaults = false;
        }

        OnPropertyChanged(nameof(NewSessionPreviewCommandText));
    }

    private void LoadDnsPresets(DnsPreset? preferredPreset = null)
    {
        var presets = _dnsPresetSettingsService.LoadAllPresets(Strings);
        ReplaceDnsPresetCollection(presets);

        if (preferredPreset is not null)
        {
            SelectedDnsPreset =
                DnsPresets.FirstOrDefault(
                    preset => preset.IsCustom == preferredPreset.IsCustom &&
                              string.Equals(preset.Name, preferredPreset.Name, StringComparison.OrdinalIgnoreCase)) ??
                FindEquivalentBuiltInPreset(preferredPreset) ??
                DnsPresets.FirstOrDefault();
        }
        else
        {
            SelectedDnsPreset = DnsPresets.FirstOrDefault(
                                    preset => string.Equals(
                                        preset.Name,
                                        Strings["DnsPresetAutomatic"],
                                        StringComparison.Ordinal)) ??
                                DnsPresets.FirstOrDefault();
        }
    }

    private void LoadDnsPresetsSafe()
    {
        try
        {
            LoadDnsPresets();
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to load DNS presets.", exception);
            ReplaceDnsPresetCollection(DnsPresetCatalog.CreateDefaultPresets(Strings));
            SelectedDnsPreset = DnsPresets.FirstOrDefault();
        }
    }

    private void ReplaceDnsPresetCollection()
    {
        var currentSelection = SelectedDnsPreset;
        ReplaceDnsPresetCollection(DnsPresets.ToList());
        SelectedDnsPreset = currentSelection is null
            ? DnsPresets.FirstOrDefault()
            : DnsPresets.FirstOrDefault(
                    preset => preset.IsCustom == currentSelection.IsCustom &&
                              string.Equals(preset.Name, currentSelection.Name, StringComparison.OrdinalIgnoreCase)) ??
                FindEquivalentBuiltInPreset(currentSelection) ??
                DnsPresets.FirstOrDefault();
    }

    private void ReplaceDnsPresetCollection(IReadOnlyList<DnsPreset> presets)
    {
        DnsPresets.Clear();

        foreach (var preset in presets)
        {
            DnsPresets.Add(preset.Clone());
        }

        OnPropertyChanged(nameof(CanApplyDnsPreset));
        OnPropertyChanged(nameof(CanEditSelectedDnsPreset));
        OnPropertyChanged(nameof(CanDeleteSelectedDnsPreset));
        OnPropertyChanged(nameof(CanEditDnsFields));
    }

    private DnsPreset BuildDnsPresetForApply()
    {
        var selectedPreset = SelectedDnsPreset;

        return new DnsPreset
        {
            Name = selectedPreset?.Name ?? Strings["DnsPresetCustom"],
            PrimaryDns = PrimaryDnsServer.Trim(),
            SecondaryDns = SecondaryDnsServer.Trim(),
            Description = selectedPreset?.Description ?? string.Empty,
            EnableDoh = DnsUseDoh,
            DohTemplate = DnsUseDoh ? DnsDohTemplate.Trim() : string.Empty,
            IsCustom = selectedPreset?.IsCustom ?? false,
            IsAutomaticPreset = selectedPreset?.IsAutomaticPreset == true
        };
    }

    private void ApplyDnsPresetToEditors(DnsPreset? preset)
    {
        if (preset is null)
        {
            return;
        }

        if (preset.IsAutomaticPreset || IsBuiltInCustomPreset(preset))
        {
            PrimaryDnsServer = string.Empty;
            SecondaryDnsServer = string.Empty;
            DnsUseDoh = false;
            DnsDohTemplate = string.Empty;
            return;
        }

        PrimaryDnsServer = preset.PrimaryDns;
        SecondaryDnsServer = preset.SecondaryDns;
        DnsUseDoh = preset.EnableDoh;
        DnsDohTemplate = preset.EnableDoh ? preset.DohTemplate : string.Empty;
    }

    private bool HasDuplicateDnsPresetName(string name, DnsPreset? ignoredPreset = null)
    {
        return DnsPresets.Any(
            preset => !ReferenceEquals(preset, ignoredPreset) &&
                      string.Equals(preset.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    private string BuildUniqueDnsPresetName(string baseName)
    {
        var candidate = $"{baseName} {Strings["DnsPresetCopySuffix"]}";

        if (!HasDuplicateDnsPresetName(candidate))
        {
            return candidate;
        }

        for (var index = 2; index < 1000; index++)
        {
            var numberedCandidate = $"{candidate} {index}";

            if (!HasDuplicateDnsPresetName(numberedCandidate))
            {
                return numberedCandidate;
            }
        }

        return $"{candidate} {DateTime.Now:HHmmss}";
    }

    private void SaveCustomDnsPresets()
    {
        try
        {
            _dnsPresetSettingsService.SaveCustomPresets(DnsPresets.Where(preset => preset.IsCustom));
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
    }

    private bool IsBuiltInCustomPreset(DnsPreset preset)
    {
        return !preset.IsCustom &&
               !preset.IsAutomaticPreset &&
               string.IsNullOrWhiteSpace(preset.PrimaryDns) &&
               string.IsNullOrWhiteSpace(preset.SecondaryDns);
    }

    private DnsPreset? FindEquivalentBuiltInPreset(DnsPreset preset)
    {
        if (preset.IsCustom)
        {
            return null;
        }

        if (preset.IsAutomaticPreset)
        {
            return DnsPresets.FirstOrDefault(
                item => !item.IsCustom &&
                        item.IsAutomaticPreset);
        }

        if (IsBuiltInCustomPreset(preset))
        {
            return DnsPresets.FirstOrDefault(item => IsBuiltInCustomPreset(item));
        }

        return DnsPresets.FirstOrDefault(
            item => !item.IsCustom &&
                    string.Equals(item.PrimaryDns, preset.PrimaryDns, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.SecondaryDns, preset.SecondaryDns, StringComparison.OrdinalIgnoreCase));
    }

    private IReadOnlyList<DnsPreset> MergeImportedDnsPresets(IReadOnlyList<DnsPreset> importedPresets)
    {
        var mergedByName = DnsPresets
            .Where(preset => preset.IsCustom)
            .ToDictionary(preset => preset.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var importedPreset in importedPresets)
        {
            mergedByName[importedPreset.Name] = importedPreset.Clone();
        }

        return mergedByName.Values
            .OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void ReplaceLaunchOptions(
        ObservableCollection<LaunchOption> target,
        IReadOnlyList<LaunchOption> values)
    {
        target.Clear();

        foreach (var value in values)
        {
            target.Add(value);
        }
    }

    private static void ReplaceStringCollection(
        ObservableCollection<string> target,
        IReadOnlyList<string> values)
    {
        target.Clear();

        foreach (var value in values)
        {
            target.Add(value);
        }
    }

    private string FormatLocalizedText(string key, params object[] args)
    {
        return args.Length == 0 ? Strings[key] : Strings.Format(key, args);
    }

    private void SetStatus(string foreground, string key, params object[] args)
    {
        _statusKey = key;
        _statusArgs = args;
        StatusForeground = foreground;
        StatusText = FormatLocalizedText(key, args);
    }

    private void SetNewSessionStatus(string foreground, string text)
    {
        NewSessionStatusForeground = foreground;
        NewSessionStatusText = text;
    }

    private void SetSetupStatus(string foreground, string text)
    {
        SetupStatusForeground = foreground;
        SetupStatusText = text;
    }

    private void SetSettingsStatus(string foreground, string key, params object[] args)
    {
        _settingsStatusKey = key;
        _settingsStatusArgs = args;
        SettingsStatusForeground = foreground;
        SettingsStatusText = FormatLocalizedText(key, args);
    }

    private void SetUpdateStatus(string foreground, string key, params object[] args)
    {
        _updateStatusKey = key;
        _updateStatusArgs = args;
        UpdateStatusForeground = foreground;
        UpdateStatusText = FormatLocalizedText(key, args);
    }

    private void SetDnsStatus(string foreground, string text)
    {
        DnsStatusForeground = foreground;
        DnsStatusText = text;
    }

    private bool PersistSelectedSessionNote(bool showStatus, bool refreshFilter)
    {
        var session = _selectedSession;

        if (session is null)
        {
            return true;
        }

        var normalizedNote = NormalizeNote(SelectedSessionNote);
        var currentNote = NormalizeNote(session.Note);

        if (string.Equals(normalizedNote, currentNote, StringComparison.Ordinal))
        {
            return true;
        }

        try
        {
            if (string.IsNullOrWhiteSpace(normalizedNote))
            {
                _sessionNotes.Remove(session.SessionId);
            }
            else
            {
                _sessionNotes[session.SessionId] = normalizedNote;
            }

            session.Note = normalizedNote;
            UpdateSessionSearchBlob(session);
            _notesService.SaveNotes(_sessionNotes);

            if (!string.Equals(SelectedSessionNote, normalizedNote, StringComparison.Ordinal))
            {
                SelectedSessionNote = normalizedNote;
            }

            if (refreshFilter)
            {
                ApplyFilter(session.SessionId);
            }

            if (showStatus)
            {
                SetStatus(
                    "#F8E7D6",
                    string.IsNullOrWhiteSpace(normalizedNote) ? "StatusNoteCleared" : "StatusNoteSaved");
            }

            OnPropertyChanged(nameof(CanSaveSelectedSessionNote));
            OnPropertyChanged(nameof(CanClearSelectedSessionNote));
            return true;
        }
        catch (Exception exception)
        {
            if (showStatus)
            {
                SetStatus("#FFD6D6", "StatusNoteSaveFailed", exception.Message);
            }

            return false;
        }
    }

    private static string NormalizeNote(string? note)
    {
        return (note ?? string.Empty)
            .Replace("\r\n", "\n")
            .Trim();
    }

    private static void UpdateSessionSearchBlob(SessionRecord session)
    {
        session.SearchBlob = string.IsNullOrWhiteSpace(session.Note)
            ? session.BaseSearchBlob
            : $"{session.BaseSearchBlob} {session.Note}";
    }

    private NewSessionLaunchOptions BuildNewSessionLaunchOptions()
    {
        return new NewSessionLaunchOptions
        {
            Prompt = NewSessionPrompt,
            WorkingDirectory = GetNormalizedNewSessionWorkingDirectory(),
            Model = NewSessionModel,
            Profile = NewSessionProfile,
            SandboxMode = SelectedSandboxMode,
            ApprovalPolicy = SelectedApprovalPolicy,
            LocalProvider = SelectedLocalProvider,
            UseSearch = NewSessionUseSearch,
            UseOss = NewSessionUseOss,
            UseFullAuto = NewSessionUseFullAuto
        };
    }

    private string GetNormalizedNewSessionWorkingDirectory()
    {
        return string.IsNullOrWhiteSpace(NewSessionWorkingDirectory)
            ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            : NewSessionWorkingDirectory.Trim();
    }

    private async Task RefreshDnsAdaptersAsync(bool preserveStatus = true)
    {
        try
        {
            IsDnsBusy = true;

            if (!preserveStatus)
            {
                SetDnsStatus("#F8E7D6", Strings["DnsStatusRefreshing"]);
            }

            var adapters = await Task.Run(_dnsManagementService.GetAdapters);
            var preferredIndex = SelectedDnsAdapter?.InterfaceIndex;

            DnsAdapters.Clear();

            foreach (var adapter in adapters)
            {
                DnsAdapters.Add(adapter);
            }

            SelectedDnsAdapter = DnsAdapters.FirstOrDefault(item => item.InterfaceIndex == preferredIndex) ??
                                 DnsAdapters.FirstOrDefault();

            if (SelectedDnsAdapter is null)
            {
                SetDnsStatus("#FFD6D6", Strings["DnsStatusNoAdapters"]);
            }
            else if (!preserveStatus)
            {
                SetDnsStatus("#F8E7D6", Strings["DnsStatusRefreshed"]);
            }
        }
        catch (Exception exception)
        {
            SetDnsStatus("#FFD6D6", Strings.Format("DnsStatusFailed", exception.Message));
        }
        finally
        {
            IsDnsBusy = false;
            RefreshDnsCommandStates();
        }
    }

    private async Task RefreshSettingsSectionAsync()
    {
        await RefreshUpdateStatusAsync();
    }

    private async Task RefreshUpdateStatusAsync()
    {
        if (IsUpdateBusy)
        {
            return;
        }

        IsUpdateBusy = true;
        SetUpdateStatus("#F8E7D6", "UpdateStatusChecking");

        try
        {
            var snapshot = await _updateService.GetLatestReleaseAsync();
            _lastAppUpdateSnapshot = snapshot;
            ApplyUpdateSnapshot(snapshot);

            if (snapshot.IsUpdateAvailable)
            {
                if (snapshot.HasInstallerAsset)
                {
                    SetUpdateStatus("#F8E7D6", "UpdateStatusAvailable", snapshot.LatestVersionDisplay);
                }
                else
                {
                    SetUpdateStatus("#FFD98C", "UpdateStatusNoInstaller");
                }
            }
            else
            {
                SetUpdateStatus("#F8E7D6", "UpdateStatusUpToDate", snapshot.CurrentVersionDisplay);
            }
        }
        catch (Exception exception)
        {
            SetUpdateStatus("#FFD6D6", "UpdateStatusCheckFailed", exception.Message);
        }
        finally
        {
            IsUpdateBusy = false;
            RefreshUpdateCommandStates();
        }
    }

    private async Task RefreshSetupStatusAsync()
    {
        if (IsSetupBusy)
        {
            return;
        }

        IsSetupBusy = true;
        SetSetupStatus("#F8E7D6", Strings["SetupStatusChecking"]);

        try
        {
            var snapshot = await Task.Run(_environmentService.GetEnvironmentSnapshot);
            _lastEnvironmentSnapshot = snapshot;
            ApplySetupSnapshot(snapshot);
            SetSetupStatus("#F8E7D6", Strings["SetupStatusChecked"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
        finally
        {
            IsSetupBusy = false;
            OnPropertyChanged(nameof(CanLaunchNewSession));
            OnPropertyChanged(nameof(CanLaunchCodexLogin));
            OnPropertyChanged(nameof(CanInstallLocalAiTools));
            OnPropertyChanged(nameof(CanInstallLocalAiModels));
            OnPropertyChanged(nameof(CanManageCreativeAiTools));
            OnPropertyChanged(nameof(CanManageAiAgents));
        }
    }

    private bool ConfirmDnsWarning(string title, string message)
    {
        return MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private static bool ConfirmLocalAiRemoval(string title, string message)
    {
        return MessageBox.Show(
            message,
            title,
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private void RefreshDnsCommandStates()
    {
        OnPropertyChanged(nameof(CanApplyDnsPreset));
        OnPropertyChanged(nameof(CanEditSelectedDnsPreset));
        OnPropertyChanged(nameof(CanDeleteSelectedDnsPreset));
        OnPropertyChanged(nameof(CanEditDnsFields));
        OnPropertyChanged(nameof(CanResetDnsAutomatic));
        OnPropertyChanged(nameof(CanRestorePreviousDns));
        OnPropertyChanged(nameof(CanRefreshDnsAdapters));
        OnPropertyChanged(nameof(DnsDohTemplateVisibility));
        OnPropertyChanged(nameof(SelectedDnsPresetDescriptionText));
        OnPropertyChanged(nameof(SelectedDnsAdapterDescriptionText));
        OnPropertyChanged(nameof(SelectedDnsAdapterServersText));
    }

    private void RefreshUpdateCommandStates()
    {
        OnPropertyChanged(nameof(CanCheckForUpdates));
        OnPropertyChanged(nameof(CanDownloadUpdate));
        OnPropertyChanged(nameof(CanOpenReleasePage));
        OnPropertyChanged(nameof(CurrentAppVersionText));
        OnPropertyChanged(nameof(LatestAppVersionText));
        OnPropertyChanged(nameof(UpdateReleaseTitleText));
        OnPropertyChanged(nameof(UpdatePublishedText));
    }

    private void ApplyUpdateSnapshot(AppUpdateSnapshot snapshot)
    {
        OnPropertyChanged(nameof(CurrentAppVersionText));
        OnPropertyChanged(nameof(LatestAppVersionText));
        OnPropertyChanged(nameof(UpdateReleaseTitleText));
        OnPropertyChanged(nameof(UpdatePublishedText));
        RefreshUpdateCommandStates();
    }

    private void ApplySetupSnapshot(CodexEnvironmentSnapshot snapshot)
    {
        RefreshLocalAiModelOptions(snapshot.InstalledOllamaModels);
        RefreshCreativeAiToolOptions(snapshot);
        RefreshAiAgentToolOptions(snapshot);
        SetupCoreChecks.Clear();
        SetupCodexChecks.Clear();
        SetupLocalAiChecks.Clear();

        SetupCoreChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckWinget"],
                snapshot.WingetAvailable ? Strings["SetupBadgeFound"] : Strings["SetupBadgeMissing"],
                snapshot.WingetAvailable ? snapshot.WingetVersion : Strings["SetupDetailWingetMissing"],
                snapshot.WingetAvailable));
        SetupCoreChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckNode"],
                snapshot.NodeAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.NodeAvailable ? snapshot.NodeVersion : Strings["SetupDetailNodeMissing"],
                snapshot.NodeAvailable));
        SetupCoreChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckNpm"],
                snapshot.NpmAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.NpmAvailable ? snapshot.NpmVersion : Strings["SetupDetailNpmMissing"],
                snapshot.NpmAvailable));
        SetupCoreChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckGit"],
                snapshot.GitAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.GitAvailable ? snapshot.GitVersion : Strings["SetupDetailGitMissing"],
                snapshot.GitAvailable));
        SetupCodexChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckCodex"],
                snapshot.CodexAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.CodexAvailable ? snapshot.CodexVersion : Strings["SetupDetailCodexMissing"],
                snapshot.CodexAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOllama"],
                snapshot.OllamaAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.OllamaAvailable ? snapshot.OllamaDetail : Strings["SetupDetailOllamaMissing"],
                snapshot.OllamaAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckLmStudio"],
                snapshot.LmStudioAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.LmStudioAvailable ? snapshot.LmStudioDetail : Strings["SetupDetailLmStudioMissing"],
                snapshot.LmStudioAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckComfyUi"],
                snapshot.ComfyUiDesktopAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.ComfyUiDesktopAvailable ? snapshot.ComfyUiDesktopDetail : Strings["SetupDetailComfyUiMissing"],
                snapshot.ComfyUiDesktopAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckPinokio"],
                snapshot.PinokioAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.PinokioAvailable ? snapshot.PinokioDetail : Strings["SetupDetailPinokioMissing"],
                snapshot.PinokioAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOpenClaw"],
                snapshot.OpenClawAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.OpenClawAvailable ? snapshot.OpenClawDetail : Strings["SetupDetailOpenClawMissing"],
                snapshot.OpenClawAvailable));
        SetupCodexChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckLogin"],
                snapshot.LoggedIn ? Strings["SetupBadgeLoggedIn"] : Strings["SetupBadgeNeedsLogin"],
                string.IsNullOrWhiteSpace(snapshot.LoginStatus)
                    ? Strings["SetupDetailLoginUnknown"]
                    : snapshot.LoginStatus,
                snapshot.LoggedIn,
                isWarning: !snapshot.LoggedIn));
        SetupCodexChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckSessionsFolder"],
                snapshot.SessionsFolderExists ? Strings["SetupBadgeExists"] : Strings["SetupBadgeMissing"],
                snapshot.SessionsFolderPath,
                snapshot.SessionsFolderExists));

        OnPropertyChanged(nameof(CanInstallLocalAiModels));
        OnPropertyChanged(nameof(CanManageCreativeAiTools));
        OnPropertyChanged(nameof(CanManageAiAgents));
        OnPropertyChanged(nameof(CanUninstallOllama));
        OnPropertyChanged(nameof(CanUninstallLmStudio));
    }

    private static SetupCheckItem CreateSetupCheckItem(
        string title,
        string status,
        string detail,
        bool isOk,
        bool isWarning = false)
    {
        var accentBrush = isOk
            ? "#1F7A52"
            : isWarning
                ? "#B86E10"
                : "#B42318";

        return new SetupCheckItem
        {
            Title = title,
            Status = status,
            Detail = detail,
            AccentBrush = accentBrush
        };
    }

    private static void OpenExplorerSelect(string path)
    {
        Process.Start(
            new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{path}\"",
                UseShellExecute = true
            });
    }

    private static string QuoteForCommandLine(string value)
    {
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private async Task RefreshSessionsAsync(bool isAutomaticRefresh)
    {
        if (_isRefreshing)
        {
            return;
        }

        if (!PersistSelectedSessionNote(showStatus: !isAutomaticRefresh, refreshFilter: false) &&
            !isAutomaticRefresh)
        {
            return;
        }

        _isRefreshing = true;
        IsLoading = true;
        SetStatus("#F8E7D6", isAutomaticRefresh ? "StatusRefreshing" : "StatusReading");

        try
        {
            var currentLanguage = Strings.CurrentLanguage;
            var refreshedSessions = await Task.Run(() => _sessionService.GetSessions(currentLanguage));
            ApplySessions(refreshedSessions);

            _lastUpdatedAtLocal = DateTime.Now;
            LastUpdatedText = Strings.Format("LastUpdated", _lastUpdatedAtLocal.Value.ToString("dd.MM.yyyy HH:mm:ss"));
            SetStatus(
                "#F8E7D6",
                refreshedSessions.Count == 0 ? "StatusNoSessions" : "StatusLoadedCount",
                refreshedSessions.Count);
        }
        catch (Exception exception)
        {
            SetStatus("#FFD6D6", "StatusError", exception.Message);
        }
        finally
        {
            IsLoading = false;
            _isRefreshing = false;
        }
    }

    private void UpdateRefreshTimer()
    {
        if (AutoRefreshEnabled)
        {
            _refreshTimer.Start();
        }
        else
        {
            _refreshTimer.Stop();
        }
    }

    private void FitToWorkArea()
    {
        var workArea = SystemParameters.WorkArea;
        var availableWidth = Math.Max(960, workArea.Width - 12);
        var availableHeight = Math.Max(600, workArea.Height - 12);

        MinWidth = Math.Min(MinWidth, availableWidth);
        MinHeight = Math.Min(MinHeight, availableHeight);
        Width = Math.Min(Width, availableWidth);
        Height = Math.Min(Height, availableHeight);
        Left = workArea.Left + Math.Max((workArea.Width - Width) / 2, 0);
        Top = workArea.Top + Math.Max((workArea.Height - Height) / 2, 0);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}



