using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using LaptopSessionViewer.Models;
using LaptopSessionViewer.Services;
using Microsoft.Win32;

namespace LaptopSessionViewer;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private static readonly TimeSpan SearchDebounceInterval = TimeSpan.FromMilliseconds(220);
    private static readonly TimeSpan SessionRefreshTimerInterval = TimeSpan.FromSeconds(25);
    private static readonly TimeSpan SessionRefreshFallbackInterval = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan SetupRefreshNormalInterval = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan SetupRefreshBusyInterval = TimeSpan.FromSeconds(8);
    private static readonly TimeSpan SetupRefreshFallbackInterval = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan SetupRefreshBoostDuration = TimeSpan.FromMinutes(3);
    private static readonly TimeSpan UpdateRefreshCacheDuration = TimeSpan.FromMinutes(15);
    private static readonly TimeSpan LayoutRefreshDebounceInterval = TimeSpan.FromMilliseconds(120);
    private readonly AppLogService _logService = new();
    private readonly AiExtensionCatalogService _extensionCatalogService = new();
    private readonly CodexEnvironmentService _environmentService = new();
    private readonly AiExtensionManagementService _extensionManagementService;
    private readonly AppUpdateService _updateService = new();
    private readonly CodexPhotoPasteFixService _photoPasteFixService;
    private readonly DnsManagementService _dnsManagementService = new();
    private readonly DnsPresetSettingsService _dnsPresetSettingsService = new();
    private readonly OpenCodeSessionBridgeService _openCodeBridgeService;
    private readonly OpenCodeSessionLinkService _openCodeLinkService = new();
    private readonly SessionFavoritesService _favoritesService = new();
    private readonly SessionNotesService _notesService = new();
    private readonly SessionService _sessionService = new();
    private readonly SessionFeedExportService _sessionFeedExportService = new();
    private readonly SessionCheckpointService _checkpointService = new();
    private readonly SessionVisibilityService _sessionVisibilityService = new();
    private readonly SessionViewerSettingsService _settingsService = new();
    private readonly DispatcherTimer _refreshTimer;
    private readonly DispatcherTimer _searchDebounceTimer;
    private readonly DispatcherTimer _setupRefreshTimer;
    private readonly DispatcherTimer _layoutRefreshTimer;
    private FileSystemWatcher? _sessionFolderWatcher;
    private FileSystemWatcher? _sessionIndexWatcher;
    private List<SessionRecord> _allSessions = [];
    private HashSet<string> _favoriteSessionIds = [];
    private HashSet<string> _hiddenSessionIds = [];
    private Dictionary<string, OpenCodeSessionLinkRecord> _openCodeLinks = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string> _sessionNotes = new(StringComparer.OrdinalIgnoreCase);
    private bool _autoRefreshEnabled = true;
    private bool _isLoading;
    private bool _isOpenCodeBusy;
    private bool _isRefreshing;
    private bool _isDnsBusy;
    private bool _isApplyingDangerousAccessDefaults;
    private bool _isBeginnerModeEnabled = true;
    private bool _beginnerOnboardingInProgress;
    private bool _hasCompletedBeginnerOnboarding;
    private bool _isSessionsSurfaceInitialized;
    private bool _isSetupBusy;
    private bool _showHiddenSessions;
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
    private string _extensionCommandOrUri = string.Empty;
    private string _extensionDescription = string.Empty;
    private bool _extensionIsEnabled = true;
    private string _extensionName = string.Empty;
    private string _extensionSearchText = string.Empty;
    private string _extensionStatusForeground = "#F8E7D6";
    private string _extensionStatusText = string.Empty;
    private string _homeLaunchStatusBackground = "#E7F6EE";
    private string _homeLaunchStatusForeground = "#1F6F4A";
    private string _homeLaunchStatusText = string.Empty;
    private DateTime? _lastUpdatedAtLocal;
    private string _lastUpdatedText = string.Empty;
    private string _newSessionModel = string.Empty;
    private string _newSessionProfile = string.Empty;
    private string _newSessionPrompt = string.Empty;
    private string _newSessionStatusForeground = "#F8E7D6";
    private string _newSessionStatusText = string.Empty;
    private bool _newSessionUseOss;
    private bool _newSessionUseSearch;
    private string _newSessionWorkingDirectory = string.Empty;
    private string _primaryDnsServer = string.Empty;
    private DnsAdapterRecord? _selectedDnsAdapter;
    private DnsPreset? _selectedDnsPreset;
    private AiExtensionItem? _selectedExtension;
    private string _selectedExtensionKind = "Plugin";
    private string _selectedExtensionTarget = "All";
    private string _extensionTargetApp = "Codex";
    private string _selectedApprovalPolicy = "on-request";
    private AppSection _selectedAppSection = AppSection.Home;
    private string _searchText = string.Empty;
    private string _secondaryDnsServer = string.Empty;
    private string _selectedLocalProvider = string.Empty;
    private string _selectedSessionTranscriptText = string.Empty;
    private string _selectedSandboxMode = "workspace-write";
    private SettingsCategoryTab _selectedSettingsCategoryTab = SettingsCategoryTab.AppSettings;
    private SessionListTab _selectedSessionListTab = SessionListTab.Sessions;
    private LanguageOption? _selectedLanguageOption;
    private bool _settingsDangerousFullAccess;
    private bool _settingsPhotoPasteFixEnabled;
    private string _settingsStatusForeground = "#F8E7D6";
    private string _settingsStatusKey = "SettingsStatusReady";
    private object[] _settingsStatusArgs = [];
    private string _settingsStatusText = string.Empty;
    private string _selectedSessionNote = string.Empty;
    private SessionRecord? _selectedSession;
    private bool _sessionRefreshPending = true;
    private int _selectedSessionTranscriptLoadVersion;
    private bool _selectedSessionTranscriptLoading;
    private bool _setupRefreshPending = true;
    private DateTime _lastSessionRefreshCompletedUtc = DateTime.MinValue;
    private DateTime _lastSetupRefreshCompletedUtc = DateTime.MinValue;
    private DateTime _lastUpdateRefreshCompletedUtc = DateTime.MinValue;
    private readonly Stopwatch _startupStopwatch = Stopwatch.StartNew();
    private bool _isNewSessionSectionInitialized;
    private bool _isExtensionsSectionInitialized;
    private bool _isSetupSectionInitialized;
    private bool _isSettingsSectionInitialized;
    private bool _isDetectedExtensionsRefreshRunning;
    private bool _isManagedExtensionsRefreshRunning;
    private bool _showCustomExtensionsTab;
    private bool _showBeginnerOnboarding;
    private bool _showInstalledExtensionsTab;
    private bool _startupRefreshScheduled;
    private DateTime _setupRefreshBoostUntilUtc = DateTime.MinValue;
    private string _setupStatusForeground = "#1F6F4A";
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
        LogStartupPhase("MainWindow constructor started.");
        var initialSettings = _settingsService.LoadSettings();
        LogStartupPhase("Settings loaded.");
        Strings.SetLanguage(initialSettings.Language);
        _settingsDangerousFullAccess = initialSettings.DefaultDangerousFullAccess;
        _settingsPhotoPasteFixEnabled = initialSettings.PhotoPasteFixEnabled;
        _isBeginnerModeEnabled = initialSettings.BeginnerModeEnabled;
        _hasCompletedBeginnerOnboarding = initialSettings.HasCompletedBeginnerOnboarding;
        _showBeginnerOnboarding = _isBeginnerModeEnabled && !_hasCompletedBeginnerOnboarding;
        _selectedLanguageOption = LanguageOptions.First(option => option.Language == initialSettings.Language);
        _extensionManagementService = new AiExtensionManagementService(_environmentService, _logService);
        _openCodeBridgeService = new OpenCodeSessionBridgeService(_logService);
        _photoPasteFixService = new CodexPhotoPasteFixService(_logService);
        LoadSessionMetadata();
        LoadOpenCodeLinks();
        _selectedSessionTranscriptText = Strings["NoTranscriptLoaded"];
        LogStartupPhase("Session metadata loaded.");

        InitializeComponent();
        DataContext = this;
        AddHandler(PreviewMouseWheelEvent, new MouseWheelEventHandler(RouteMouseWheelToScrollableParent), true);
        LogStartupPhase("InitializeComponent completed.");
        RefreshLocalizedChromeText();
        RefreshSectionChromeText();
        LogStartupPhase("Initial chrome text refreshed.");

        _refreshTimer = new DispatcherTimer
        {
            Interval = SessionRefreshTimerInterval
        };
        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = SearchDebounceInterval
        };
        _setupRefreshTimer = new DispatcherTimer
        {
            Interval = SetupRefreshNormalInterval
        };
        _layoutRefreshTimer = new DispatcherTimer
        {
            Interval = LayoutRefreshDebounceInterval
        };

        SourceInitialized += MainWindow_SourceInitialized;
        _refreshTimer.Tick += RefreshTimer_Tick;
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;
        _setupRefreshTimer.Tick += SetupRefreshTimer_Tick;
        _layoutRefreshTimer.Tick += LayoutRefreshTimer_Tick;
        Activated += MainWindow_Activated;
        Deactivated += MainWindow_Deactivated;
        Loaded += MainWindow_Loaded;
        Closed += MainWindow_Closed;
        SizeChanged += MainWindow_SizeChanged;
        StateChanged += MainWindow_StateChanged;
        LogStartupPhase("MainWindow constructor finished.");
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public LocalizationService Strings { get; } = new();

    public ObservableCollection<DnsAdapterRecord> DnsAdapters { get; } = [];

    public ObservableCollection<DnsPreset> DnsPresets { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupCoreChecks { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupCodexChecks { get; } = [];

    public ObservableCollection<SetupCheckItem> SetupLocalAiChecks { get; } = [];

    public ObservableCollection<SetupCheckItem> OllamaQuickChecks { get; } = [];

    public ObservableCollection<LocalAiModelOption> LocalAiModelOptions { get; } = [];

    public ObservableCollection<CreativeAiToolOption> CreativeAiToolOptions { get; } = [];

    public ObservableCollection<CreativeAiToolOption> AiAgentToolOptions { get; } = [];

    public ObservableCollection<SetupCheckItem> OpenClawSetupModes { get; } = [];

    public ObservableCollection<SetupCheckItem> OpenClawCapabilityChecks { get; } = [];

    public ObservableCollection<AiExtensionItem> AiExtensions { get; } = [];

    public ObservableCollection<AiExtensionItem> SuggestedAiExtensions { get; } = [];

    public ObservableCollection<AiExtensionItem> InstalledAiExtensions { get; } = [];

    public ObservableCollection<AiExtensionItem> CustomAiExtensions { get; } = [];

    public ObservableCollection<LaunchOption> ExtensionKindOptions { get; } = [];

    public ObservableCollection<LaunchOption> ExtensionTargetOptions { get; } = [];

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

    public ObservableCollection<SessionRecord> HomeRecentSessions { get; } = [];

    public Thickness AppOuterMargin =>
        IsWideWindowLayout ? new Thickness(16) : IsCompactWindowLayout ? new Thickness(10) : new Thickness(12);

    public double AppContentMaxWidth => IsWideWindowLayout ? 1820 : 1720;

    public double AppContentWidth
    {
        get
        {
            var measuredWidth = ActualWidth > 0 ? ActualWidth : Width;
            var margin = AppOuterMargin;
            var availableWidth = Math.Max(0, measuredWidth - margin.Left - margin.Right - 16);
            return Math.Min(AppContentMaxWidth, availableWidth);
        }
    }

    public GridLength ShellSidebarColumnWidth =>
        new(IsWideWindowLayout ? 240 : IsCompactWindowLayout ? 224 : 232);

    public GridLength ShellMainGapColumnWidth =>
        new(IsWideWindowLayout ? 16 : IsCompactWindowLayout ? 10 : 12);

    public GridLength SectionRailGapColumnWidth =>
        new(IsWideWindowLayout ? 18 : IsCompactWindowLayout ? 12 : 16);

    public GridLength SessionsHeaderSearchColumnWidth =>
        new(IsWideWindowLayout ? 320 : IsCompactWindowLayout ? 232 : 280);

    public GridLength SessionsActionButtonColumnWidth =>
        new(IsWideWindowLayout ? 164 : IsCompactWindowLayout ? 138 : 150);

    public GridLength SessionsDetailColumnWidth =>
        SharedAsideColumnWidth;

    public GridLength NewSessionAsideColumnWidth =>
        SharedAsideColumnWidth;

    public GridLength SettingsAsideColumnWidth =>
        SharedAsideColumnWidth;

    public double HomeTitleFontSize => IsCompactWindowLayout ? 30 : 36;

    public double HomePromptHeight => IsCompactWindowLayout ? 54 : 78;

    public GridLength HomeSafetyColumnWidth => new(IsCompactWindowLayout ? 260 : 250);

    private GridLength SharedAsideColumnWidth =>
        new(IsWideWindowLayout ? 416 : IsCompactWindowLayout ? 336 : 370);

    public AppSection SelectedAppSection
    {
        get => _selectedAppSection;
        set
        {
            if (SetField(ref _selectedAppSection, value))
            {
                EnsureSectionDataInitialized(value);
                OnPropertyChanged(nameof(HomeSectionButtonBackground));
                OnPropertyChanged(nameof(HomeSectionButtonForeground));
                OnPropertyChanged(nameof(SessionsSectionButtonBackground));
                OnPropertyChanged(nameof(SessionsSectionButtonForeground));
                OnPropertyChanged(nameof(NewSessionSectionButtonBackground));
                OnPropertyChanged(nameof(NewSessionSectionButtonForeground));
                OnPropertyChanged(nameof(ExtensionsSectionButtonBackground));
                OnPropertyChanged(nameof(ExtensionsSectionButtonForeground));
                OnPropertyChanged(nameof(SetupSectionButtonBackground));
                OnPropertyChanged(nameof(SetupSectionButtonForeground));
                OnPropertyChanged(nameof(SettingsSectionButtonBackground));
                OnPropertyChanged(nameof(SettingsSectionButtonForeground));
                OnPropertyChanged(nameof(HomeSectionVisibility));
                OnPropertyChanged(nameof(SessionsSectionVisibility));
                OnPropertyChanged(nameof(NewSessionSectionVisibility));
                OnPropertyChanged(nameof(ExtensionsSectionVisibility));
                OnPropertyChanged(nameof(SetupSectionVisibility));
                OnPropertyChanged(nameof(SettingsSectionVisibility));
                UpdateRefreshTimer();
                UpdateSetupRefreshTimer();

                if (value == AppSection.Setup && IsLoaded)
                {
                    _ = RefreshSetupSectionAsync(preserveDnsStatus: true);
                }

                if (value == AppSection.Home && IsLoaded && _lastEnvironmentSnapshot is null)
                {
                    EnsureSectionDataInitialized(AppSection.Setup);
                    _ = RefreshSetupSectionAsync(preserveDnsStatus: true);
                }

                if (value == AppSection.Settings && IsLoaded)
                {
                    _ = RefreshSettingsSectionAsync();
                }

                if (value == AppSection.Sessions && IsLoaded)
                {
                    EnsureSessionsSurfaceInitialized();
                    _ = RefreshSessionsAsync(isAutomaticRefresh: false);
                }
            }
        }
    }

    public string HomeSectionButtonBackground =>
        SelectedAppSection == AppSection.Home ? "#F7F3EC" : "#1D3545";

    public string HomeSectionButtonForeground =>
        SelectedAppSection == AppSection.Home ? "#16212B" : "#FFFDF9";

    public string SessionsSectionButtonBackground =>
        SelectedAppSection == AppSection.Sessions ? "#F7F3EC" : "#1D3545";

    public string SessionsSectionButtonForeground =>
        SelectedAppSection == AppSection.Sessions ? "#16212B" : "#FFFDF9";

    public string NewSessionSectionButtonBackground =>
        SelectedAppSection == AppSection.NewSession ? "#F7F3EC" : "#1D3545";

    public string NewSessionSectionButtonForeground =>
        SelectedAppSection == AppSection.NewSession ? "#16212B" : "#FFFDF9";

    public string ExtensionsSectionButtonBackground =>
        SelectedAppSection == AppSection.Extensions ? "#F7F3EC" : "#1D3545";

    public string ExtensionsSectionButtonForeground =>
        SelectedAppSection == AppSection.Extensions ? "#16212B" : "#FFFDF9";

    public string SetupSectionButtonBackground =>
        SelectedAppSection == AppSection.Setup ? "#F7F3EC" : "#1D3545";

    public string SetupSectionButtonForeground =>
        SelectedAppSection == AppSection.Setup ? "#16212B" : "#FFFDF9";

    public string SettingsSectionButtonBackground =>
        SelectedAppSection == AppSection.Settings ? "#F7F3EC" : "#1D3545";

    public string SettingsSectionButtonForeground =>
        SelectedAppSection == AppSection.Settings ? "#16212B" : "#FFFDF9";

    public Visibility SessionsSectionVisibility =>
        SelectedAppSection == AppSection.Sessions ? Visibility.Visible : Visibility.Collapsed;

    public Visibility HomeSectionVisibility =>
        SelectedAppSection == AppSection.Home ? Visibility.Visible : Visibility.Collapsed;

    public bool IsBeginnerModeEnabled
    {
        get => _isBeginnerModeEnabled;
        set
        {
            if (!SetField(ref _isBeginnerModeEnabled, value))
            {
                return;
            }

            _settingsService.SaveBeginnerModeEnabled(value);

            if (value)
            {
                SelectedSettingsCategoryTab = SettingsCategoryTab.AppSettings;
                _showBeginnerOnboarding = !_hasCompletedBeginnerOnboarding;
                SelectedAppSection = AppSection.Home;
            }

            OnPropertyChanged(nameof(ExpertOnlyVisibility));
            OnPropertyChanged(nameof(BeginnerOnlyVisibility));
            OnPropertyChanged(nameof(BeginnerOnboardingVisibility));
            OnPropertyChanged(nameof(HomeWorkspaceVisibility));
            OnPropertyChanged(nameof(SetupSubtitleText));
        }
    }

    public Visibility ExpertOnlyVisibility =>
        IsBeginnerModeEnabled ? Visibility.Collapsed : Visibility.Visible;

    public Visibility BeginnerOnlyVisibility =>
        IsBeginnerModeEnabled ? Visibility.Visible : Visibility.Collapsed;

    public Visibility BeginnerOnboardingVisibility =>
        IsBeginnerModeEnabled && _showBeginnerOnboarding ? Visibility.Visible : Visibility.Collapsed;

    public Visibility HomeWorkspaceVisibility =>
        !IsBeginnerModeEnabled || !_showBeginnerOnboarding ? Visibility.Visible : Visibility.Collapsed;

    public bool HasHomeRecentSessions => HomeRecentSessions.Count > 0;

    public Visibility HomeRecentSessionsVisibility =>
        HasHomeRecentSessions ? Visibility.Visible : Visibility.Collapsed;

    public Visibility HomeRecentSessionsEmptyVisibility =>
        HasHomeRecentSessions ? Visibility.Collapsed : Visibility.Visible;

    public string SetupSubtitleText =>
        IsBeginnerModeEnabled ? Strings["SetupBeginnerSubtitle"] : Strings["SetupSubtitle"];

    public bool IsHomeEnvironmentReady =>
        _lastEnvironmentSnapshot is not null &&
        _lastEnvironmentSnapshot.CodexAvailable &&
        _lastEnvironmentSnapshot.LoggedIn;

    public bool CanStartHomeSession =>
        IsHomeEnvironmentReady &&
        !string.IsNullOrWhiteSpace(NewSessionPrompt);

    public string HomeReadinessText =>
        _lastEnvironmentSnapshot is null
            ? Strings["HomeReadinessChecking"]
            : !_lastEnvironmentSnapshot.CodexAvailable
                ? Strings["HomeReadinessNeedsCodex"]
                : !_lastEnvironmentSnapshot.LoggedIn
                    ? Strings["HomeReadinessNeedsLogin"]
                    : Strings["HomeReadinessReady"];

    public string HomeReadinessBrush =>
        _lastEnvironmentSnapshot is null ? "#355364" : IsHomeEnvironmentReady ? "#1F7A52" : "#8A4B08";

    public string HomeStartHelpText =>
        !IsHomeEnvironmentReady
            ? Strings["HomeLaunchSetupRequired"]
            : string.IsNullOrWhiteSpace(NewSessionPrompt)
                ? Strings["HomeLaunchPromptRequired"]
                : Strings["HomeStartSafe"];

    public string HomeLaunchStatusText
    {
        get => _homeLaunchStatusText;
        private set
        {
            if (SetField(ref _homeLaunchStatusText, value))
            {
                OnPropertyChanged(nameof(HomeLaunchStatusVisibility));
            }
        }
    }

    public string HomeLaunchStatusForeground
    {
        get => _homeLaunchStatusForeground;
        private set => SetField(ref _homeLaunchStatusForeground, value);
    }

    public string HomeLaunchStatusBackground
    {
        get => _homeLaunchStatusBackground;
        private set => SetField(ref _homeLaunchStatusBackground, value);
    }

    public Visibility HomeLaunchStatusVisibility =>
        string.IsNullOrWhiteSpace(HomeLaunchStatusText) ? Visibility.Collapsed : Visibility.Visible;

    public string BeginnerSetupLocalAiStatusText =>
        _lastEnvironmentSnapshot is null
            ? Strings["SetupBeginnerChecking"]
            : _lastEnvironmentSnapshot.OllamaAvailable ||
              _lastEnvironmentSnapshot.OllamaAppAvailable ||
              _lastEnvironmentSnapshot.LmStudioAvailable
                ? Strings["SetupBeginnerLocalAiInstalled"]
                : Strings["SetupBeginnerLocalAiOptional"];

    public string BeginnerSetupLocalAiStatusBrush =>
        _lastEnvironmentSnapshot?.OllamaAvailable == true ||
        _lastEnvironmentSnapshot?.OllamaAppAvailable == true ||
        _lastEnvironmentSnapshot?.LmStudioAvailable == true
            ? "#1F7A52"
            : "#667085";

    public object? SessionsMainContent => _isSessionsSurfaceInitialized ? true : null;

    public Visibility SessionsMainPlaceholderVisibility =>
        _isSessionsSurfaceInitialized ? Visibility.Collapsed : Visibility.Visible;

    public Visibility NewSessionSectionVisibility =>
        SelectedAppSection == AppSection.NewSession ? Visibility.Visible : Visibility.Collapsed;

    public object? NewSessionSectionContent => _isNewSessionSectionInitialized ? true : null;

    public Visibility ExtensionsSectionVisibility =>
        SelectedAppSection == AppSection.Extensions ? Visibility.Visible : Visibility.Collapsed;

    public object? ExtensionsSectionContent => _isExtensionsSectionInitialized ? true : null;

    public Visibility SetupSectionVisibility =>
        SelectedAppSection == AppSection.Setup ? Visibility.Visible : Visibility.Collapsed;

    public object? SetupSectionContent => _isSetupSectionInitialized ? true : null;

    public Visibility SettingsSectionVisibility =>
        SelectedAppSection == AppSection.Settings ? Visibility.Visible : Visibility.Collapsed;

    public object? SettingsSectionContent => _isSettingsSectionInitialized ? true : null;

    public SettingsCategoryTab SelectedSettingsCategoryTab
    {
        get => _selectedSettingsCategoryTab;
        set
        {
            if (SetField(ref _selectedSettingsCategoryTab, value))
            {
                OnPropertyChanged(nameof(NeuralSettingsTabBackground));
                OnPropertyChanged(nameof(NeuralSettingsTabForeground));
                OnPropertyChanged(nameof(AppSettingsTabBackground));
                OnPropertyChanged(nameof(AppSettingsTabForeground));
                OnPropertyChanged(nameof(NeuralSettingsTabVisibility));
                OnPropertyChanged(nameof(AppSettingsTabVisibility));
            }
        }
    }

    public string NeuralSettingsTabBackground =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.NeuralSettings ? "#E7D8CA" : "#16212B";

    public string NeuralSettingsTabForeground =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.NeuralSettings ? "#16212B" : "#FFFDF9";

    public string AppSettingsTabBackground =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.AppSettings ? "#E7D8CA" : "#16212B";

    public string AppSettingsTabForeground =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.AppSettings ? "#16212B" : "#FFFDF9";

    public Visibility NeuralSettingsTabVisibility =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.NeuralSettings ? Visibility.Visible : Visibility.Collapsed;

    public Visibility AppSettingsTabVisibility =>
        SelectedSettingsCategoryTab == SettingsCategoryTab.AppSettings ? Visibility.Visible : Visibility.Collapsed;

    public string SessionsTabText => $"{Strings["SessionsTab"]} ({RegularSessions})";

    public string FavoritesTabText => $"{Strings["FavoritesTab"]} ({FavoriteSessions})";

    public string HiddenSessionsToggleText => Strings.Format("ShowHiddenSessions", HiddenSessions);

    public bool ShowHiddenSessions
    {
        get => _showHiddenSessions;
        set
        {
            if (SetField(ref _showHiddenSessions, value))
            {
                ApplyFilter();
            }
        }
    }

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

    public bool CanArchiveSelectedSession =>
        SelectedSession is not null && File.Exists(SelectedSession.FilePath);

    public bool CanToggleSelectedSessionHidden => SelectedSession is not null;

    public string ToggleSessionHiddenButtonText =>
        SelectedSession?.IsHidden == true ? Strings["ShowSessionInList"] : Strings["HideSessionFromList"];

    public bool CanResumeSelectedSession =>
        SelectedSession is not null &&
        !string.IsNullOrWhiteSpace(SelectedSession.SessionId) &&
        ((SelectedSession.IsCodexSession && File.Exists(_environmentService.CodexCommandPath)) ||
         (SelectedSession.IsClaudeSession && File.Exists(_environmentService.ClaudeCommandPath)));

    public bool CanOpenSelectedSessionDirectory =>
        SelectedSession is not null &&
        !string.IsNullOrWhiteSpace(SelectedSession.WorkingDirectory) &&
        SelectedSession.WorkingDirectory != "-" &&
        Directory.Exists(SelectedSession.WorkingDirectory);

    public bool IsOpenCodeBusy
    {
        get => _isOpenCodeBusy;
        private set
        {
            if (SetField(ref _isOpenCodeBusy, value))
            {
                OnPropertyChanged(nameof(CanResumeSelectedSessionInOpenCode));
                OnPropertyChanged(nameof(CanRefreshSelectedSessionOpenCodeBridge));
            }
        }
    }

    public bool CanResumeSelectedSessionInOpenCode =>
        SelectedSession is not null &&
        !IsOpenCodeBusy;

    public bool CanRefreshSelectedSessionOpenCodeBridge =>
        SelectedSession is not null &&
        !IsOpenCodeBusy;

    public string OpenCodeResumeButtonText =>
        GetSelectedSessionOpenCodeLink() is null
            ? Strings["CreateOpenCodeBridge"]
            : Strings["ResumeInOpenCode"];

    public string SelectedSessionOpenCodeBridgeText => BuildSelectedSessionOpenCodeBridgeText();

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

    public bool CanCreateSelectedSessionCheckpoint => SelectedSession is not null;

    public bool CanLaunchNewSession =>
        !string.IsNullOrWhiteSpace(NewSessionPrompt) &&
        File.Exists(_environmentService.CodexCommandPath) &&
        Directory.Exists(GetNormalizedNewSessionWorkingDirectory());

    public bool CanInstallCodexDesktopApp => !IsSetupBusy;

    public bool CanOpenCodexDesktopStorePage => !IsSetupBusy;

    public bool CanLaunchCodexLogin => File.Exists(_environmentService.CodexCommandPath);

    public bool CanInstallOpenCode =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.NpmAvailable == true;

    public bool CanLaunchOpenCode =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenCodeAvailable == true;

    public bool CanUninstallOpenCode =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenCodeAvailable == true;

    public string OpenCodeSetupDetailText =>
        _lastEnvironmentSnapshot?.OpenCodeAvailable == true
            ? _lastEnvironmentSnapshot.OpenCodeDetail
            : Strings["SetupDetailOpenCodeMissing"];

    public bool CanInstallLocalAiTools => !IsSetupBusy;

    public bool CanInstallBaseComponents =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.WingetAvailable == true;

    public bool CanRepairWinget => !IsSetupBusy;

    public bool CanInstallLocalAiModels =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAvailable == true;

    public bool CanLaunchOllamaApp =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAppAvailable == true;

    public bool CanStartOllamaServer =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAvailable == true;

    public bool CanStopOllamaProcesses =>
        !IsSetupBusy &&
        (_lastEnvironmentSnapshot?.OllamaServerRunning == true ||
         _lastEnvironmentSnapshot?.OllamaTrayRunning == true);

    public bool CanInstallStarterOllamaModel =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OllamaAvailable == true;

    public string OllamaQuickGuidanceText => BuildOllamaQuickGuidanceText(_lastEnvironmentSnapshot);

    public bool CanManageCreativeAiTools => !IsSetupBusy;

    public bool CanManageAiAgents => !IsSetupBusy;

    public bool CanApplyOpenClawModes =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenClawAvailable == true;

    public bool CanInspectOpenClawStatus =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenClawAvailable == true;

    public bool CanInstallOpenClawNode =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenClawAvailable == true;

    public bool CanInspectOpenClawNode =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenClawAvailable == true;

    public bool CanInspectOpenClawBrowser =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot?.OpenClawAvailable == true;

    public bool CanOpenOpenClawConfig =>
        !IsSetupBusy &&
        _lastEnvironmentSnapshot is not null &&
        (File.Exists(_lastEnvironmentSnapshot.OpenClawConfigPath) ||
         Directory.Exists(Path.GetDirectoryName(_lastEnvironmentSnapshot.OpenClawConfigPath) ?? string.Empty));

    public string OpenClawDetectedConfigText => BuildOpenClawDetectedConfigText(_lastEnvironmentSnapshot);

    public string OpenClawRecommendationText => BuildOpenClawRecommendationText(_lastEnvironmentSnapshot);

    public bool CanUninstallOllama =>
        !IsSetupBusy &&
        (_lastEnvironmentSnapshot?.OllamaAvailable == true ||
         _lastEnvironmentSnapshot?.OllamaAppAvailable == true);

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
                OnPropertyChanged(nameof(CanInstallBaseComponents));
                OnPropertyChanged(nameof(CanRepairWinget));
                OnPropertyChanged(nameof(CanInstallCodexDesktopApp));
                OnPropertyChanged(nameof(CanOpenCodexDesktopStorePage));
                OnPropertyChanged(nameof(CanInstallOpenCode));
                OnPropertyChanged(nameof(CanLaunchOpenCode));
                OnPropertyChanged(nameof(CanUninstallOpenCode));
                OnPropertyChanged(nameof(CanInstallLocalAiTools));
                OnPropertyChanged(nameof(CanInstallLocalAiModels));
                OnPropertyChanged(nameof(CanLaunchOllamaApp));
                OnPropertyChanged(nameof(CanStartOllamaServer));
                OnPropertyChanged(nameof(CanStopOllamaProcesses));
                OnPropertyChanged(nameof(CanInstallStarterOllamaModel));
                OnPropertyChanged(nameof(CanManageCreativeAiTools));
                OnPropertyChanged(nameof(CanManageAiAgents));
                OnPropertyChanged(nameof(CanApplyOpenClawModes));
                OnPropertyChanged(nameof(CanInspectOpenClawStatus));
                OnPropertyChanged(nameof(CanInstallOpenClawNode));
                OnPropertyChanged(nameof(CanInspectOpenClawNode));
                OnPropertyChanged(nameof(CanInspectOpenClawBrowser));
                OnPropertyChanged(nameof(CanOpenOpenClawConfig));
                OnPropertyChanged(nameof(CanUninstallOllama));
                OnPropertyChanged(nameof(CanUninstallLmStudio));
                OnPropertyChanged(nameof(OpenCodeSetupDetailText));
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
                OnPropertyChanged(nameof(SetupCoreSectionCollapsedIndicatorVisibility));
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
                OnPropertyChanged(nameof(SetupCodexSectionCollapsedIndicatorVisibility));
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
                OnPropertyChanged(nameof(SetupLocalAiSectionCollapsedIndicatorVisibility));
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
                OnPropertyChanged(nameof(SetupDnsSectionCollapsedIndicatorVisibility));
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

    public Visibility SetupCoreSectionCollapsedIndicatorVisibility =>
        IsSetupCoreSectionExpanded ? Visibility.Collapsed : Visibility.Visible;

    public Visibility SetupCodexSectionCollapsedIndicatorVisibility =>
        IsSetupCodexSectionExpanded ? Visibility.Collapsed : Visibility.Visible;

    public Visibility SetupLocalAiSectionCollapsedIndicatorVisibility =>
        IsSetupLocalAiSectionExpanded ? Visibility.Collapsed : Visibility.Visible;

    public Visibility SetupDnsSectionCollapsedIndicatorVisibility =>
        IsSetupDnsSectionExpanded ? Visibility.Collapsed : Visibility.Visible;

    public string SetupLiveStatusHintText =>
        IsSetupRefreshBoostActive
            ? Strings["SetupLiveStatusBoostHint"]
            : Strings["SetupLiveStatusHint"];

    public string SetupRecommendedNextStepText =>
        GetSetupRecommendedNextStepText(_lastEnvironmentSnapshot);

    public string SetupCoreProgressText =>
        GetSetupSectionProgressText(GetCoreReadyCount(_lastEnvironmentSnapshot), 4);

    public string SetupCodexProgressText =>
        GetSetupSectionProgressText(GetCodexReadyCount(_lastEnvironmentSnapshot), 4);

    public string SetupLocalAiProgressText =>
        GetSetupSectionProgressText(GetLocalAiReadyCount(_lastEnvironmentSnapshot), 4);

    public string SetupCoreNextStepText =>
        GetSetupCoreNextStepText(_lastEnvironmentSnapshot);

    public string SetupCodexNextStepText =>
        GetSetupCodexNextStepText(_lastEnvironmentSnapshot);

    public string SetupLocalAiNextStepText =>
        GetSetupLocalAiNextStepText(_lastEnvironmentSnapshot);

    public string SetupCoreSummaryBrush =>
        _lastEnvironmentSnapshot is null ? "#2D5366" : GetSetupSummaryBrush(GetCoreReadyCount(_lastEnvironmentSnapshot), 4);

    public string SetupCodexSummaryBrush =>
        _lastEnvironmentSnapshot is null ? "#2D5366" : GetSetupSummaryBrush(GetCodexReadyCount(_lastEnvironmentSnapshot), 4);

    public string SetupLocalAiSummaryBrush =>
        _lastEnvironmentSnapshot is null ? "#2D5366" : GetSetupSummaryBrush(GetLocalAiReadyCount(_lastEnvironmentSnapshot), 4);

    public string HardwareOverviewText
    {
        get
        {
            var snapshot = _lastEnvironmentSnapshot;

            if (snapshot is null)
            {
                return Strings["SetupHardwarePending"];
            }

            var gpuName = string.IsNullOrWhiteSpace(snapshot.GpuName)
                ? Strings["SetupHardwareGpuUnknown"]
                : snapshot.GpuName;
            var gpuMemory = snapshot.GpuMemoryBytes > 0
                ? FormatByteSize(snapshot.GpuMemoryBytes)
                : Strings["SetupHardwareMemoryUnknown"];

            return Strings.Format(
                "SetupHardwareOverviewFormat",
                gpuName,
                gpuMemory,
                FormatByteSize(snapshot.TotalPhysicalMemoryBytes),
                FormatByteSize(snapshot.SystemDriveFreeBytes));
        }
    }

    public string HardwareRecommendationText =>
        GetHardwareRecommendationText(_lastEnvironmentSnapshot);

    public string HardwareStatusBrush =>
        GetHardwareStatusBrush(_lastEnvironmentSnapshot);

    public string LocalAiStorageSummaryText
    {
        get
        {
            var snapshot = _lastEnvironmentSnapshot;

            if (snapshot is null)
            {
                return Strings["SetupHardwarePending"];
            }

            return Strings.Format(
                "SetupLocalAiStorageSummaryFormat",
                snapshot.OllamaModelCount,
                FormatByteSize(snapshot.OllamaModelStorageBytes),
                FormatByteSize(snapshot.SystemDriveFreeBytes));
        }
    }

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

    public bool SettingsPhotoPasteFixEnabled
    {
        get => _settingsPhotoPasteFixEnabled;
        set
        {
            if (SetField(ref _settingsPhotoPasteFixEnabled, value))
            {
                OnPropertyChanged(nameof(SettingsPhotoPasteFixStateText));
                OnPropertyChanged(nameof(SettingsPhotoPasteFixStateBrush));
                OnPropertyChanged(nameof(SettingsPhotoPasteFixStateForeground));
            }
        }
    }

    public string SettingsPhotoPasteFixStateText =>
        SettingsPhotoPasteFixEnabled
            ? Strings["SettingsPhotoPasteFixStateEnabled"]
            : Strings["SettingsPhotoPasteFixStateDisabled"];

    public string SettingsPhotoPasteFixStateBrush =>
        SettingsPhotoPasteFixEnabled ? "#DDF5E8" : "#FFF1E6";

    public string SettingsPhotoPasteFixStateForeground =>
        SettingsPhotoPasteFixEnabled ? "#1F7A52" : "#B86E10";

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
                HomeLaunchStatusText = string.Empty;
                OnPropertyChanged(nameof(CanLaunchNewSession));
                OnPropertyChanged(nameof(CanStartHomeSession));
                OnPropertyChanged(nameof(HomeStartHelpText));
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
                NotifyNewSessionAccessSummaryChanged();
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
                OnPropertyChanged(nameof(NewSessionDangerousWarningVisibility));
                NotifyNewSessionAccessSummaryChanged();
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
                OnPropertyChanged(nameof(NewSessionDangerousWarningVisibility));
                NotifyNewSessionAccessSummaryChanged();
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
                OnPropertyChanged(nameof(NewSessionDataRouteText));
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
                OnPropertyChanged(nameof(NewSessionDataRouteText));
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
                OnPropertyChanged(nameof(NewSessionDataRouteText));
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

    public Visibility NewSessionDangerousWarningVisibility =>
        ShouldUseDangerousBypassForNewSession() ? Visibility.Visible : Visibility.Collapsed;

    public string NewSessionAccessSummaryTitle =>
        GetNewSessionAccessLevel() switch
        {
            NewSessionAccessLevel.Critical => Strings["NewSessionAccessCriticalTitle"],
            NewSessionAccessLevel.Caution => Strings["NewSessionAccessCautionTitle"],
            _ => Strings["NewSessionAccessSafeTitle"]
        };

    public string NewSessionAccessSummaryText
    {
        get
        {
            var folder = string.IsNullOrWhiteSpace(NewSessionWorkingDirectory)
                ? Strings["NewSessionAccessFolderNotSelected"]
                : NewSessionWorkingDirectory.Trim();

            if (string.Equals(SelectedSandboxMode, "read-only", StringComparison.OrdinalIgnoreCase))
            {
                return Strings.Format("NewSessionAccessReadOnlyText", folder);
            }

            if (string.Equals(SelectedSandboxMode, "danger-full-access", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(SelectedApprovalPolicy, "never", StringComparison.OrdinalIgnoreCase)
                    ? Strings["NewSessionAccessFullNeverText"]
                    : Strings["NewSessionAccessFullConfirmText"];
            }

            return string.Equals(SelectedApprovalPolicy, "never", StringComparison.OrdinalIgnoreCase)
                ? Strings.Format("NewSessionAccessWorkspaceNeverText", folder)
                : Strings.Format("NewSessionAccessWorkspaceConfirmText", folder);
        }
    }

    public string NewSessionAccessSummaryBackground =>
        GetNewSessionAccessLevel() switch
        {
            NewSessionAccessLevel.Critical => "#FDECEC",
            NewSessionAccessLevel.Caution => "#FFF3E0",
            _ => "#E7F6EE"
        };

    public string NewSessionAccessSummaryForeground =>
        GetNewSessionAccessLevel() switch
        {
            NewSessionAccessLevel.Critical => "#B42318",
            NewSessionAccessLevel.Caution => "#7A4208",
            _ => "#1F6F4A"
        };

    public string NewSessionAccessSummaryBorder =>
        GetNewSessionAccessLevel() switch
        {
            NewSessionAccessLevel.Critical => "#F04438",
            NewSessionAccessLevel.Caution => "#EAAA08",
            _ => "#32A66A"
        };

    public string NewSessionDataRouteText
    {
        get
        {
            if (NewSessionUseOss)
            {
                var provider = LocalProviderOptions
                    .FirstOrDefault(option => option.Value == SelectedLocalProvider)
                    ?.DisplayName;
                provider = string.IsNullOrWhiteSpace(provider)
                    ? Strings["NewSessionDataRouteLocalProvider"]
                    : provider;

                return NewSessionUseSearch
                    ? Strings.Format("NewSessionDataRouteLocalWithWeb", provider)
                    : Strings.Format("NewSessionDataRouteLocal", provider);
            }

            return NewSessionUseSearch
                ? Strings["NewSessionDataRouteCloudWithWeb"]
                : Strings["NewSessionDataRouteCloud"];
        }
    }

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
                ScheduleSearchFilter();
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

    public AiExtensionItem? SelectedExtension
    {
        get => _selectedExtension;
        set
        {
            if (SetField(ref _selectedExtension, value))
            {
                PopulateExtensionEditor(value);
                OnPropertyChanged(nameof(CanDeleteSelectedExtension));
                OnPropertyChanged(nameof(CanOpenSelectedExtensionLocation));
                OnPropertyChanged(nameof(CanInstallSelectedExtension));
                OnPropertyChanged(nameof(CanEnableSelectedExtension));
                OnPropertyChanged(nameof(CanDisableSelectedExtension));
                OnPropertyChanged(nameof(CanRemoveSelectedExtension));
                OnPropertyChanged(nameof(CanSaveSelectedExtension));
                OnPropertyChanged(nameof(SelectedExtensionDetailsText));
                OnPropertyChanged(nameof(SelectedExtensionPrimaryActionText));
            }
        }
    }

    public string ExtensionName
    {
        get => _extensionName;
        set
        {
            if (SetField(ref _extensionName, value))
            {
                OnPropertyChanged(nameof(CanSaveSelectedExtension));
            }
        }
    }

    public string SelectedExtensionKind
    {
        get => _selectedExtensionKind;
        set => SetField(ref _selectedExtensionKind, value);
    }

    public string SelectedExtensionTarget
    {
        get => _selectedExtensionTarget;
        set
        {
            if (SetField(ref _selectedExtensionTarget, value))
            {
                RefreshExtensionViews(SelectedExtension?.Id);
            }
        }
    }

    public string ExtensionTargetApp
    {
        get => _extensionTargetApp;
        set => SetField(ref _extensionTargetApp, string.IsNullOrWhiteSpace(value) ? "Codex" : value);
    }

    public string ExtensionCommandOrUri
    {
        get => _extensionCommandOrUri;
        set => SetField(ref _extensionCommandOrUri, value);
    }

    public string ExtensionDescription
    {
        get => _extensionDescription;
        set => SetField(ref _extensionDescription, value);
    }

    public string ExtensionSearchText
    {
        get => _extensionSearchText;
        set
        {
            if (SetField(ref _extensionSearchText, value))
            {
                RefreshExtensionViews(SelectedExtension?.Id);
            }
        }
    }

    public bool ExtensionIsEnabled
    {
        get => _extensionIsEnabled;
        set => SetField(ref _extensionIsEnabled, value);
    }

    public string ExtensionStatusText
    {
        get => _extensionStatusText;
        private set => SetField(ref _extensionStatusText, value);
    }

    public string ExtensionStatusForeground
    {
        get => _extensionStatusForeground;
        private set => SetField(ref _extensionStatusForeground, value);
    }

    public bool CanDeleteSelectedExtension =>
        SelectedExtension is not null &&
        (SelectedExtension.IsCustom || TryGetExtensionFileSystemTarget(SelectedExtension, forDelete: true, out _, out _));

    public bool CanOpenSelectedExtensionLocation =>
        SelectedExtension is not null &&
        TryGetExtensionFileSystemTarget(SelectedExtension, forDelete: false, out _, out _);

    public bool CanInstallSelectedExtension =>
        SelectedExtension is { CanProvision: true, IsBusy: false } item &&
        (!item.IsInstalled || string.Equals(item.ManagementKind, "endpoint", StringComparison.OrdinalIgnoreCase));

    public bool CanEnableSelectedExtension =>
        SelectedExtension is { IsCustom: true, IsInstalled: true, IsEnabled: false, IsBusy: false };

    public bool CanDisableSelectedExtension =>
        SelectedExtension is { IsCustom: true, IsInstalled: true, IsEnabled: true, IsBusy: false };

    public bool CanRemoveSelectedExtension =>
        SelectedExtension is { IsBusy: false } item &&
        !item.IsDetected &&
        ((item.CanUninstall && item.IsInstalled) || item.IsCustom);

    public string SelectedExtensionPrimaryActionText =>
        string.Equals(SelectedExtension?.ManagementKind, "endpoint", StringComparison.OrdinalIgnoreCase)
            ? Strings["ExtensionsCheckConnectionButton"]
            : Strings["ExtensionsInstallButton"];

    public bool CanSaveSelectedExtension => !string.IsNullOrWhiteSpace(ExtensionName);

    public string ExtensionsStoragePath => _extensionCatalogService.GetStoragePath();

    public ObservableCollection<AiExtensionItem> DisplayedAiExtensions =>
        _showCustomExtensionsTab
            ? CustomAiExtensions
            : _showInstalledExtensionsTab
                ? InstalledAiExtensions
                : SuggestedAiExtensions;

    public string SuggestedExtensionsTabText =>
        Strings.Format("ExtensionsSuggestedTab", SuggestedAiExtensions.Count);

    public string InstalledExtensionsTabText =>
        Strings.Format("ExtensionsInstalledTab", InstalledAiExtensions.Count);

    public string CustomExtensionsTabText =>
        Strings.Format("ExtensionsCustomTab", CustomAiExtensions.Count);

    public string SuggestedExtensionsTabButtonBackground =>
        _showInstalledExtensionsTab || _showCustomExtensionsTab ? "#E8D8C8" : "#16212B";

    public string SuggestedExtensionsTabButtonForeground =>
        _showInstalledExtensionsTab || _showCustomExtensionsTab ? "#16212B" : "#FFFDF9";

    public string InstalledExtensionsTabButtonBackground =>
        _showInstalledExtensionsTab ? "#16212B" : "#E8D8C8";

    public string InstalledExtensionsTabButtonForeground =>
        _showInstalledExtensionsTab ? "#FFFDF9" : "#16212B";

    public string CustomExtensionsTabButtonBackground =>
        _showCustomExtensionsTab ? "#16212B" : "#E8D8C8";

    public string CustomExtensionsTabButtonForeground =>
        _showCustomExtensionsTab ? "#FFFDF9" : "#16212B";

    public string AllExtensionsTargetTabText => Strings["ExtensionsTargetAllTab"];

    public string CodexExtensionsTargetTabText => Strings["ExtensionsTargetCodexTab"];

    public string OpenCodeExtensionsTargetTabText => Strings["ExtensionsTargetOpenCodeTab"];

    public string LmStudioExtensionsTargetTabText => Strings["ExtensionsTargetLmStudioTab"];

    public string AllExtensionsTargetTabBackground =>
        SelectedExtensionTarget == "All" ? "#16212B" : "#E8D8C8";

    public string AllExtensionsTargetTabForeground =>
        SelectedExtensionTarget == "All" ? "#FFFDF9" : "#16212B";

    public string CodexExtensionsTargetTabBackground =>
        SelectedExtensionTarget == "Codex" ? "#16212B" : "#E8D8C8";

    public string CodexExtensionsTargetTabForeground =>
        SelectedExtensionTarget == "Codex" ? "#FFFDF9" : "#16212B";

    public string OpenCodeExtensionsTargetTabBackground =>
        SelectedExtensionTarget == "OpenCode" ? "#16212B" : "#E8D8C8";

    public string OpenCodeExtensionsTargetTabForeground =>
        SelectedExtensionTarget == "OpenCode" ? "#FFFDF9" : "#16212B";

    public string LmStudioExtensionsTargetTabBackground =>
        SelectedExtensionTarget == "LmStudio" ? "#16212B" : "#E8D8C8";

    public string LmStudioExtensionsTargetTabForeground =>
        SelectedExtensionTarget == "LmStudio" ? "#FFFDF9" : "#16212B";

    public string SelectedExtensionDetailsText
    {
        get
        {
            if (SelectedExtension is null)
            {
                return Strings["ExtensionsNoSelection"];
            }

            var lines = new List<string>
            {
                Strings.Format("ExtensionsDetailName", SelectedExtension.Name),
                Strings.Format("ExtensionsDetailType", SelectedExtension.KindLabel),
                Strings.Format("ExtensionsDetailTarget", SelectedExtension.TargetAppDisplayLabel),
                Strings.Format("ExtensionsDetailSource", SelectedExtension.SourceDisplayLabel),
                Strings.Format("ExtensionsDetailStatus", SelectedExtension.InstallStateLabel)
            };

            if (!string.IsNullOrWhiteSpace(SelectedExtension.PackageVersion))
            {
                lines.Add(Strings.Format("ExtensionsDetailVersion", SelectedExtension.PackageVersion));
            }

            if (!string.IsNullOrWhiteSpace(SelectedExtension.RequestedAccess))
            {
                lines.Add(Strings.Format("ExtensionsDetailAccess", SelectedExtension.RequestedAccess));
            }

            lines.Add(Strings.Format("ExtensionsDetailCommand", SelectedExtension.CommandOrUri));

            if (!string.IsNullOrWhiteSpace(SelectedExtension.VerificationDetail))
            {
                lines.Add(Strings.Format("ExtensionsDetailVerification", SelectedExtension.VerificationDetail));
            }

            lines.Add(SelectedExtension.Description);
            return string.Join(Environment.NewLine, lines);
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
                QueueSelectedSessionTranscriptLoad(value);
                OnPropertyChanged(nameof(CanOpenSelectedFile));
                OnPropertyChanged(nameof(CanOpenSelectedSessionDirectory));
                OnPropertyChanged(nameof(CanDeleteSelectedSession));
                OnPropertyChanged(nameof(CanArchiveSelectedSession));
                OnPropertyChanged(nameof(CanToggleSelectedSessionHidden));
                OnPropertyChanged(nameof(CanResumeSelectedSession));
                OnPropertyChanged(nameof(CanResumeSelectedSessionInOpenCode));
                OnPropertyChanged(nameof(CanRefreshSelectedSessionOpenCodeBridge));
                OnPropertyChanged(nameof(CanUseSelectedSessionDirectory));
                OnPropertyChanged(nameof(CanCreateSelectedSessionCheckpoint));
                OnPropertyChanged(nameof(CanEditSelectedSessionNote));
                OnPropertyChanged(nameof(CanSaveSelectedSessionNote));
                OnPropertyChanged(nameof(CanClearSelectedSessionNote));
                OnPropertyChanged(nameof(FavoriteButtonText));
                OnPropertyChanged(nameof(OpenCodeResumeButtonText));
                OnPropertyChanged(nameof(SelectedSessionOpenCodeBridgeText));
                OnPropertyChanged(nameof(SelectedSessionTitleText));
                OnPropertyChanged(nameof(SelectedSessionPreviewText));
                OnPropertyChanged(nameof(SelectedSessionTranscriptText));
                OnPropertyChanged(nameof(SelectedSessionFavoriteText));
                OnPropertyChanged(nameof(SelectedSessionHealthTitle));
                OnPropertyChanged(nameof(SelectedSessionHealthText));
                OnPropertyChanged(nameof(SelectedSessionHealthBackground));
                OnPropertyChanged(nameof(SelectedSessionHealthForeground));
                OnPropertyChanged(nameof(SelectedSessionHealthBorder));
                OnPropertyChanged(nameof(ToggleSessionHiddenButtonText));
            }
        }
    }

    public string SelectedSessionTitleText =>
        SelectedSession?.DisplayTitle ?? Strings["NoSessionSelected"];

    public string SelectedSessionPreviewText =>
        SelectedSession?.Preview ?? Strings["SelectSessionHint"];

    public string SelectedSessionTranscriptText
    {
        get => _selectedSessionTranscriptText;
        private set => SetField(ref _selectedSessionTranscriptText, value);
    }

    public bool IsSelectedSessionTranscriptLoading
    {
        get => _selectedSessionTranscriptLoading;
        private set => SetField(ref _selectedSessionTranscriptLoading, value);
    }

    public string SelectedSessionFavoriteText =>
        SelectedSession is null ? "-" : SelectedSession.IsFavorite ? Strings["Yes"] : Strings["No"];

    public string SelectedSessionHealthTitle =>
        GetSelectedSessionHealthLevel() switch
        {
            SessionHealthLevel.Overloaded => Strings["SessionHealthOverloadedTitle"],
            SessionHealthLevel.Long => Strings["SessionHealthLongTitle"],
            _ => Strings["SessionHealthStableTitle"]
        };

    public string SelectedSessionHealthText =>
        GetSelectedSessionHealthLevel() switch
        {
            SessionHealthLevel.Overloaded => Strings["SessionHealthOverloadedText"],
            SessionHealthLevel.Long => Strings["SessionHealthLongText"],
            _ => Strings["SessionHealthStableText"]
        };

    public string SelectedSessionHealthBackground =>
        GetSelectedSessionHealthLevel() switch
        {
            SessionHealthLevel.Overloaded => "#3B2022",
            SessionHealthLevel.Long => "#3B3220",
            _ => "#17392D"
        };

    public string SelectedSessionHealthForeground =>
        GetSelectedSessionHealthLevel() switch
        {
            SessionHealthLevel.Overloaded => "#FFB4AB",
            SessionHealthLevel.Long => "#FFE08A",
            _ => "#A9EBC9"
        };

    public string SelectedSessionHealthBorder =>
        GetSelectedSessionHealthLevel() switch
        {
            SessionHealthLevel.Overloaded => "#F97066",
            SessionHealthLevel.Long => "#EAAA08",
            _ => "#32A66A"
        };

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

    public int FavoriteSessions => _allSessions.Count(session => session.IsFavorite && !session.IsHidden);

    public int RegularSessions => _allSessions.Count(session => !session.IsFavorite && !session.IsHidden);

    public int HiddenSessions => _allSessions.Count(session => session.IsHidden);

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
        LogStartupPhase("MainWindow loaded.");
        UpdateRefreshTimer();
        UpdateSetupRefreshTimer();
        EnsureSessionWatchersInitialized();

        if (_startupRefreshScheduled)
        {
            return;
        }

        _startupRefreshScheduled = true;
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);
        LogStartupPhase("First frame yielded to background.");
        _ = RunInitialRefreshAsync();
        _ = RunDeferredStartupInitializationAsync();
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        FitToWorkArea();
        RefreshAdaptiveLayoutBindings();
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ScheduleAdaptiveLayoutRefresh();
    }

    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        ScheduleAdaptiveLayoutRefresh();
    }

    private void MainWindow_Activated(object? sender, EventArgs e)
    {
        UpdateRefreshTimer();
        UpdateSetupRefreshTimer();

        if (IsLoaded && SelectedAppSection == AppSection.Sessions)
        {
            _ = RefreshSessionsAsync(isAutomaticRefresh: false);
        }

        if (IsLoaded && SelectedAppSection == AppSection.Setup)
        {
            _ = RefreshSetupSectionAsync(preserveDnsStatus: true);
        }
        else if (IsLoaded && IsSetupRefreshBoostActive)
        {
            _ = RefreshSetupSectionAsync(preserveDnsStatus: true);
        }

        if (IsLoaded && SelectedAppSection == AppSection.Settings)
        {
            _ = RefreshSettingsSectionAsync();
        }
    }

    private void MainWindow_Deactivated(object? sender, EventArgs e)
    {
        UpdateRefreshTimer();
        UpdateSetupRefreshTimer();
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        PersistSelectedSessionNote(showStatus: false, refreshFilter: false);
        _refreshTimer.Stop();
        _searchDebounceTimer.Stop();
        _setupRefreshTimer.Stop();
        _layoutRefreshTimer.Stop();
        DisposeSessionWatchers();
        _photoPasteFixService.Dispose();
    }

    private async Task RunInitialRefreshAsync()
    {
        LogStartupPhase("Initial refresh started.");
        await RefreshSessionsAsync(isAutomaticRefresh: false, forceRefresh: true);
        LogStartupPhase("Initial sessions refresh finished.");
        EnsureSessionsSurfaceInitialized();

        if (SelectedAppSection is AppSection.Home or AppSection.Setup)
        {
            EnsureSectionDataInitialized(AppSection.Setup);
            await RefreshSetupSectionAsync(preserveDnsStatus: true, forceRefresh: true);
            LogStartupPhase("Initial setup refresh finished.");
        }
        else if (SelectedAppSection == AppSection.Settings && _lastAppUpdateSnapshot is null)
        {
            await RefreshSettingsSectionAsync(forceRefresh: true);
            LogStartupPhase("Initial settings refresh finished.");
        }
    }

    private async Task RunDeferredStartupInitializationAsync()
    {
        await Task.Delay(150);

        try
        {
            await Task.Run(() => _photoPasteFixService.UpdateConfiguration(_settingsPhotoPasteFixEnabled));
            LogStartupPhase("Photo paste fix configuration applied.");
        }
        catch (Exception exception)
        {
            _settingsPhotoPasteFixEnabled = false;
            _settingsService.SavePhotoPasteFixEnabled(false);
            _logService.Error(nameof(MainWindow), "Failed to enable the photo paste fix on deferred startup.", exception);
        }
    }

    private void EnsureSectionDataInitialized(AppSection section)
    {
        switch (section)
        {
            case AppSection.NewSession when !_isNewSessionSectionInitialized:
                _isNewSessionSectionInitialized = true;
                OnPropertyChanged(nameof(NewSessionSectionContent));
                RefreshLaunchOptionCollections();
                LoadNewSessionConfigurationInfoSafe();
                ApplyDangerousAccessDefaultsToNewSession();
                LogStartupPhase("New Session section initialized.");
                break;
            case AppSection.Extensions when !_isExtensionsSectionInitialized:
                _isExtensionsSectionInitialized = true;
                OnPropertyChanged(nameof(ExtensionsSectionContent));
                RefreshExtensionKindOptions();
                RefreshExtensionTargetOptions();
                LoadExtensionsSafe();
                LogStartupPhase("Extensions section initialized.");
                break;
            case AppSection.Setup when !_isSetupSectionInitialized:
                _isSetupSectionInitialized = true;
                OnPropertyChanged(nameof(SetupSectionContent));
                RefreshLocalAiModelOptions();
                RefreshCreativeAiToolOptions();
                RefreshAiAgentToolOptions();
                RefreshOpenClawSetupModes();
                RefreshOpenClawCapabilityChecks();
                LoadDnsPresetsSafe();
                LogStartupPhase("Setup section initialized.");
                break;
            case AppSection.Settings when !_isSettingsSectionInitialized:
                _isSettingsSectionInitialized = true;
                OnPropertyChanged(nameof(SettingsSectionContent));
                LogStartupPhase("Settings section initialized.");
                break;
        }
    }

    private void EnsureSessionsSurfaceInitialized()
    {
        if (_isSessionsSurfaceInitialized)
        {
            return;
        }

        _isSessionsSurfaceInitialized = true;
        OnPropertyChanged(nameof(SessionsMainContent));
        OnPropertyChanged(nameof(SessionsMainPlaceholderVisibility));
        LogStartupPhase("Sessions surface initialized.");
    }

    private void LogStartupPhase(string phase)
    {
        _logService.Info(nameof(MainWindow), $"Startup +{_startupStopwatch.ElapsedMilliseconds} ms: {phase}");
    }

    private void LoadSessionMetadata()
    {
        _favoriteSessionIds = _favoritesService.LoadFavorites();
        _hiddenSessionIds = _sessionVisibilityService.LoadHiddenSessions();
        _sessionNotes = _notesService.LoadNotes();
    }

    private void LoadOpenCodeLinks()
    {
        _openCodeLinks = _openCodeLinkService.LoadLinks();
    }

    private void ScheduleSearchFilter()
    {
        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void QueueSelectedSessionTranscriptLoad(SessionRecord? session)
    {
        var loadVersion = ++_selectedSessionTranscriptLoadVersion;

        if (session is null)
        {
            IsSelectedSessionTranscriptLoading = false;
            SelectedSessionTranscriptText = Strings["NoTranscriptLoaded"];
            return;
        }

        if (!string.IsNullOrWhiteSpace(session.TranscriptText))
        {
            IsSelectedSessionTranscriptLoading = false;
            SelectedSessionTranscriptText = session.TranscriptText;
            return;
        }

        var language = SelectedLanguageOption?.Language ?? AppLanguage.English;
        IsSelectedSessionTranscriptLoading = true;
        SelectedSessionTranscriptText = Strings["TranscriptLoading"];

        _ = Task.Run(() =>
            {
                try
                {
                    return _sessionService.LoadTranscriptText(session, language);
                }
                catch (Exception exception)
                {
                    _logService.Error(nameof(MainWindow), "Failed to load transcript for the selected session.", exception);
                    return Strings["NoTranscriptLoaded"];
                }
            })
            .ContinueWith(
                task =>
                {
                    _ = Dispatcher.InvokeAsync(() =>
                    {
                        if (loadVersion != _selectedSessionTranscriptLoadVersion ||
                            !ReferenceEquals(SelectedSession, session))
                        {
                            return;
                        }

                        SelectedSessionTranscriptText = task.Result;
                        IsSelectedSessionTranscriptLoading = false;
                    });
                },
                TaskScheduler.Default);
    }

    private OpenCodeSessionLinkRecord? GetSelectedSessionOpenCodeLink()
    {
        return SelectedSession is not null &&
               _openCodeLinks.TryGetValue(SelectedSession.SessionId, out var linkRecord)
            ? linkRecord
            : null;
    }

    private bool IsSelectedSessionOpenCodeBridgeStale()
    {
        var session = SelectedSession;
        var linkRecord = GetSelectedSessionOpenCodeLink();

        return session is not null &&
               linkRecord is not null &&
               IsOpenCodeLinkStale(session, linkRecord);
    }

    private static bool IsOpenCodeLinkStale(SessionRecord session, OpenCodeSessionLinkRecord linkRecord)
    {
        return session.UpdatedAtUtc > linkRecord.CodexUpdatedAtUtc ||
               string.IsNullOrWhiteSpace(linkRecord.HandoffPath) ||
               !File.Exists(linkRecord.HandoffPath) ||
               AiHelperWorkspaceService.IsUnsafeWorkspace(linkRecord.WorkingDirectory);
    }

    private string BuildSelectedSessionOpenCodeBridgeText()
    {
        var session = SelectedSession;
        var issue = GetOpenCodeBridgeIssue(session);

        if (!string.IsNullOrWhiteSpace(issue))
        {
            return issue;
        }

        var linkRecord = GetSelectedSessionOpenCodeLink();
        if (linkRecord is null)
        {
            return Strings["OpenCodeBridgeMissing"];
        }

        var shortId = linkRecord.OpenCodeSessionId.Length <= 18
            ? linkRecord.OpenCodeSessionId
            : linkRecord.OpenCodeSessionId[..18];

        return IsSelectedSessionOpenCodeBridgeStale()
            ? Strings.Format("OpenCodeBridgeLinkedStale", shortId)
            : Strings.Format("OpenCodeBridgeLinkedReady", shortId);
    }

    private string? GetOpenCodeBridgeIssue(SessionRecord? session)
    {
        if (session is null)
        {
            return Strings["OpenCodeBridgeNoSelection"];
        }

        if (!_openCodeBridgeService.IsOpenCodeDesktopAvailable)
        {
            return Strings["OpenCodeBridgeNotInstalled"];
        }

        if (!File.Exists(session.FilePath) || (session.UserMessageCount + session.AssistantMessageCount) <= 0)
        {
            return Strings["OpenCodeBridgeSessionUnavailable"];
        }

        return null;
    }

    private void RefreshOpenCodeBindings()
    {
        OnPropertyChanged(nameof(CanResumeSelectedSessionInOpenCode));
        OnPropertyChanged(nameof(CanRefreshSelectedSessionOpenCodeBridge));
        OnPropertyChanged(nameof(OpenCodeResumeButtonText));
        OnPropertyChanged(nameof(SelectedSessionOpenCodeBridgeText));
    }

    private async Task<OpenCodeSessionLinkRecord> EnsureOpenCodeLinkAsync(
        SessionRecord selectedSession,
        bool forceRefresh)
    {
        if (!forceRefresh &&
            _openCodeLinks.TryGetValue(selectedSession.SessionId, out var existingLink) &&
            !IsOpenCodeLinkStale(selectedSession, existingLink))
        {
            RefreshOpenCodeBindings();
            return existingLink;
        }

        var conversation = await Task.Run(() => _sessionService.GetConversation(selectedSession));
        var linkRecord = await Task.Run(() => _openCodeBridgeService.CreateBridge(conversation));
        _openCodeLinks[selectedSession.SessionId] = linkRecord;
        _openCodeLinkService.SaveLinks(_openCodeLinks);
        RefreshOpenCodeBindings();
        return linkRecord;
    }

    private void EnsureSessionWatchersInitialized()
    {
        if (_sessionFolderWatcher is null)
        {
            var sessionsFolder = _environmentService.SessionsFolder;

            if (Directory.Exists(sessionsFolder))
            {
                _sessionFolderWatcher = CreateSessionWatcher(
                    sessionsFolder,
                    "*.jsonl",
                    includeSubdirectories: true);
            }
        }

        if (_sessionIndexWatcher is null)
        {
            var codexHomeFolder = _environmentService.CodexHomeFolder;

            if (Directory.Exists(codexHomeFolder))
            {
                _sessionIndexWatcher = CreateSessionWatcher(
                    codexHomeFolder,
                    "session_index.jsonl",
                    includeSubdirectories: false);
            }
        }
    }

    private FileSystemWatcher CreateSessionWatcher(string path, string filter, bool includeSubdirectories)
    {
        var watcher = new FileSystemWatcher(path, filter)
        {
            IncludeSubdirectories = includeSubdirectories,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
            EnableRaisingEvents = true
        };

        watcher.Changed += SessionWatcher_Changed;
        watcher.Created += SessionWatcher_Changed;
        watcher.Deleted += SessionWatcher_Changed;
        watcher.Renamed += SessionWatcher_Renamed;
        watcher.Error += SessionWatcher_Error;

        return watcher;
    }

    private void SessionWatcher_Changed(object sender, FileSystemEventArgs e)
    {
        MarkSessionsRefreshPending();
    }

    private void SessionWatcher_Renamed(object sender, RenamedEventArgs e)
    {
        MarkSessionsRefreshPending();
    }

    private void SessionWatcher_Error(object sender, ErrorEventArgs e)
    {
        _logService.Error(nameof(MainWindow), "The session watcher reported an error.", e.GetException());

        if (ReferenceEquals(sender, _sessionFolderWatcher))
        {
            DisposeSessionWatcher(ref _sessionFolderWatcher);
        }
        else if (ReferenceEquals(sender, _sessionIndexWatcher))
        {
            DisposeSessionWatcher(ref _sessionIndexWatcher);
        }

        MarkSessionsRefreshPending();
        _ = Dispatcher.BeginInvoke(EnsureSessionWatchersInitialized, DispatcherPriority.Background);
    }

    private void MarkSessionsRefreshPending()
    {
        _sessionRefreshPending = true;
    }

    private void MarkSetupRefreshPending()
    {
        _setupRefreshPending = true;
    }

    private void DisposeSessionWatchers()
    {
        DisposeSessionWatcher(ref _sessionFolderWatcher);
        DisposeSessionWatcher(ref _sessionIndexWatcher);
    }

    private static void DisposeSessionWatcher(ref FileSystemWatcher? watcher)
    {
        watcher?.Dispose();
        watcher = null;
    }

    private static void RouteMouseWheelToScrollableParent(object sender, MouseWheelEventArgs e)
    {
        if (e.OriginalSource is not DependencyObject source)
        {
            return;
        }

        var comboBox = FindVisualParent<ComboBox>(source);
        if (comboBox?.IsDropDownOpen == true)
        {
            return;
        }

        var scrollViewer = FindWheelScrollViewer(source, e.Delta);
        if (scrollViewer is null)
        {
            return;
        }

        ScrollByWheel(scrollViewer, e.Delta);
        e.Handled = true;
    }

    private static ScrollViewer? FindWheelScrollViewer(DependencyObject source, int delta)
    {
        for (DependencyObject? current = source; current is not null; current = GetVisualOrLogicalParent(current))
        {
            if (current is ScrollViewer scrollViewer && CanScrollByWheel(scrollViewer, delta))
            {
                return scrollViewer;
            }
        }

        return null;
    }

    private static bool CanScrollByWheel(ScrollViewer scrollViewer, int delta)
    {
        if (scrollViewer.ScrollableHeight <= 0)
        {
            return false;
        }

        return delta < 0
            ? scrollViewer.VerticalOffset < scrollViewer.ScrollableHeight
            : scrollViewer.VerticalOffset > 0;
    }

    private static void ScrollByWheel(ScrollViewer scrollViewer, int delta)
    {
        var lines = SystemParameters.WheelScrollLines <= 0
            ? 3
            : Math.Min(SystemParameters.WheelScrollLines, 12);

        for (var index = 0; index < lines; index++)
        {
            if (delta < 0)
            {
                scrollViewer.LineDown();
            }
            else
            {
                scrollViewer.LineUp();
            }
        }
    }

    private static T? FindVisualParent<T>(DependencyObject source)
        where T : DependencyObject
    {
        for (DependencyObject? current = source; current is not null; current = GetVisualOrLogicalParent(current))
        {
            if (current is T typed)
            {
                return typed;
            }
        }

        return null;
    }

    private static DependencyObject? GetVisualOrLogicalParent(DependencyObject source)
    {
        if (source is Visual or Visual3D)
        {
            var visualParent = VisualTreeHelper.GetParent(source);
            if (visualParent is not null)
            {
                return visualParent;
            }
        }

        return LogicalTreeHelper.GetParent(source);
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

    private async void CreateSelectedSessionCheckpointButton_Click(object sender, RoutedEventArgs e)
    {
        var checkpointPath = await CreateSelectedSessionCheckpointAsync(showSuccessStatus: true);
        if (!string.IsNullOrWhiteSpace(checkpointPath))
        {
            OpenExplorerSelect(checkpointPath);
        }
    }

    private async void StartFreshSessionFromCheckpointButton_Click(object sender, RoutedEventArgs e)
    {
        var session = SelectedSession;
        if (session is null)
        {
            return;
        }

        var checkpointPath = await CreateSelectedSessionCheckpointAsync(showSuccessStatus: false);
        if (string.IsNullOrWhiteSpace(checkpointPath))
        {
            return;
        }

        EnsureSectionDataInitialized(AppSection.NewSession);
        NewSessionPrompt = Strings.Format("SessionCheckpointContinuePrompt", checkpointPath);
        if (Directory.Exists(session.WorkingDirectory))
        {
            NewSessionWorkingDirectory = AiHelperWorkspaceService.ResolveSafeWorkspace(
                session.WorkingDirectory,
                session.SessionId,
                session.Title,
                out _);
        }

        SelectedSandboxMode = "workspace-write";
        SelectedApprovalPolicy = "on-request";
        SelectedAppSection = AppSection.NewSession;
        SetNewSessionStatus(
            "#1F6F4A",
            Strings.Format("SessionCheckpointReadyForFreshSession", checkpointPath));
    }

    private async Task<string?> CreateSelectedSessionCheckpointAsync(bool showSuccessStatus)
    {
        var session = SelectedSession;
        if (session is null)
        {
            return null;
        }

        try
        {
            var language = Strings.CurrentLanguage;
            var transcript = !string.IsNullOrWhiteSpace(session.TranscriptText)
                ? session.TranscriptText
                : await Task.Run(() => _sessionService.LoadTranscriptText(session, language));
            var path = await Task.Run(
                () => _checkpointService.CreateCheckpoint(session, transcript, language));

            if (showSuccessStatus)
            {
                SetStatus("#A9EBC9", "SessionCheckpointCreated", path);
            }

            return path;
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to create a session checkpoint.", exception);
            SetStatus("#FFD6D6", "SessionCheckpointFailed", exception.Message);
            return null;
        }
    }

    private void OpenSelectedSessionDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null ||
            string.IsNullOrWhiteSpace(selectedSession.WorkingDirectory) ||
            selectedSession.WorkingDirectory == "-" ||
            !Directory.Exists(selectedSession.WorkingDirectory))
        {
            return;
        }

        CodexEnvironmentService.OpenFolder(selectedSession.WorkingDirectory);
        SetStatus("#F8E7D6", "StatusSessionFolderOpened", selectedSession.WorkingDirectory);
    }

    private void CopySelectedSessionIdButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null || string.IsNullOrWhiteSpace(selectedSession.SessionId))
        {
            return;
        }

        Clipboard.SetText(selectedSession.SessionId);
        SetStatus("#F8E7D6", "StatusSessionIdCopied");
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
            _hiddenSessionIds.Remove(selectedSession.SessionId);
            _openCodeLinks.Remove(selectedSession.SessionId);
            _sessionNotes.Remove(selectedSession.SessionId);
            _favoritesService.SaveFavorites(_favoriteSessionIds);
            _sessionVisibilityService.SaveHiddenSessions(_hiddenSessionIds);
            _openCodeLinkService.SaveLinks(_openCodeLinks);
            _notesService.SaveNotes(_sessionNotes);
            await RefreshSessionsAsync(isAutomaticRefresh: false, forceRefresh: true);
        }
        catch (Exception exception)
        {
            SetStatus("#FFD6D6", "StatusDeleteFailed", exception.Message);
        }
    }

    private void ToggleSelectedSessionHiddenButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null)
        {
            return;
        }

        if (selectedSession.IsHidden)
        {
            _hiddenSessionIds.Remove(selectedSession.SessionId);
            selectedSession.IsHidden = false;
            SetStatus("#F8E7D6", "StatusSessionShown", selectedSession.Title);
        }
        else
        {
            _hiddenSessionIds.Add(selectedSession.SessionId);
            selectedSession.IsHidden = true;
            SetStatus("#F8E7D6", "StatusSessionHidden", selectedSession.Title);
        }

        _sessionVisibilityService.SaveHiddenSessions(_hiddenSessionIds);
        RefreshSessionCountBindings();
        OnPropertyChanged(nameof(ToggleSessionHiddenButtonText));
        ExportSessionsFeedSafe();
        ApplyFilter(selectedSession.SessionId);
    }

    private async void ArchiveSelectedSessionButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        if (selectedSession is null || !File.Exists(selectedSession.FilePath))
        {
            return;
        }

        var result = MessageBox.Show(
            Strings.Format("ArchiveSessionDialogMessage", selectedSession.Title, selectedSession.SessionId),
            Strings["ArchiveSessionDialogTitle"],
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var archivePath = _sessionService.ArchiveSession(selectedSession);
            _favoriteSessionIds.Remove(selectedSession.SessionId);
            _hiddenSessionIds.Remove(selectedSession.SessionId);
            _openCodeLinks.Remove(selectedSession.SessionId);
            _sessionNotes.Remove(selectedSession.SessionId);
            _favoritesService.SaveFavorites(_favoriteSessionIds);
            _sessionVisibilityService.SaveHiddenSessions(_hiddenSessionIds);
            _openCodeLinkService.SaveLinks(_openCodeLinks);
            _notesService.SaveNotes(_sessionNotes);
            SetStatus("#F8E7D6", "StatusSessionArchived", archivePath);
            await RefreshSessionsAsync(isAutomaticRefresh: false, forceRefresh: true);
        }
        catch (Exception exception)
        {
            SetStatus("#FFD6D6", "StatusSessionArchiveFailed", exception.Message);
        }
    }

    private void OpenSessionArchiveButton_Click(object sender, RoutedEventArgs e)
    {
        var archivePath = _sessionService.GetSessionArchiveRootPath();
        Directory.CreateDirectory(archivePath);
        CodexEnvironmentService.OpenFolder(archivePath);
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

        if (selectedSession.IsClaudeSession)
        {
            if (!File.Exists(_environmentService.ClaudeCommandPath))
            {
                SetStatus("#FFD6D6", "StatusClaudeCmdMissing", _environmentService.ClaudeCommandPath);
                return;
            }

            MarkSessionsRefreshPending();
            _environmentService.LaunchClaudeResumeSession(
                selectedSession.SessionId,
                selectedSession.WorkingDirectory);
            SetStatus("#F8E7D6", "StatusResumeStarted");
            return;
        }

        if (!File.Exists(_environmentService.CodexCommandPath))
        {
            SetStatus("#FFD6D6", "StatusCodexCmdMissing", _environmentService.CodexCommandPath);
            return;
        }

        var workingDirectory = AiHelperWorkspaceService.ResolveSafeWorkspace(
            selectedSession.WorkingDirectory,
            selectedSession.SessionId,
            selectedSession.Title,
            out _);

        MarkSessionsRefreshPending();
        _environmentService.LaunchResumeSession(selectedSession.SessionId, workingDirectory);
        SetStatus("#F8E7D6", "StatusResumeStarted");
    }

    private async void ResumeSelectedSessionInOpenCodeButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        _logService.Info(
            nameof(MainWindow),
            $"OpenCode bridge click. Session={selectedSession?.SessionId ?? "-"}; Action=ResumeOrCreate.");

        if (selectedSession is null)
        {
            return;
        }

        if (!selectedSession.IsCodexSession)
        {
            SetStatus("#FFD6D6", "StatusResumeOnlyCodex");
            return;
        }

        var issue = GetOpenCodeBridgeIssue(selectedSession);
        if (!string.IsNullOrWhiteSpace(issue))
        {
            SetStatus("#FFD6D6", "StatusOpenCodeBridgeFailed", issue);
            MessageBox.Show(
                issue,
                Strings["DetailOpenCodeBridge"],
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            IsOpenCodeBusy = true;
            SetStatus("#F8E7D6", "StatusOpenCodeBridgePreparing");
            var hadExistingLink = GetSelectedSessionOpenCodeLink() is not null;
            var linkRecord = await EnsureOpenCodeLinkAsync(selectedSession, forceRefresh: false);
            _openCodeBridgeService.LaunchSession(linkRecord);
            _logService.Info(
                nameof(MainWindow),
                $"OpenCode bridge ready. Session={selectedSession.SessionId}; OpenCode={linkRecord.OpenCodeSessionId}; Existing={hadExistingLink}.");
            SetStatus(
                "#F8E7D6",
                hadExistingLink
                    ? "StatusOpenCodeStarted"
                    : "StatusOpenCodeBridgeCreated");
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "OpenCode bridge resume/create failed.", exception);
            SetStatus("#FFD6D6", "StatusOpenCodeBridgeFailed", exception.Message);
            MessageBox.Show(
                Strings.Format("StatusOpenCodeBridgeFailed", exception.Message),
                Strings["DetailOpenCodeBridge"],
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsOpenCodeBusy = false;
        }
    }

    private async void RefreshSelectedSessionOpenCodeBridgeButton_Click(object sender, RoutedEventArgs e)
    {
        var selectedSession = SelectedSession;

        _logService.Info(
            nameof(MainWindow),
            $"OpenCode bridge click. Session={selectedSession?.SessionId ?? "-"}; Action=Refresh.");

        if (selectedSession is null)
        {
            return;
        }

        var issue = GetOpenCodeBridgeIssue(selectedSession);
        if (!string.IsNullOrWhiteSpace(issue))
        {
            SetStatus("#FFD6D6", "StatusOpenCodeBridgeFailed", issue);
            MessageBox.Show(
                issue,
                Strings["DetailOpenCodeBridge"],
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            IsOpenCodeBusy = true;
            SetStatus("#F8E7D6", "StatusOpenCodeBridgePreparing");
            await EnsureOpenCodeLinkAsync(selectedSession, forceRefresh: true);
            _logService.Info(
                nameof(MainWindow),
                $"OpenCode bridge refreshed. Session={selectedSession.SessionId}.");
            SetStatus("#F8E7D6", "StatusOpenCodeBridgeRefreshed");
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "OpenCode bridge refresh failed.", exception);
            SetStatus("#FFD6D6", "StatusOpenCodeBridgeFailed", exception.Message);
            MessageBox.Show(
                Strings.Format("StatusOpenCodeBridgeFailed", exception.Message),
                Strings["DetailOpenCodeBridge"],
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsOpenCodeBusy = false;
        }
    }

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshSessionsAsync(isAutomaticRefresh: false, forceRefresh: true);
    }

    private async void RefreshTimer_Tick(object? sender, EventArgs e)
    {
        await RefreshSessionsAsync(isAutomaticRefresh: true);
    }

    private void SearchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        ApplyFilter();
    }

    private void LayoutRefreshTimer_Tick(object? sender, EventArgs e)
    {
        _layoutRefreshTimer.Stop();
        RefreshAdaptiveLayoutBindings();
    }

    private async void SetupRefreshTimer_Tick(object? sender, EventArgs e)
    {
        ExpireSetupRefreshBoostIfNeeded();

        if (SelectedAppSection != AppSection.Setup && !IsSetupRefreshBoostActive)
        {
            return;
        }

        await RefreshSetupSectionAsync(preserveDnsStatus: true);
    }

    private bool IsSetupRefreshBoostActive => _setupRefreshBoostUntilUtc > DateTime.UtcNow;

    private void BeginSetupAction(string statusText, Action focusAction)
    {
        _environmentService.InvalidateSnapshotCaches();
        focusAction();
        BeginSetupRefreshBoost();
        SetSetupStatus("#F8E7D6", statusText);
    }

    private void BeginSetupRefreshBoost()
    {
        _setupRefreshBoostUntilUtc = DateTime.UtcNow.Add(SetupRefreshBoostDuration);
        MarkSetupRefreshPending();
        RefreshSetupOverviewBindings();
        UpdateSetupRefreshTimer();
    }

    private void ExpireSetupRefreshBoostIfNeeded()
    {
        if (_setupRefreshBoostUntilUtc == DateTime.MinValue || IsSetupRefreshBoostActive)
        {
            return;
        }

        _setupRefreshBoostUntilUtc = DateTime.MinValue;
        RefreshSetupOverviewBindings();
        UpdateSetupRefreshTimer();
    }

    private void FocusSetupCoreSection()
    {
        SelectedAppSection = AppSection.Setup;
        IsSetupCoreSectionExpanded = true;
        IsSetupCodexSectionExpanded = false;
        IsSetupLocalAiSectionExpanded = false;
    }

    private void FocusSetupCodexSection()
    {
        SelectedAppSection = AppSection.Setup;
        IsSetupCoreSectionExpanded = false;
        IsSetupCodexSectionExpanded = true;
        IsSetupLocalAiSectionExpanded = false;
    }

    private void FocusSetupLocalAiSection()
    {
        SelectedAppSection = AppSection.Setup;
        IsSetupCoreSectionExpanded = false;
        IsSetupCodexSectionExpanded = false;
        IsSetupLocalAiSectionExpanded = true;
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
        RefreshSessionCountBindings();
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        ExportSessionsFeedSafe();
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

    private void HomeSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Home;
    }

    private void HomeExampleButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string prompt })
        {
            NewSessionPrompt = prompt;
        }
    }

    private void HomeStartSafeSessionButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NewSessionPrompt))
        {
            SetHomeLaunchStatus(
                "#FDECEC",
                "#B42318",
                Strings["HomeLaunchPromptRequired"]);
            HomePromptTextBox.Focus();
            return;
        }

        if (!IsHomeEnvironmentReady)
        {
            SetHomeLaunchStatus(
                "#FFF3E0",
                "#7A4208",
                Strings["HomeLaunchSetupRequired"]);
            return;
        }

        EnsureSectionDataInitialized(AppSection.NewSession);
        NewSessionModel = string.Empty;
        NewSessionProfile = string.Empty;
        NewSessionUseOss = false;
        SelectedSandboxMode = "workspace-write";
        SelectedApprovalPolicy = "on-request";

        try
        {
            MarkSessionsRefreshPending();
            var workspace = _environmentService.LaunchInteractiveSession(
                BuildNewSessionLaunchOptions(workingDirectoryOverride: string.Empty));
            SetHomeLaunchStatus(
                "#E7F6EE",
                "#1F6F4A",
                Strings.Format("HomeLaunchStarted", workspace));
        }
        catch (Exception exception)
        {
            SetHomeLaunchStatus(
                "#FDECEC",
                "#B42318",
                Strings.Format("HomeLaunchFailed", exception.Message));
        }
    }

    private void HomeOpenAdvancedLaunchButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.NewSession;
    }

    private void HomeOpenSetupButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Setup;
        FocusSetupCoreSection();
    }

    private void HomeRecentSessionButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: SessionRecord session })
        {
            return;
        }

        SelectedSession = session;
        SelectedAppSection = AppSection.Sessions;
    }

    private void StartBeginnerOnboardingButton_Click(object sender, RoutedEventArgs e)
    {
        _beginnerOnboardingInProgress = true;
        _showBeginnerOnboarding = false;
        OnPropertyChanged(nameof(BeginnerOnboardingVisibility));
        OnPropertyChanged(nameof(HomeWorkspaceVisibility));
        SelectedAppSection = AppSection.Setup;
        FocusSetupCoreSection();

        if (IsHomeEnvironmentReady)
        {
            CompleteBeginnerOnboarding();
        }
    }

    private void SkipBeginnerOnboardingButton_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show(
                Strings["OnboardingSkipConfirmText"],
                Strings["OnboardingSkipConfirmTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        CompleteBeginnerOnboarding();
    }

    private void ReplayBeginnerOnboardingButton_Click(object sender, RoutedEventArgs e)
    {
        _beginnerOnboardingInProgress = false;
        _hasCompletedBeginnerOnboarding = false;
        _showBeginnerOnboarding = true;
        _settingsService.SaveHasCompletedBeginnerOnboarding(false);
        OnPropertyChanged(nameof(BeginnerOnboardingVisibility));
        OnPropertyChanged(nameof(HomeWorkspaceVisibility));
        SelectedAppSection = AppSection.Home;
    }

    private void CompleteBeginnerOnboarding()
    {
        _beginnerOnboardingInProgress = false;
        _hasCompletedBeginnerOnboarding = true;
        _showBeginnerOnboarding = false;
        _settingsService.SaveHasCompletedBeginnerOnboarding(true);
        OnPropertyChanged(nameof(BeginnerOnboardingVisibility));
        OnPropertyChanged(nameof(HomeWorkspaceVisibility));
    }

    private void ExtensionsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Extensions;
    }

    private void SuggestedExtensionsTabButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_showInstalledExtensionsTab && !_showCustomExtensionsTab)
        {
            return;
        }

        _showInstalledExtensionsTab = false;
        _showCustomExtensionsTab = false;
        RefreshExtensionTabBindings();
        SelectVisibleExtension(SelectedExtension?.Id);
    }

    private void InstalledExtensionsTabButton_Click(object sender, RoutedEventArgs e)
    {
        if (_showInstalledExtensionsTab)
        {
            return;
        }

        _showInstalledExtensionsTab = true;
        _showCustomExtensionsTab = false;
        RefreshExtensionTabBindings();
        SelectVisibleExtension(SelectedExtension?.Id);
    }

    private void CustomExtensionsTabButton_Click(object sender, RoutedEventArgs e)
    {
        if (_showCustomExtensionsTab)
        {
            return;
        }

        _showInstalledExtensionsTab = false;
        _showCustomExtensionsTab = true;
        RefreshExtensionTabBindings();
        SelectVisibleExtension(SelectedExtension?.Id);
    }

    private void AllExtensionsTargetTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedExtensionTarget = "All";
    }

    private void CodexExtensionsTargetTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedExtensionTarget = "Codex";
    }

    private void OpenCodeExtensionsTargetTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedExtensionTarget = "OpenCode";
    }

    private void LmStudioExtensionsTargetTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedExtensionTarget = "LmStudio";
    }

    private void NewCustomPluginButton_Click(object sender, RoutedEventArgs e)
    {
        PrepareNewCustomExtension(AiExtensionKind.Plugin);
    }

    private void NewCustomMcpButton_Click(object sender, RoutedEventArgs e)
    {
        PrepareNewCustomExtension(AiExtensionKind.Mcp);
    }

    private void SetupSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Setup;
    }

    private void OpenSetupCoreSectionButton_Click(object sender, RoutedEventArgs e)
    {
        FocusSetupCoreSection();
    }

    private void OpenSetupCodexSectionButton_Click(object sender, RoutedEventArgs e)
    {
        FocusSetupCodexSection();
    }

    private void OpenSetupLocalAiSectionButton_Click(object sender, RoutedEventArgs e)
    {
        FocusSetupLocalAiSection();
    }

    private void ToggleSetupCoreSectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsSetupCoreSectionExpanded)
        {
            IsSetupCoreSectionExpanded = false;
            return;
        }

        FocusSetupCoreSection();
    }

    private void ToggleSetupCodexSectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsSetupCodexSectionExpanded)
        {
            IsSetupCodexSectionExpanded = false;
            return;
        }

        FocusSetupCodexSection();
    }

    private void ToggleSetupLocalAiSectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsSetupLocalAiSectionExpanded)
        {
            IsSetupLocalAiSectionExpanded = false;
            return;
        }

        FocusSetupLocalAiSection();
    }

    private void ToggleSetupDnsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        IsSetupDnsSectionExpanded = !IsSetupDnsSectionExpanded;
    }

    private void SettingsSectionButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedAppSection = AppSection.Settings;
    }

    private void NeuralSettingsTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedSettingsCategoryTab = SettingsCategoryTab.NeuralSettings;
    }

    private void AppSettingsTabButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedSettingsCategoryTab = SettingsCategoryTab.AppSettings;
    }

    private void UseSelectedSessionDirectoryButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanUseSelectedSessionDirectory || SelectedSession is null)
        {
            return;
        }

        NewSessionWorkingDirectory = AiHelperWorkspaceService.ResolveSafeWorkspace(
            SelectedSession.WorkingDirectory,
            SelectedSession.SessionId,
            SelectedSession.Title,
            out _);
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionStatusDirectoryCopied"]);
    }

    private void CopyNewSessionPreviewCommandButton_Click(object sender, RoutedEventArgs e)
    {
        Clipboard.SetText(NewSessionPreviewCommandText);
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionStatusCommandCopied"]);
    }

    private void LaunchNewSessionButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NewSessionPrompt))
        {
            SetNewSessionStatus("#FFD6D6", Strings["HomeLaunchPromptRequired"]);
            return;
        }

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

        if (!ConfirmDangerousNewSessionLaunch())
        {
            return;
        }

        try
        {
            MarkSessionsRefreshPending();
            _environmentService.LaunchInteractiveSession(BuildNewSessionLaunchOptions());
            SetNewSessionStatus("#F8E7D6", Strings["NewSessionStatusStarted"]);
        }
        catch (Exception exception)
        {
            SetNewSessionStatus("#FFD6D6", Strings.Format("NewSessionStatusLaunchFailed", exception.Message));
        }
    }

    private void LaunchNewSessionWithImageButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NewSessionPrompt))
        {
            SetNewSessionStatus("#FFD6D6", Strings["HomeLaunchPromptRequired"]);
            return;
        }

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

        var imagePath = GetLaunchImagePath();

        if (string.IsNullOrWhiteSpace(imagePath))
        {
            SetNewSessionStatus("#F8E7D6", Strings["StatusImageSelectionCanceled"]);
            return;
        }

        if (!ConfirmDangerousNewSessionLaunch())
        {
            return;
        }

        try
        {
            MarkSessionsRefreshPending();
            _environmentService.LaunchInteractiveSession(BuildNewSessionLaunchOptions([imagePath]));
            SetNewSessionStatus("#F8E7D6", Strings.Format("NewSessionStatusStartedWithImage", Path.GetFileName(imagePath)));
        }
        catch (Exception exception)
        {
            SetNewSessionStatus("#FFD6D6", Strings.Format("NewSessionStatusLaunchFailed", exception.Message));
        }
    }

    private async void RefreshSetupStatusButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshSetupSectionAsync(preserveDnsStatus: false, forceRefresh: true);
    }

    private void InstallBaseComponentsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchPrerequisitesInstallTerminal();
            BeginSetupAction(Strings["SetupStatusBaseComponentsInstallStarted"], FocusSetupCoreSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallNodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchNodeInstallTerminal();
            BeginSetupAction(Strings["SetupStatusNodeInstallStarted"], FocusSetupCoreSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallGitButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchGitInstallTerminal();
            BeginSetupAction(Strings["SetupStatusGitInstallStarted"], FocusSetupCoreSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void RepairWingetButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchWingetRepairTerminal();
            BeginSetupAction(Strings["SetupStatusWingetRepairStarted"], FocusSetupCoreSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private async void CheckForUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshUpdateStatusAsync(forceRefresh: true);
    }

    private async void DownloadUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        var snapshot = _lastAppUpdateSnapshot;

        if (snapshot is null)
        {
            await RefreshUpdateStatusAsync(forceRefresh: true);
            snapshot = _lastAppUpdateSnapshot;
        }

        if (snapshot is null || !snapshot.IsUpdateAvailable || !snapshot.HasInstallerAsset)
        {
            return;
        }

        var latestVersion = snapshot.LatestVersionDisplay.TrimStart('v', 'V');
        var installerPath = _updateService.GetDefaultInstallerPath(latestVersion);

        try
        {
            IsUpdateBusy = true;
            SetUpdateStatus("#F8E7D6", "UpdateStatusDownloading");
            await _updateService.DownloadInstallerAsync(
                snapshot.InstallerDownloadUrl,
                snapshot.InstallerChecksumUrl,
                installerPath);
            SetUpdateStatus("#F8E7D6", "UpdateStatusDownloaded", installerPath);

            var executablePath = Environment.ProcessPath ??
                                 Process.GetCurrentProcess().MainModule?.FileName ??
                                 string.Empty;
            _updateService.StartSilentInstallerAndRestart(installerPath, executablePath);
            SetUpdateStatus("#F8E7D6", "UpdateStatusInstallerStarted");

            await Task.Delay(250);
            Application.Current.Shutdown(0);
        }
        catch (Exception exception)
        {
            SetUpdateStatus("#FFD6D6", "UpdateStatusLaunchFailed", exception.Message);
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

    private void ApplyPhotoPasteFixSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        if (SettingsPhotoPasteFixEnabled)
        {
            var confirmation = MessageBox.Show(
                Strings["SettingsPhotoPasteFixWarningMessage"],
                Strings["SettingsPhotoPasteFixWarningTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                SettingsPhotoPasteFixEnabled = false;
                return;
            }
        }

        try
        {
            _photoPasteFixService.UpdateConfiguration(SettingsPhotoPasteFixEnabled);
            _settingsService.SavePhotoPasteFixEnabled(SettingsPhotoPasteFixEnabled);
            SetSettingsStatus(
                "#F8E7D6",
                SettingsPhotoPasteFixEnabled
                    ? "SettingsStatusPhotoPasteFixEnabled"
                    : "SettingsStatusPhotoPasteFixDisabled");
        }
        catch (Exception exception)
        {
            SettingsPhotoPasteFixEnabled = _photoPasteFixService.IsEnabled;
            SetSettingsStatus("#FFD6D6", "SettingsStatusPhotoPasteFixFailed", exception.Message);
            _logService.Error(nameof(MainWindow), "Failed to apply the photo paste fix setting.", exception);
        }
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
            BeginSetupAction(Strings["SetupStatusInstallerStarted"], FocusSetupCodexSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallCodexDesktopAppButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchCodexDesktopInstallTerminal();
            BeginSetupAction(Strings["SetupStatusCodexDesktopInstallStarted"], FocusSetupCodexSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenCodexDesktopStorePageButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.OpenCodexDesktopStorePage();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusCodexDesktopStoreOpened"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallOpenCodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenCodeInstallTerminal();
            BeginSetupAction(Strings["SetupStatusOpenCodeInstallStarted"], FocusSetupCodexSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void LaunchOpenCodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenCodeTerminal();
            BeginSetupAction(Strings["SetupStatusOpenCodeStarted"], FocusSetupCodexSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void UninstallOpenCodeButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmLocalAiRemoval(
                Strings["SetupRemoveRuntimeWarningTitle"],
                Strings.Format("SetupRemoveRuntimeWarningMessage", "OpenCode")))
        {
            return;
        }

        try
        {
            _environmentService.LaunchOpenCodeUninstallTerminal();
            BeginSetupAction(Strings["SetupStatusOpenCodeUninstallStarted"], FocusSetupCodexSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenOpenCodeDocsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.OpenOpenCodeDocsPage();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenCodeDocsOpened"]);
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
            BeginSetupAction(Strings["SetupStatusOllamaInstallStarted"], FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void LaunchOllamaAppButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOllamaApp();
            BeginSetupAction(Strings["SetupStatusOllamaAppStarted"], FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void StartOllamaServerButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOllamaServeTerminal();
            BeginSetupAction(Strings["SetupStatusOllamaServerStarted"], FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void StopOllamaProcessesButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ConfirmLocalAiRemoval(
                Strings["SetupStopOllamaWarningTitle"],
                Strings["SetupStopOllamaWarningMessage"]))
        {
            return;
        }

        try
        {
            _environmentService.LaunchOllamaStopTerminal();
            BeginSetupAction(Strings["SetupStatusOllamaStopStarted"], FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallStarterOllamaModelButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_environmentService.IsOllamaInstalled())
        {
            SetSetupStatus("#FFD6D6", Strings["SetupStatusOllamaMissing"]);
            return;
        }

        var option = LocalAiModelOptions.FirstOrDefault(model => model.IsRecommended) ??
                     LocalAiModelOptions.FirstOrDefault();

        if (option is null)
        {
            SetSetupStatus("#FFD6D6", Strings["SetupHardwarePending"]);
            return;
        }

        if (!option.CanInstall)
        {
            SetSetupStatus("#FFD6D6", Strings["SetupStatusModelDoesNotFit"]);
            return;
        }

        try
        {
            _environmentService.LaunchOllamaModelInstallTerminal(option.ModelTag);
            BeginSetupAction(
                Strings.Format("SetupStatusModelInstallStarted", option.Name),
                FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings["SetupStatusLmStudioInstallStarted"], FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings["SetupStatusOllamaUninstallStarted"], FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings["SetupStatusLmStudioUninstallStarted"], FocusSetupLocalAiSection);
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

        if (!option.CanInstall)
        {
            SetSetupStatus("#FFD6D6", Strings["SetupStatusModelDoesNotFit"]);
            return;
        }

        try
        {
            _environmentService.LaunchOllamaModelInstallTerminal(option.ModelTag);
            BeginSetupAction(Strings.Format("SetupStatusModelInstallStarted", option.Name), FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings.Format("SetupStatusModelRemoveStarted", option.Name), FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings.Format("SetupStatusCreativeToolInstallStarted", option.Name), FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings.Format("SetupStatusCreativeToolRemoveStarted", option.Name), FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings.Format("SetupStatusAiAgentInstallStarted", option.Name), FocusSetupLocalAiSection);
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
            BeginSetupAction(Strings.Format("SetupStatusAiAgentRemoveStarted", option.Name), FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenClawStatusButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenClawStatusTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenClawStatusStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void InstallOpenClawNodeButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenClawNodeInstallTerminal();
            BeginSetupAction(Strings["SetupStatusOpenClawNodeInstallStarted"], FocusSetupLocalAiSection);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenClawNodeStatusButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenClawNodeStatusTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenClawNodeStatusStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenClawBrowserStatusButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _environmentService.LaunchOpenClawBrowserStatusTerminal();
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenClawBrowserStatusStarted"]);
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void OpenClawConfigButton_Click(object sender, RoutedEventArgs e)
    {
        var configPath = _lastEnvironmentSnapshot?.OpenClawConfigPath ?? _environmentService.OpenClawConfigFilePath;
        var configDirectory = Path.GetDirectoryName(configPath);

        if (File.Exists(configPath))
        {
            OpenExplorerSelect(configPath);
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenClawConfigOpened"]);
            return;
        }

        if (!string.IsNullOrWhiteSpace(configDirectory) && Directory.Exists(configDirectory))
        {
            CodexEnvironmentService.OpenFolder(configDirectory);
            SetSetupStatus("#F8E7D6", Strings["SetupStatusOpenClawConfigOpened"]);
            return;
        }

        SetSetupStatus("#FFD6D6", Strings.Format("StatusFolderNotFound", configPath));
    }

    private async void ApplyOpenClawQuickModeButton_Click(object sender, RoutedEventArgs e)
    {
        await ApplyOpenClawModeAsync(
            () => _environmentService.ApplyOpenClawQuickStartMode(),
            "SetupStatusOpenClawQuickModeApplied");
    }

    private async void ApplyOpenClawAdvancedModeButton_Click(object sender, RoutedEventArgs e)
    {
        await ApplyOpenClawModeAsync(
            () => _environmentService.ApplyOpenClawAdvancedMode(),
            "SetupStatusOpenClawAdvancedModeApplied");
    }

    private async void PrepareOpenClawAlmostFullModeButton_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show(
                Strings["SetupOpenClawAlmostFullWarningMessage"],
                Strings["SetupOpenClawAlmostFullWarningTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning) != MessageBoxResult.Yes)
        {
            return;
        }

        await ApplyOpenClawModeAsync(
            () => _environmentService.PrepareOpenClawAlmostFullMode(),
            "SetupStatusOpenClawAlmostFullModeApplied");
    }

    private async Task ApplyOpenClawModeAsync(
        Func<OpenClawConfigApplyResult> applyMode,
        string successStatusKey)
    {
        try
        {
            var result = await Task.Run(applyMode);
            await RefreshSetupSectionAsync(preserveDnsStatus: true, forceRefresh: true);

            if (string.IsNullOrWhiteSpace(result.BackupPath))
            {
                SetSetupStatus(
                    "#F8E7D6",
                    Strings.Format(successStatusKey, result.PrimaryModel, result.ToolsProfile));
            }
            else
            {
                SetSetupStatus(
                    "#F8E7D6",
                    Strings.Format(
                        "SetupStatusOpenClawModeAppliedWithBackup",
                        result.PrimaryModel,
                        result.ToolsProfile,
                        result.BackupPath));
            }
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
        }
    }

    private void ApplyBeginnerCloudPresetButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyBeginnerCloudPreset();
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionPresetCloudApplied"]);
    }

    private void ApplyBeginnerLocalPresetButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyBeginnerLocalPreset();
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionPresetLocalApplied"]);
    }

    private void ApplyBeginnerCloudPreset()
    {
        NewSessionUseOss = false;
        NewSessionUseSearch = false;
        SelectedLocalProvider = string.Empty;
        SelectedSandboxMode = SettingsDangerousFullAccess ? "danger-full-access" : "workspace-write";
        SelectedApprovalPolicy = SettingsDangerousFullAccess ? "never" : "on-request";

        if (NewSessionModel.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase))
        {
            NewSessionModel = _configuredCodexModel;
        }

        if (!ProfileSuggestions.Contains(NewSessionProfile))
        {
            NewSessionProfile = string.Empty;
        }
    }

    private void ApplyBeginnerLocalPreset()
    {
        NewSessionUseOss = true;
        NewSessionUseSearch = false;
        SelectedLocalProvider = "ollama";
        SelectedSandboxMode = "workspace-write";
        SelectedApprovalPolicy = "on-request";

        if (string.IsNullOrWhiteSpace(NewSessionModel) ||
            !NewSessionModel.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase))
        {
            NewSessionModel = string.Empty;
        }

        NewSessionProfile = string.Empty;
    }

    private void ApplyProjectReviewExampleButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyBeginnerCloudPreset();
        NewSessionPrompt = Strings["NewSessionExampleProjectPrompt"];
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionExampleProjectApplied"]);
    }

    private void ApplyBugFixExampleButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyBeginnerCloudPreset();
        NewSessionPrompt = Strings["NewSessionExampleBugfixPrompt"];
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionExampleBugfixApplied"]);
    }

    private void ApplyLocalExampleButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyBeginnerLocalPreset();
        NewSessionPrompt = Strings["NewSessionExampleLocalPrompt"];
        SetNewSessionStatus("#F8E7D6", Strings["NewSessionExampleLocalApplied"]);
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
            BeginSetupAction(Strings["SetupStatusLoginStarted"], FocusSetupCodexSection);
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
        IEnumerable<SessionRecord> query = SelectedSessionListTab == SessionListTab.Favorites
            ? _allSessions.Where(session => session.IsFavorite)
            : _allSessions.Where(session => !session.IsFavorite);

        if (!ShowHiddenSessions)
        {
            query = query.Where(session => !session.IsHidden);
        }

        if (!string.IsNullOrWhiteSpace(filter))
        {
            query = query.Where(session => session.SearchBlob.Contains(filter, StringComparison.OrdinalIgnoreCase));
        }

        var filtered = query
            .OrderByDescending(session => session.UpdatedAtUtc)
            .ToList();

        ReplaceVisibleSessions(filtered);

        if (Sessions.Count == 0)
        {
            SelectedSession = null;
            OnPropertyChanged(nameof(HasVisibleSessions));
            return;
        }

        SelectedSession = Sessions.FirstOrDefault(session => session.SessionId == currentId) ?? Sessions[0];
        OnPropertyChanged(nameof(HasVisibleSessions));
    }

    private void ReplaceVisibleSessions(IReadOnlyList<SessionRecord> filteredSessions)
    {
        if (Sessions.Count == filteredSessions.Count)
        {
            var sequenceMatches = true;

            for (var index = 0; index < filteredSessions.Count; index++)
            {
                if (!ReferenceEquals(Sessions[index], filteredSessions[index]))
                {
                    sequenceMatches = false;
                    break;
                }
            }

            if (sequenceMatches)
            {
                return;
            }
        }

        Sessions.Clear();

        foreach (var session in filteredSessions)
        {
            Sessions.Add(session);
        }
    }

    private void ApplySessions(IReadOnlyList<SessionRecord> refreshedSessions)
    {
        _allSessions = refreshedSessions.ToList();

        foreach (var session in _allSessions)
        {
            session.IsFavorite = _favoriteSessionIds.Contains(session.SessionId);
            session.IsHidden = _hiddenSessionIds.Contains(session.SessionId);
            session.Note = _sessionNotes.TryGetValue(session.SessionId, out var note) ? note : string.Empty;
            UpdateSessionSearchBlob(session);
        }

        TotalSessions = _allSessions.Count;
        UpdatedTodaySessions = _allSessions.Count(
            session => session.UpdatedAtUtc.ToLocalTime().Date == DateTime.Today);
        TotalMessages = _allSessions.Sum(session => session.TotalMessageCount);
        TotalToolCalls = _allSessions.Sum(session => session.ToolCallCount);
        RefreshSessionCountBindings();
        RefreshHomeRecentSessions();
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        ExportSessionsFeedSafe();
        ApplyFilter();
        RefreshOpenCodeBindings();
    }

    private void RefreshHomeRecentSessions()
    {
        HomeRecentSessions.Clear();

        foreach (var session in _allSessions
                     .Where(session => !session.IsHidden)
                     .OrderByDescending(session => session.UpdatedAtUtc)
                     .Take(3))
        {
            HomeRecentSessions.Add(session);
        }

        OnPropertyChanged(nameof(HasHomeRecentSessions));
        OnPropertyChanged(nameof(HomeRecentSessionsVisibility));
        OnPropertyChanged(nameof(HomeRecentSessionsEmptyVisibility));
    }

    private void ExportSessionsFeedSafe()
    {
        try
        {
            _sessionFeedExportService.SaveSessions(_allSessions);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to export AIHelper session feed.", exception);
        }
    }

    private void RefreshSessionCountBindings()
    {
        OnPropertyChanged(nameof(FavoriteSessions));
        OnPropertyChanged(nameof(RegularSessions));
        OnPropertyChanged(nameof(HiddenSessions));
        OnPropertyChanged(nameof(SessionsTabText));
        OnPropertyChanged(nameof(FavoritesTabText));
        OnPropertyChanged(nameof(HiddenSessionsToggleText));
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
        OnPropertyChanged(nameof(HiddenSessionsToggleText));
        OnPropertyChanged(nameof(EmptySessionsText));
        OnPropertyChanged(nameof(SelectedSessionTitleText));
        OnPropertyChanged(nameof(SelectedSessionPreviewText));
        OnPropertyChanged(nameof(SelectedSessionTranscriptText));
        OnPropertyChanged(nameof(SelectedSessionFavoriteText));
        OnPropertyChanged(nameof(SelectedSessionHealthTitle));
        OnPropertyChanged(nameof(SelectedSessionHealthText));
        OnPropertyChanged(nameof(ToggleSessionHiddenButtonText));
        OnPropertyChanged(nameof(CanArchiveSelectedSession));
        OnPropertyChanged(nameof(CanToggleSelectedSessionHidden));
        OnPropertyChanged(nameof(CanOpenSelectedSessionDirectory));
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
        NotifyNewSessionAccessSummaryChanged();
        OnPropertyChanged(nameof(NewSessionDataRouteText));
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
        OnPropertyChanged(nameof(CanLaunchOllamaApp));
        OnPropertyChanged(nameof(CanStartOllamaServer));
        OnPropertyChanged(nameof(CanStopOllamaProcesses));
        OnPropertyChanged(nameof(CanInstallStarterOllamaModel));
        OnPropertyChanged(nameof(OllamaQuickGuidanceText));
        OnPropertyChanged(nameof(CanManageCreativeAiTools));
        OnPropertyChanged(nameof(CanManageAiAgents));
        OnPropertyChanged(nameof(CanApplyOpenClawModes));
        OnPropertyChanged(nameof(CanInspectOpenClawStatus));
        OnPropertyChanged(nameof(CanInstallOpenClawNode));
        OnPropertyChanged(nameof(CanInspectOpenClawNode));
        OnPropertyChanged(nameof(CanInspectOpenClawBrowser));
        OnPropertyChanged(nameof(CanOpenOpenClawConfig));
        OnPropertyChanged(nameof(OpenClawDetectedConfigText));
        OnPropertyChanged(nameof(OpenClawRecommendationText));
        OnPropertyChanged(nameof(CanInstallOpenCode));
        OnPropertyChanged(nameof(CanLaunchOpenCode));
        OnPropertyChanged(nameof(CanUninstallOpenCode));
        OnPropertyChanged(nameof(OpenCodeSetupDetailText));
        OnPropertyChanged(nameof(OpenCodeResumeButtonText));
        OnPropertyChanged(nameof(SelectedSessionOpenCodeBridgeText));
        OnPropertyChanged(nameof(SelectedExtensionDetailsText));
        OnPropertyChanged(nameof(CanDeleteSelectedExtension));
        OnPropertyChanged(nameof(CanOpenSelectedExtensionLocation));
        OnPropertyChanged(nameof(CanInstallSelectedExtension));
        OnPropertyChanged(nameof(CanEnableSelectedExtension));
        OnPropertyChanged(nameof(CanDisableSelectedExtension));
        OnPropertyChanged(nameof(CanRemoveSelectedExtension));
        OnPropertyChanged(nameof(CanSaveSelectedExtension));

        RefreshLaunchOptionCollections();
        RefreshExtensionKindOptions();
        RefreshExtensionTargetOptions();
        if (_isExtensionsSectionInitialized)
        {
            LoadExtensionsSafe();
        }

        RefreshLocalAiModelOptions();
        RefreshCreativeAiToolOptions(_lastEnvironmentSnapshot);
        RefreshAiAgentToolOptions(_lastEnvironmentSnapshot);
        RefreshOpenClawSetupModes(_lastEnvironmentSnapshot);
        RefreshOpenClawCapabilityChecks(_lastEnvironmentSnapshot);
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

        if (SelectedSession is not null)
        {
            SelectedSession.TranscriptText = string.Empty;
            QueueSelectedSessionTranscriptLoad(SelectedSession);
        }

        if (IsLoaded)
        {
            _ = RefreshSessionsAsync(isAutomaticRefresh: false, forceRefresh: true);
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
        RefreshSetupOverviewBindings();
        OnPropertyChanged(nameof(HomeReadinessText));
        OnPropertyChanged(nameof(HomeStartHelpText));
        OnPropertyChanged(nameof(SetupSubtitleText));
        OnPropertyChanged(nameof(SettingsPhotoPasteFixStateText));
        RefreshOpenCodeBindings();
    }

    private void RefreshSectionChromeText()
    {
        NewSessionStatusText = Strings["NewSessionStatusReady"];
        NewSessionStatusForeground = "#F8E7D6";
        SetupStatusText = Strings["SetupStatusReady"];
        SetupStatusForeground = "#1F6F4A";
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
        ExtensionStatusText = Strings["ExtensionsStatusReady"];
        ExtensionStatusForeground = "#F8E7D6";
        RefreshSetupOverviewBindings();
        RefreshOpenCodeBindings();
    }

    private void RefreshExtensionKindOptions()
    {
        ReplaceLaunchOptions(
            ExtensionKindOptions,
            [
                new LaunchOption
                {
                    Value = "Plugin",
                    DisplayName = Strings["ExtensionsKindPlugin"],
                    Description = Strings["ExtensionsKindPluginDescription"]
                },
                new LaunchOption
                {
                    Value = "Skill",
                    DisplayName = Strings["ExtensionsKindSkill"],
                    Description = Strings["ExtensionsKindSkillDescription"]
                },
                new LaunchOption
                {
                    Value = "MCP",
                    DisplayName = Strings["ExtensionsKindMcp"],
                    Description = Strings["ExtensionsKindMcpDescription"]
                }
            ]);
    }

    private void RefreshExtensionTargetOptions()
    {
        ReplaceLaunchOptions(
            ExtensionTargetOptions,
            [
                new LaunchOption
                {
                    Value = "Codex",
                    DisplayName = Strings["ExtensionsTargetCodexTab"],
                    Description = Strings["ExtensionsTargetCodexDescription"]
                },
                new LaunchOption
                {
                    Value = "OpenCode",
                    DisplayName = Strings["ExtensionsTargetOpenCodeTab"],
                    Description = Strings["ExtensionsTargetOpenCodeDescription"]
                },
                new LaunchOption
                {
                    Value = "LmStudio",
                    DisplayName = Strings["ExtensionsTargetLmStudioTab"],
                    Description = Strings["ExtensionsTargetLmStudioDescription"]
                }
            ]);
    }

    private void LoadExtensionsSafe()
    {
        try
        {
            var selectedId = SelectedExtension?.Id;
            var items = _extensionCatalogService.LoadExtensions(Strings);
            AiExtensions.Clear();

            foreach (var item in items)
            {
                LocalizeExtensionItem(item);
                AiExtensions.Add(item);
            }

            RemoveObsoleteDetectedToolEntries();
            var detected = MergeFastDetectedExtensions();
            detected |= RemoveMissingDetectedExtensions();
            RefreshExtensionViews(selectedId);
            if (detected)
            {
                SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusDetectedLoaded"]);
            }

            _ = RefreshDetectedExtensionsAsync();
            _ = RefreshManagedExtensionsAsync();

            if (!detected)
            {
                SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusLoaded"]);
            }
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to load extension catalog.", exception);
            SetExtensionStatus("#FFD6D6", Strings.Format("ExtensionsStatusLoadFailed", exception.Message));
        }
    }

    private void PopulateExtensionEditor(AiExtensionItem? item)
    {
        if (item is null)
        {
            ExtensionName = string.Empty;
            SelectedExtensionKind = "Plugin";
            ExtensionTargetApp = NormalizeExtensionTargetAppValue(SelectedExtensionTarget == "All" ? "Codex" : SelectedExtensionTarget);
            ExtensionCommandOrUri = string.Empty;
            ExtensionDescription = string.Empty;
            ExtensionIsEnabled = true;
            return;
        }

        ExtensionName = item.Name;
        SelectedExtensionKind = FormatExtensionKindValue(item.Kind);
        ExtensionTargetApp = NormalizeExtensionTargetAppValue(item.TargetApp);
        ExtensionCommandOrUri = item.CommandOrUri;
        ExtensionDescription = item.Description;
        ExtensionIsEnabled = item.IsEnabled;
    }

    private async void InstallSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        if (!selected.CanProvision)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusNoTrustedInstaller"]);
            return;
        }

        if (!string.Equals(selected.ManagementKind, "endpoint", StringComparison.OrdinalIgnoreCase))
        {
            var confirmation = MessageBox.Show(
                Strings.Format(
                    "ExtensionsInstallConfirmationMessage",
                    selected.Name,
                    string.IsNullOrWhiteSpace(selected.PackageVersion) ? "—" : selected.PackageVersion,
                    string.IsNullOrWhiteSpace(selected.RequestedAccess)
                        ? Strings["ExtensionsAccessUnknown"]
                        : selected.RequestedAccess,
                    selected.CommandOrUri),
                Strings["ExtensionsInstallConfirmationTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }
        }

        selected.IsBusy = true;
        RefreshExtensionGridBindings(selected.Id);
        SetExtensionStatus("#F8E7D6", Strings.Format("ExtensionsStatusInstalling", selected.Name));
        var result = await _extensionManagementService.InstallAsync(selected);
        LocalizeExtensionItem(selected);
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(selected.Id);

        SetExtensionStatus(
            result.Success ? "#B7F7D1" : "#FFD6D6",
            result.Success
                ? Strings.Format("ExtensionsStatusInstalledVerified", selected.Name)
                : Strings.Format("ExtensionsStatusOperationFailed", selected.Name, result.Detail));
    }

    private void EnableSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        selected.IsInstalled = true;
        selected.IsEnabled = true;
        ExtensionIsEnabled = true;
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(selected.Id);
        SetExtensionStatus("#F8E7D6", Strings.Format("ExtensionsStatusEnabled", selected.Name));
    }

    private void DisableSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        selected.IsEnabled = false;
        ExtensionIsEnabled = false;
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(selected.Id);
        SetExtensionStatus("#F8E7D6", Strings.Format("ExtensionsStatusDisabled", selected.Name));
    }

    private async void RemoveSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        if (selected.IsCustom)
        {
            DeleteSelectedExtensionButton_Click(sender, e);
            return;
        }

        var result = MessageBox.Show(
            Strings.Format("ExtensionsRemovePresetWarningMessage", selected.Name),
            Strings["ExtensionsRemovePresetWarningTitle"],
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        if (!selected.CanUninstall)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusNoTrustedRemoval"]);
            return;
        }

        selected.IsBusy = true;
        RefreshExtensionGridBindings(selected.Id);
        SetExtensionStatus("#F8E7D6", Strings.Format("ExtensionsStatusRemoving", selected.Name));
        var operation = await _extensionManagementService.RemoveAsync(selected);
        LocalizeExtensionItem(selected);
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(selected.Id);

        SetExtensionStatus(
            operation.Success ? "#B7F7D1" : "#FFD6D6",
            operation.Success
                ? Strings.Format("ExtensionsStatusRemovedVerified", selected.Name)
                : Strings.Format("ExtensionsStatusOperationFailed", selected.Name, operation.Detail));
    }

    private void OpenSelectedExtensionLocationButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null ||
            !TryGetExtensionFileSystemTarget(selected, forDelete: false, out var targetPath, out var isDirectory))
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsLocationNotFound"]);
            return;
        }

        if (isDirectory)
        {
            CodexEnvironmentService.OpenFolder(targetPath);
        }
        else
        {
            OpenExplorerSelect(targetPath);
        }
    }

    private void ExtensionActiveToggle_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.CheckBox { DataContext: AiExtensionItem item } checkBox)
        {
            return;
        }

        SelectedExtension = item;

        if (!item.IsCustom)
        {
            checkBox.IsChecked = item.IsActive;
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusManagedToggleBlocked"]);
            return;
        }

        var isEnabled = checkBox.IsChecked == true;

        item.IsEnabled = isEnabled;
        if (isEnabled)
        {
            item.IsInstalled = true;
        }

        ExtensionIsEnabled = isEnabled;
        LocalizeExtensionItem(item);
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(item.Id);
        OnPropertyChanged(nameof(SelectedExtensionDetailsText));

        SetExtensionStatus(
            "#F8E7D6",
            Strings.Format(
                isEnabled ? "ExtensionsStatusEnabled" : "ExtensionsStatusDisabled",
                item.Name));
    }

    private void AddExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanSaveSelectedExtension)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusNameRequired"]);
            return;
        }

        var item = BuildExtensionFromEditor();
        item.IsPreset = false;
        item.IsInstalled = true;
        LocalizeExtensionItem(item);
        AiExtensions.Add(item);
        SaveExtensionsSafe();
        _showInstalledExtensionsTab = false;
        _showCustomExtensionsTab = true;
        RefreshExtensionGridBindings(item.Id);
        SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusAdded"]);
    }

    private void SaveSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanSaveSelectedExtension)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusNameRequired"]);
            return;
        }

        if (SelectedExtension?.IsCustom != true)
        {
            var item = BuildExtensionFromEditor();
            item.IsPreset = false;
            item.IsInstalled = true;
            LocalizeExtensionItem(item);
            AiExtensions.Add(item);
            SaveExtensionsSafe();
            _showInstalledExtensionsTab = false;
            _showCustomExtensionsTab = true;
            RefreshExtensionGridBindings(item.Id);
            SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusPresetCopied"]);
            return;
        }

        SelectedExtension.Name = ExtensionName.Trim();
        SelectedExtension.Kind = ParseExtensionKind(SelectedExtensionKind);
        SelectedExtension.TargetApp = NormalizeExtensionTargetAppValue(ExtensionTargetApp);
        SelectedExtension.CommandOrUri = ExtensionCommandOrUri.Trim();
        SelectedExtension.Description = ExtensionDescription.Trim();
        SelectedExtension.IsEnabled = ExtensionIsEnabled;
        SelectedExtension.IsInstalled = SelectedExtension.IsInstalled || ExtensionIsEnabled;
        LocalizeExtensionItem(SelectedExtension);
        SaveExtensionsSafe();
        RefreshExtensionGridBindings(SelectedExtension.Id);
        OnPropertyChanged(nameof(SelectedExtensionDetailsText));
        SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusSaved"]);
    }

    private void DuplicateExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        var copy = selected.Clone();
        copy.Id = Guid.NewGuid().ToString("N");
        copy.Name = Strings.Format("ExtensionsDuplicateName", selected.Name);
        copy.IsPreset = false;
        copy.IsInstalled = true;
        LocalizeExtensionItem(copy);
        AiExtensions.Add(copy);
        SaveExtensionsSafe();
        _showInstalledExtensionsTab = false;
        _showCustomExtensionsTab = true;
        RefreshExtensionGridBindings(copy.Id);
        SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusPresetCopied"]);
    }

    private void DeleteSelectedExtensionButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = SelectedExtension;

        if (selected is null)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsNoSelection"]);
            return;
        }

        if (selected.IsDetected)
        {
            if (!TryGetExtensionFileSystemTarget(selected, forDelete: true, out var targetPath, out var isDirectory))
            {
                SetExtensionStatus("#FFD6D6", Strings["ExtensionsLocationNotFound"]);
                return;
            }

            var detectedResult = MessageBox.Show(
                Strings.Format("ExtensionsDeleteDetectedWarningMessage", selected.Name, targetPath),
                Strings["ExtensionsDeleteWarningTitle"],
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (detectedResult != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                var trashPath = MoveExtensionTargetToTrash(selected, targetPath, isDirectory);
                AiExtensions.Remove(selected);
                SaveExtensionsSafe();
                RefreshExtensionGridBindings();
                SetExtensionStatus("#F8E7D6", Strings.Format("ExtensionsStatusMovedToTrash", trashPath));
            }
            catch (Exception exception)
            {
                _logService.Error(nameof(MainWindow), "Failed to move detected extension to trash.", exception);
                SetExtensionStatus("#FFD6D6", Strings.Format("ExtensionsStatusDeleteFailed", exception.Message));
            }

            return;
        }

        if (!selected.IsCustom)
        {
            SetExtensionStatus("#FFD6D6", Strings["ExtensionsStatusPresetDeleteBlocked"]);
            return;
        }

        var result = MessageBox.Show(
            Strings.Format("ExtensionsDeleteWarningMessage", selected.Name),
            Strings["ExtensionsDeleteWarningTitle"],
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        AiExtensions.Remove(selected);
        SaveExtensionsSafe();
        RefreshExtensionGridBindings();
        SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusDeleted"]);
    }

    private static bool TryGetExtensionFileSystemTarget(
        AiExtensionItem item,
        bool forDelete,
        out string targetPath,
        out bool isDirectory)
    {
        targetPath = string.Empty;
        isDirectory = false;

        var detectionPath = item.DetectionPath?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(detectionPath) && !forDelete)
        {
            if (Directory.Exists(detectionPath))
            {
                targetPath = detectionPath;
                isDirectory = true;
                return true;
            }

            if (File.Exists(detectionPath))
            {
                targetPath = detectionPath;
                isDirectory = false;
                return true;
            }
        }

        if (forDelete &&
            item.IsDetected &&
            !string.IsNullOrWhiteSpace(detectionPath) &&
            !string.Equals(detectionPath, item.CommandOrUri?.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var path = item.CommandOrUri?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        if (Directory.Exists(path))
        {
            targetPath = path;
            isDirectory = true;
            return true;
        }

        if (!File.Exists(path))
        {
            return false;
        }

        if (forDelete &&
            string.Equals(Path.GetFileName(path), "SKILL.md", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path)))
        {
            targetPath = Path.GetDirectoryName(path) ?? path;
            isDirectory = Directory.Exists(targetPath);
            return isDirectory;
        }

        targetPath = path;
        isDirectory = false;
        return true;
    }

    private static string MoveExtensionTargetToTrash(AiExtensionItem item, string targetPath, bool isDirectory)
    {
        var trashRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIHelper",
            "extension-trash");
        Directory.CreateDirectory(trashRoot);

        var targetName = Path.GetFileName(targetPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var safeName = SanitizeFileName(string.IsNullOrWhiteSpace(targetName) ? item.Name : targetName);
        var destinationPath = CreateUniqueExtensionTrashPath(
            Path.Combine(trashRoot, $"{DateTime.Now:yyyyMMdd-HHmmss}-{safeName}"),
            isDirectory);

        if (isDirectory)
        {
            Directory.Move(targetPath, destinationPath);
        }
        else
        {
            var destinationFilePath = destinationPath;
            var destinationDirectory = Path.GetDirectoryName(destinationFilePath);

            if (!string.IsNullOrWhiteSpace(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            File.Move(targetPath, destinationFilePath);
        }

        return destinationPath;
    }

    private static string CreateUniqueExtensionTrashPath(string path, bool isDirectory)
    {
        if (isDirectory)
        {
            if (!Directory.Exists(path) && !File.Exists(path))
            {
                return path;
            }

            for (var index = 1; index < 10_000; index++)
            {
                var candidate = $"{path}-{index}";

                if (!Directory.Exists(candidate) && !File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return $"{path}-{Guid.NewGuid():N}";
        }

        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        var candidatePath = Path.Combine(directory, $"{name}{extension}");

        if (!File.Exists(candidatePath) && !Directory.Exists(candidatePath))
        {
            return candidatePath;
        }

        for (var index = 1; index < 10_000; index++)
        {
            candidatePath = Path.Combine(directory, $"{name}-{index}{extension}");

            if (!File.Exists(candidatePath) && !Directory.Exists(candidatePath))
            {
                return candidatePath;
            }
        }

        return Path.Combine(directory, $"{name}-{Guid.NewGuid():N}{extension}");
    }

    private static string SanitizeFileName(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var safe = new string(value
            .Select(character => invalidCharacters.Contains(character) ? '_' : character)
            .ToArray())
            .Trim();

        return string.IsNullOrWhiteSpace(safe) ? "extension" : safe;
    }

    private AiExtensionItem BuildExtensionFromEditor()
    {
        return new AiExtensionItem
        {
            Id = Guid.NewGuid().ToString("N"),
            Name = ExtensionName.Trim(),
            Kind = ParseExtensionKind(SelectedExtensionKind),
            TargetApp = NormalizeExtensionTargetAppValue(ExtensionTargetApp),
            CommandOrUri = ExtensionCommandOrUri.Trim(),
            Description = ExtensionDescription.Trim(),
            IsInstalled = true,
            IsEnabled = ExtensionIsEnabled
        };
    }

    private void PrepareNewCustomExtension(AiExtensionKind kind)
    {
        SelectedExtension = null;
        SelectedExtensionKind = FormatExtensionKindValue(kind);
        ExtensionTargetApp = NormalizeExtensionTargetAppValue(SelectedExtensionTarget == "All"
            ? "Codex"
            : SelectedExtensionTarget);
        ExtensionName = string.Empty;
        ExtensionCommandOrUri = kind == AiExtensionKind.Mcp
            ? "mcp:"
            : "plugin:";
        ExtensionDescription = string.Empty;
        ExtensionIsEnabled = true;
        _showInstalledExtensionsTab = false;
        _showCustomExtensionsTab = true;
        RefreshExtensionTabBindings();
        SetExtensionStatus(
            "#F8E7D6",
            kind == AiExtensionKind.Mcp
                ? Strings["ExtensionsStatusNewMcpDraft"]
                : Strings["ExtensionsStatusNewPluginDraft"]);
    }

    private void SaveExtensionsSafe()
    {
        try
        {
            _extensionCatalogService.SaveExtensions(AiExtensions);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to save custom extensions.", exception);
            SetExtensionStatus("#FFD6D6", Strings.Format("ExtensionsStatusSaveFailed", exception.Message));
        }
    }

    private void RefreshExtensionGridBindings(string? selectedId = null)
    {
        selectedId ??= SelectedExtension?.Id;
        var currentItems = AiExtensions.ToList();
        foreach (var item in currentItems)
        {
            LocalizeExtensionItem(item);
        }

        AiExtensions.Clear();

        foreach (var item in currentItems)
        {
            AiExtensions.Add(item);
        }

        RefreshExtensionViews(selectedId);
    }

    private void RefreshExtensionViews(string? selectedId = null)
    {
        SuggestedAiExtensions.Clear();
        InstalledAiExtensions.Clear();
        CustomAiExtensions.Clear();

        foreach (var item in AiExtensions)
        {
            LocalizeExtensionItem(item);
            if (!MatchesSelectedExtensionTarget(item))
            {
                continue;
            }

            if (!MatchesExtensionSearch(item))
            {
                continue;
            }

            if (item.IsCustom)
            {
                CustomAiExtensions.Add(item);
            }

            if (item.IsInstalled || item.IsCustom)
            {
                InstalledAiExtensions.Add(item);
            }
            else if (!item.IsDetected)
            {
                SuggestedAiExtensions.Add(item);
            }
        }

        RefreshExtensionTabBindings();
        SelectVisibleExtension(selectedId);
    }

    private void RefreshExtensionTabBindings()
    {
        OnPropertyChanged(nameof(DisplayedAiExtensions));
        OnPropertyChanged(nameof(SuggestedExtensionsTabText));
        OnPropertyChanged(nameof(InstalledExtensionsTabText));
        OnPropertyChanged(nameof(CustomExtensionsTabText));
        OnPropertyChanged(nameof(SuggestedExtensionsTabButtonBackground));
        OnPropertyChanged(nameof(SuggestedExtensionsTabButtonForeground));
        OnPropertyChanged(nameof(InstalledExtensionsTabButtonBackground));
        OnPropertyChanged(nameof(InstalledExtensionsTabButtonForeground));
        OnPropertyChanged(nameof(CustomExtensionsTabButtonBackground));
        OnPropertyChanged(nameof(CustomExtensionsTabButtonForeground));
        OnPropertyChanged(nameof(AllExtensionsTargetTabText));
        OnPropertyChanged(nameof(CodexExtensionsTargetTabText));
        OnPropertyChanged(nameof(OpenCodeExtensionsTargetTabText));
        OnPropertyChanged(nameof(LmStudioExtensionsTargetTabText));
        OnPropertyChanged(nameof(AllExtensionsTargetTabBackground));
        OnPropertyChanged(nameof(AllExtensionsTargetTabForeground));
        OnPropertyChanged(nameof(CodexExtensionsTargetTabBackground));
        OnPropertyChanged(nameof(CodexExtensionsTargetTabForeground));
        OnPropertyChanged(nameof(OpenCodeExtensionsTargetTabBackground));
        OnPropertyChanged(nameof(OpenCodeExtensionsTargetTabForeground));
        OnPropertyChanged(nameof(LmStudioExtensionsTargetTabBackground));
        OnPropertyChanged(nameof(LmStudioExtensionsTargetTabForeground));
    }

    private void SelectVisibleExtension(string? selectedId = null)
    {
        SelectedExtension = DisplayedAiExtensions.FirstOrDefault(item =>
                                string.Equals(item.Id, selectedId, StringComparison.OrdinalIgnoreCase)) ??
                            DisplayedAiExtensions.FirstOrDefault();

        OnPropertyChanged(nameof(CanDeleteSelectedExtension));
        OnPropertyChanged(nameof(CanOpenSelectedExtensionLocation));
        OnPropertyChanged(nameof(CanInstallSelectedExtension));
        OnPropertyChanged(nameof(CanEnableSelectedExtension));
        OnPropertyChanged(nameof(CanDisableSelectedExtension));
        OnPropertyChanged(nameof(CanRemoveSelectedExtension));
        OnPropertyChanged(nameof(CanSaveSelectedExtension));
        OnPropertyChanged(nameof(SelectedExtensionDetailsText));
        OnPropertyChanged(nameof(SelectedExtensionPrimaryActionText));
    }

    private bool MatchesSelectedExtensionTarget(AiExtensionItem item)
    {
        return MatchesExtensionTarget(item, SelectedExtensionTarget);
    }

    private bool MatchesExtensionSearch(AiExtensionItem item)
    {
        if (string.IsNullOrWhiteSpace(ExtensionSearchText))
        {
            return true;
        }

        var query = ExtensionSearchText.Trim();
        return item.Name.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.Description.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.CommandOrUri.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.KindDisplayLabel.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.TargetAppDisplayLabel.Contains(query, StringComparison.OrdinalIgnoreCase) ||
               item.SourceDisplayLabel.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesExtensionTarget(AiExtensionItem item, string target)
    {
        if (string.Equals(target, "All", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return string.Equals(
            NormalizeExtensionTargetAppValue(item.TargetApp),
            NormalizeExtensionTargetAppValue(target),
            StringComparison.OrdinalIgnoreCase);
    }

    private async Task RefreshDetectedExtensionsAsync()
    {
        if (_isDetectedExtensionsRefreshRunning)
        {
            return;
        }

        _isDetectedExtensionsRefreshRunning = true;

        try
        {
            var snapshot = _lastEnvironmentSnapshot ??
                           await Task.Run(_environmentService.GetEnvironmentSnapshot);

            if (_lastEnvironmentSnapshot is null)
            {
                _lastEnvironmentSnapshot = snapshot;
            }

            var selectedId = SelectedExtension?.Id;
            var changed = MergeDetectedExtensions(snapshot);
            changed |= RemoveMissingDetectedExtensions();

            if (changed)
            {
                RefreshExtensionGridBindings(selectedId);
                SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusDetectedLoaded"]);
            }
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to detect local extensions.", exception);
        }
        finally
        {
            _isDetectedExtensionsRefreshRunning = false;
        }
    }

    private async Task RefreshManagedExtensionsAsync()
    {
        if (_isManagedExtensionsRefreshRunning)
        {
            return;
        }

        _isManagedExtensionsRefreshRunning = true;

        try
        {
            var selectedId = SelectedExtension?.Id;
            await _extensionManagementService.RefreshAsync(AiExtensions);
            RefreshExtensionGridBindings(selectedId);
            SetExtensionStatus("#F8E7D6", Strings["ExtensionsStatusVerified"]);
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Failed to verify managed extensions.", exception);
            SetExtensionStatus(
                "#FFD6D6",
                Strings.Format("ExtensionsStatusVerificationFailed", exception.Message));
        }
        finally
        {
            _isManagedExtensionsRefreshRunning = false;
        }
    }

    private bool MergeDetectedExtensions(CodexEnvironmentSnapshot snapshot)
    {
        var changed = false;

        changed |= MergeDetectedCodexSkills(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "skills"),
            "detected-skill");
        changed |= MergeDetectedCodexPlugins(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "plugins"),
            "detected-plugin");
        changed |= MergeDetectedCodexConfig(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "config.toml"));
        changed |= MergeDetectedOpenCodeConfig();
        changed |= MergeDetectedLmStudioConfig();

        return changed;
    }

    private bool MergeDetectedCodexSkills(
        string rootDirectory,
        string idPrefix,
        string targetApp = "Codex")
    {
        if (!Directory.Exists(rootDirectory))
        {
            return false;
        }

        var changed = false;

        foreach (var skillFilePath in Directory.EnumerateFiles(rootDirectory, "SKILL.md", SearchOption.AllDirectories).Take(240))
        {
            var skillDirectory = Path.GetDirectoryName(skillFilePath);
            if (string.IsNullOrWhiteSpace(skillDirectory))
            {
                continue;
            }

            var manifest = ReadSkillManifest(skillFilePath);
            var fallbackName = Path.GetFileName(skillDirectory);
            var name = string.IsNullOrWhiteSpace(manifest.Name) ? fallbackName : manifest.Name;

            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            changed |= UpsertDetectedExtension(
                $"{idPrefix}-{SanitizeExtensionId(GetRelativePathSafe(rootDirectory, skillDirectory))}",
                name,
                AiExtensionKind.Skill,
                skillFilePath,
                skillFilePath,
                string.IsNullOrWhiteSpace(manifest.Description)
                    ? Strings["ExtensionDetectedCodexSkillDescription"]
                    : manifest.Description,
                targetApp);
        }

        return changed;
    }

    private bool MergeDetectedCodexPlugins(
        string rootDirectory,
        string idPrefix)
    {
        if (!Directory.Exists(rootDirectory))
        {
            return false;
        }

        var changed = false;

        foreach (var pluginJsonPath in Directory.EnumerateFiles(rootDirectory, "plugin.json", SearchOption.AllDirectories).Take(120))
        {
            var manifestDirectory = Path.GetDirectoryName(pluginJsonPath);
            if (string.IsNullOrWhiteSpace(manifestDirectory) ||
                !string.Equals(Path.GetFileName(manifestDirectory), ".codex-plugin", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var pluginDirectory = Directory.GetParent(manifestDirectory)?.FullName;
            if (string.IsNullOrWhiteSpace(pluginDirectory))
            {
                continue;
            }

            var manifest = ReadPluginManifest(pluginJsonPath);
            var fallbackName = Path.GetFileName(pluginDirectory);
            var name = string.IsNullOrWhiteSpace(manifest.Name) ? fallbackName : manifest.Name;
            var description = string.IsNullOrWhiteSpace(manifest.Description)
                ? Strings["ExtensionDetectedCodexPluginDescription"]
                : manifest.Description;

            if (!string.IsNullOrWhiteSpace(name))
            {
                changed |= UpsertDetectedExtension(
                    $"{idPrefix}-{SanitizeExtensionId(GetRelativePathSafe(rootDirectory, pluginDirectory))}",
                    name,
                    AiExtensionKind.Plugin,
                    pluginJsonPath,
                    pluginDirectory,
                    description);
            }

            var pluginSkillsDirectory = Path.Combine(pluginDirectory, "skills");
            if (Directory.Exists(pluginSkillsDirectory))
            {
                changed |= MergeDetectedCodexSkills(pluginSkillsDirectory, $"{idPrefix}-skill-{SanitizeExtensionId(name)}");
            }
        }

        return changed;
    }

    private bool MergeDetectedCodexConfig(string configPath)
    {
        if (!File.Exists(configPath))
        {
            return false;
        }

        var changed = false;

        try
        {
            var text = File.ReadAllText(configPath);
            changed |= MergeDetectedTomlMcpServers(
                text,
                configPath,
                "detected-codex-mcp",
                "Codex",
                Strings["ExtensionDetectedCodexMcpDescription"]);

            foreach (var plugin in ReadCodexPluginSections(text))
            {
                changed |= UpsertDetectedExtension(
                    $"detected-codex-config-plugin-{SanitizeExtensionId(plugin.Name)}",
                    plugin.Name,
                    AiExtensionKind.Plugin,
                    configPath,
                    $"plugin:{plugin.Name}",
                    Strings["ExtensionDetectedCodexConfigPluginDescription"],
                    "Codex",
                    plugin.Enabled,
                    configPath);
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            _logService.Info(nameof(MainWindow), $"Failed to read Codex config: {exception.Message}");
        }

        return changed;
    }

    private bool MergeDetectedOpenCodeConfig()
    {
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var configDirectory = Path.Combine(profile, ".config", "opencode");
        var changed = false;

        changed |= MergeDetectedOpenCodeJson(Path.Combine(configDirectory, "opencode.json"));
        changed |= MergeDetectedOpenCodeJson(Path.Combine(configDirectory, "opencode-swarm.json"));
        changed |= MergeDetectedOpenCodeJson(Path.Combine(configDirectory, "dcp.jsonc"));
        changed |= MergeDetectedCodexSkills(
            Path.Combine(configDirectory, "skill-libraries"),
            "detected-opencode-skill",
            "OpenCode");

        return changed;
    }

    private bool MergeDetectedOpenCodeJson(string configPath)
    {
        if (!File.Exists(configPath))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(ReadJsonWithComments(configPath));
            var root = document.RootElement;
            var changed = false;

            if (root.ValueKind == JsonValueKind.Object &&
                TryGetJsonPropertyIgnoreCase(root, "mcp", out var mcpElement) &&
                mcpElement.ValueKind == JsonValueKind.Object)
            {
                changed |= MergeDetectedJsonMcpServers(
                    mcpElement,
                    configPath,
                    "detected-opencode-mcp",
                    "OpenCode",
                    Strings["ExtensionDetectedOpenCodeMcpDescription"]);
            }

            if (root.ValueKind == JsonValueKind.Object &&
                TryGetJsonPropertyIgnoreCase(root, "plugin", out var pluginElement))
            {
                changed |= MergeDetectedJsonPluginList(
                    pluginElement,
                    configPath,
                    "detected-opencode-plugin",
                    "OpenCode",
                    Strings["ExtensionDetectedOpenCodePluginDescription"]);
            }

            return changed;
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            _logService.Info(nameof(MainWindow), $"Failed to read OpenCode config {configPath}: {exception.Message}");
            return false;
        }
    }

    private bool MergeDetectedLmStudioConfig()
    {
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var lmStudioDirectory = Path.Combine(profile, ".lmstudio");
        var changed = false;

        changed |= MergeDetectedLmStudioMcpJson(Path.Combine(lmStudioDirectory, "mcp.json"));
        changed |= MergeDetectedLmStudioMcpServerDirectories(Path.Combine(lmStudioDirectory, "mcp-servers"));
        changed |= MergeDetectedLmStudioPluginManifests(Path.Combine(lmStudioDirectory, "extensions", "plugins"));

        return changed;
    }

    private bool MergeDetectedLmStudioMcpJson(string configPath)
    {
        if (!File.Exists(configPath))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(configPath));
            var root = document.RootElement;

            if (root.ValueKind != JsonValueKind.Object ||
                !TryGetJsonPropertyIgnoreCase(root, "mcpServers", out var serversElement) ||
                serversElement.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            return MergeDetectedJsonMcpServers(
                serversElement,
                configPath,
                "detected-lmstudio-mcp",
                "LmStudio",
                Strings["ExtensionDetectedLmStudioMcpDescription"]);
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            _logService.Info(nameof(MainWindow), $"Failed to read LM Studio MCP config: {exception.Message}");
            return false;
        }
    }

    private bool MergeDetectedLmStudioMcpServerDirectories(string rootDirectory)
    {
        if (!Directory.Exists(rootDirectory))
        {
            return false;
        }

        var changed = false;

        foreach (var directory in SafeEnumerateDirectories(rootDirectory).Take(80))
        {
            var name = Path.GetFileName(directory);
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var bridgeConfigPath = Path.Combine(directory, "mcp-bridge-config.json");
            var commandOrUri = File.Exists(bridgeConfigPath)
                ? ReadMcpCommandSummaryFromJsonFile(bridgeConfigPath)
                : directory;

            changed |= UpsertDetectedExtension(
                $"detected-lmstudio-mcp-dir-{SanitizeExtensionId(name)}",
                name,
                AiExtensionKind.Mcp,
                directory,
                string.IsNullOrWhiteSpace(commandOrUri) ? directory : commandOrUri,
                Strings["ExtensionDetectedLmStudioMcpDirectoryDescription"],
                "LmStudio",
                true,
                directory);
        }

        return changed;
    }

    private bool MergeDetectedLmStudioPluginManifests(string rootDirectory)
    {
        if (!Directory.Exists(rootDirectory))
        {
            return false;
        }

        var changed = false;

        foreach (var manifestPath in SafeEnumerateFiles(rootDirectory, "manifest.json").Take(160))
        {
            if (manifestPath.Contains($"{Path.DirectorySeparatorChar}node_modules{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var directory = Path.GetDirectoryName(manifestPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                continue;
            }

            var manifest = ReadLmStudioPluginManifest(manifestPath);
            var fallbackName = Path.GetFileName(directory);
            var name = string.IsNullOrWhiteSpace(manifest.Name) ? fallbackName : manifest.Name;
            var kind = string.Equals(manifest.Owner, "mcp", StringComparison.OrdinalIgnoreCase)
                ? AiExtensionKind.Mcp
                : AiExtensionKind.Plugin;
            var commandOrUri = kind == AiExtensionKind.Mcp
                ? ReadMcpCommandSummaryFromJsonFile(Path.Combine(directory, "mcp-bridge-config.json"))
                : directory;

            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            changed |= UpsertDetectedExtension(
                $"detected-lmstudio-plugin-{SanitizeExtensionId(GetRelativePathSafe(rootDirectory, directory))}",
                name,
                kind,
                manifestPath,
                string.IsNullOrWhiteSpace(commandOrUri) ? directory : commandOrUri,
                kind == AiExtensionKind.Mcp
                    ? Strings["ExtensionDetectedLmStudioMcpDescription"]
                    : Strings["ExtensionDetectedLmStudioPluginDescription"],
                "LmStudio",
                true,
                directory);
        }

        return changed;
    }

    private bool UpsertDetectedExtension(
        string id,
        string name,
        AiExtensionKind kind,
        string detail,
        string commandOrUri,
        string description,
        string targetApp = "Codex",
        bool isEnabled = true,
        string? detectionPath = null)
    {
        var existing = AiExtensions.FirstOrDefault(item =>
            string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
        var normalizedDetectionPath = string.IsNullOrWhiteSpace(detectionPath)
            ? commandOrUri
            : detectionPath;

        if (existing is null)
        {
            AiExtensions.Add(
                new AiExtensionItem
                {
                    Id = id,
                    Name = name,
                    Kind = kind,
                    TargetApp = NormalizeExtensionTargetAppValue(targetApp),
                    Description = BuildDetectedDescription(description, detail),
                    CommandOrUri = commandOrUri,
                    DetectionPath = normalizedDetectionPath,
                    IsDetected = true,
                    IsInstalled = true,
                    IsEnabled = isEnabled
                });
            return true;
        }

        existing.Name = name;
        existing.Kind = kind;
        existing.TargetApp = NormalizeExtensionTargetAppValue(targetApp);
        existing.Description = BuildDetectedDescription(description, detail);
        existing.CommandOrUri = commandOrUri;
        existing.DetectionPath = normalizedDetectionPath;
        existing.IsPreset = false;
        existing.IsDetected = true;
        existing.IsInstalled = true;
        existing.IsEnabled = isEnabled;

        return true;
    }

    private static string BuildDetectedDescription(string description, string detail)
    {
        return string.IsNullOrWhiteSpace(detail)
            ? description
            : $"{description}{Environment.NewLine}{Environment.NewLine}{detail}";
    }

    private bool MergeDetectedJsonMcpServers(
        JsonElement serversElement,
        string configPath,
        string idPrefix,
        string targetApp,
        string description)
    {
        var changed = false;

        foreach (var server in serversElement.EnumerateObject())
        {
            if (server.Value.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var commandSummary = BuildJsonMcpCommandSummary(server.Value);
            var enabled = !TryGetJsonPropertyIgnoreCase(server.Value, "enabled", out var enabledElement) ||
                          enabledElement.ValueKind != JsonValueKind.False;

            changed |= UpsertDetectedExtension(
                $"{idPrefix}-{SanitizeExtensionId(server.Name)}",
                server.Name,
                AiExtensionKind.Mcp,
                configPath,
                string.IsNullOrWhiteSpace(commandSummary) ? $"mcp:{server.Name}" : commandSummary,
                description,
                targetApp,
                enabled,
                configPath);
        }

        return changed;
    }

    private bool MergeDetectedJsonPluginList(
        JsonElement pluginElement,
        string configPath,
        string idPrefix,
        string targetApp,
        string description)
    {
        var changed = false;

        if (pluginElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in pluginElement.EnumerateArray())
            {
                var name = item.ValueKind == JsonValueKind.String
                    ? item.GetString()
                    : ReadJsonString(item, "name");

                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                changed |= UpsertDetectedExtension(
                    $"{idPrefix}-{SanitizeExtensionId(name)}",
                    name,
                    AiExtensionKind.Plugin,
                    configPath,
                    $"plugin:{name}",
                    description,
                    targetApp,
                    true,
                    configPath);
            }
        }
        else if (pluginElement.ValueKind == JsonValueKind.Object)
        {
            foreach (var item in pluginElement.EnumerateObject())
            {
                changed |= UpsertDetectedExtension(
                    $"{idPrefix}-{SanitizeExtensionId(item.Name)}",
                    item.Name,
                    AiExtensionKind.Plugin,
                    configPath,
                    $"plugin:{item.Name}",
                    description,
                    targetApp,
                    true,
                    configPath);
            }
        }

        return changed;
    }

    private bool MergeDetectedTomlMcpServers(
        string text,
        string configPath,
        string idPrefix,
        string targetApp,
        string description)
    {
        var sections = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        string? currentName = null;

        foreach (var rawLine in text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#'))
            {
                continue;
            }

            var sectionMatch = Regex.Match(
                line,
                @"^\[(?:mcp_servers|mcpServers)\.(?:""(?<name>[^""]+)""|'(?<name>[^']+)'|(?<name>[^\].]+))(?:\.[^\]]+)?\]$",
                RegexOptions.IgnoreCase);

            if (sectionMatch.Success)
            {
                currentName = sectionMatch.Groups["name"].Value.Trim();
                if (!sections.ContainsKey(currentName))
                {
                    sections[currentName] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }

                continue;
            }

            if (string.IsNullOrWhiteSpace(currentName) || !sections.TryGetValue(currentName, out var values))
            {
                continue;
            }

            var separatorIndex = line.IndexOf('=');
            if (separatorIndex <= 0)
            {
                continue;
            }

            var key = line[..separatorIndex].Trim();
            var value = CleanManifestValue(line[(separatorIndex + 1)..].Trim());
            values[key] = value;
        }

        var changed = false;

        foreach (var section in sections)
        {
            var commandSummary = BuildTomlMcpCommandSummary(section.Key, section.Value);
            var enabled = !section.Value.TryGetValue("enabled", out var enabledText) ||
                          !string.Equals(enabledText, "false", StringComparison.OrdinalIgnoreCase);

            changed |= UpsertDetectedExtension(
                $"{idPrefix}-{SanitizeExtensionId(section.Key)}",
                section.Key,
                AiExtensionKind.Mcp,
                configPath,
                commandSummary,
                description,
                targetApp,
                enabled,
                configPath);
        }

        return changed;
    }

    private static IEnumerable<(string Name, bool Enabled)> ReadCodexPluginSections(string text)
    {
        string? currentName = null;
        var enabled = true;

        foreach (var rawLine in text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            var line = rawLine.Trim();
            var sectionMatch = Regex.Match(
                line,
                @"^\[plugins\.(?:""(?<name>[^""]+)""|'(?<name>[^']+)'|(?<name>[^\]]+))\]$",
                RegexOptions.IgnoreCase);

            if (sectionMatch.Success)
            {
                if (!string.IsNullOrWhiteSpace(currentName))
                {
                    yield return (currentName, enabled);
                }

                currentName = sectionMatch.Groups["name"].Value.Trim();
                enabled = true;
                continue;
            }

            if (string.IsNullOrWhiteSpace(currentName))
            {
                continue;
            }

            var enabledMatch = Regex.Match(line, @"^enabled\s*=\s*(?<value>true|false)\s*$", RegexOptions.IgnoreCase);
            if (enabledMatch.Success)
            {
                enabled = string.Equals(enabledMatch.Groups["value"].Value, "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        if (!string.IsNullOrWhiteSpace(currentName))
        {
            yield return (currentName, enabled);
        }
    }

    private static string BuildTomlMcpCommandSummary(string name, IReadOnlyDictionary<string, string> values)
    {
        if (values.TryGetValue("url", out var url) && !string.IsNullOrWhiteSpace(url))
        {
            return url;
        }

        if (values.TryGetValue("command", out var command) && !string.IsNullOrWhiteSpace(command))
        {
            if (values.TryGetValue("args", out var args) && !string.IsNullOrWhiteSpace(args))
            {
                return $"{command} {args}";
            }

            return command;
        }

        return $"mcp:{name}";
    }

    private static string BuildJsonMcpCommandSummary(JsonElement serverElement)
    {
        if (TryGetJsonPropertyIgnoreCase(serverElement, "url", out var urlElement) &&
            urlElement.ValueKind == JsonValueKind.String)
        {
            return urlElement.GetString() ?? string.Empty;
        }

        if (!TryGetJsonPropertyIgnoreCase(serverElement, "command", out var commandElement))
        {
            return string.Empty;
        }

        var parts = new List<string>();
        AppendJsonCommandParts(commandElement, parts);

        if (TryGetJsonPropertyIgnoreCase(serverElement, "args", out var argsElement))
        {
            AppendJsonCommandParts(argsElement, parts);
        }

        return string.Join(" ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));
    }

    private static void AppendJsonCommandParts(JsonElement element, ICollection<string> parts)
    {
        if (element.ValueKind == JsonValueKind.String)
        {
            parts.Add(element.GetString() ?? string.Empty);
            return;
        }

        if (element.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        foreach (var item in element.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.String)
            {
                parts.Add(item.GetString() ?? string.Empty);
            }
        }
    }

    private static string ReadMcpCommandSummaryFromJsonFile(string path)
    {
        if (!File.Exists(path))
        {
            return string.Empty;
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(path));
            return document.RootElement.ValueKind == JsonValueKind.Object
                ? BuildJsonMcpCommandSummary(document.RootElement)
                : string.Empty;
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            return string.Empty;
        }
    }

    private static (string Name, string Owner, string Description) ReadLmStudioPluginManifest(string manifestPath)
    {
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(manifestPath));
            var root = document.RootElement;
            return (
                ReadJsonString(root, "name"),
                ReadJsonString(root, "owner"),
                ReadJsonString(root, "description"));
        }
        catch (Exception exception) when (exception is JsonException or IOException or UnauthorizedAccessException)
        {
            return (string.Empty, string.Empty, string.Empty);
        }
    }

    private static bool TryGetJsonPropertyIgnoreCase(JsonElement element, string propertyName, out JsonElement property)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var candidate in element.EnumerateObject())
            {
                if (string.Equals(candidate.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    property = candidate.Value;
                    return true;
                }
            }
        }

        property = default;
        return false;
    }

    private static string ReadJsonWithComments(string path)
    {
        var text = File.ReadAllText(path);
        return StripJsonLineComments(text);
    }

    private static string StripJsonLineComments(string text)
    {
        var result = new System.Text.StringBuilder(text.Length);
        var inString = false;
        var escaped = false;

        for (var index = 0; index < text.Length; index++)
        {
            var current = text[index];

            if (current == '"' && !escaped)
            {
                inString = !inString;
            }

            if (!inString && current == '/' && index + 1 < text.Length && text[index + 1] == '/')
            {
                while (index < text.Length && text[index] != '\r' && text[index] != '\n')
                {
                    index++;
                }

                if (index < text.Length)
                {
                    result.Append(text[index]);
                }

                escaped = false;
                continue;
            }

            result.Append(current);
            escaped = current == '\\' && !escaped;
            if (current != '\\')
            {
                escaped = false;
            }
        }

        return result.ToString();
    }

    private static IEnumerable<string> SafeEnumerateDirectories(string rootDirectory)
    {
        try
        {
            return Directory.EnumerateDirectories(rootDirectory);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return [];
        }
    }

    private static IEnumerable<string> SafeEnumerateFiles(string rootDirectory, string pattern)
    {
        var pending = new Stack<string>();
        pending.Push(rootDirectory);

        while (pending.Count > 0)
        {
            var current = pending.Pop();
            IEnumerable<string> files;
            IEnumerable<string> directories;

            try
            {
                files = Directory.EnumerateFiles(current, pattern);
                directories = Directory.EnumerateDirectories(current);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                continue;
            }

            foreach (var file in files)
            {
                yield return file;
            }

            foreach (var directory in directories)
            {
                if (directory.Contains($"{Path.DirectorySeparatorChar}node_modules", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                pending.Push(directory);
            }
        }
    }

    private static (string Name, string Description) ReadSkillManifest(string skillFilePath)
    {
        try
        {
            using var reader = File.OpenText(skillFilePath);
            var firstLine = reader.ReadLine();
            if (!string.Equals(firstLine?.Trim(), "---", StringComparison.Ordinal))
            {
                return (string.Empty, string.Empty);
            }

            var name = string.Empty;
            var description = string.Empty;

            for (var index = 0; index < 80; index++)
            {
                var line = reader.ReadLine();
                if (line is null || string.Equals(line.Trim(), "---", StringComparison.Ordinal))
                {
                    break;
                }

                var separatorIndex = line.IndexOf(':');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                var key = line[..separatorIndex].Trim();
                var value = CleanManifestValue(line[(separatorIndex + 1)..]);

                if (string.Equals(key, "name", StringComparison.OrdinalIgnoreCase))
                {
                    name = value;
                }
                else if (string.Equals(key, "description", StringComparison.OrdinalIgnoreCase))
                {
                    description = value;
                }
            }

            return (name, description);
        }
        catch
        {
            return (string.Empty, string.Empty);
        }
    }

    private static (string Name, string Description) ReadPluginManifest(string pluginJsonPath)
    {
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(pluginJsonPath));
            var root = document.RootElement;
            var name = ReadJsonString(root, "name");
            var description = ReadJsonString(root, "description");

            if (root.TryGetProperty("interface", out var interfaceElement) &&
                interfaceElement.ValueKind == JsonValueKind.Object)
            {
                var displayName = ReadJsonString(interfaceElement, "displayName");
                var shortDescription = ReadJsonString(interfaceElement, "shortDescription");

                if (!string.IsNullOrWhiteSpace(displayName))
                {
                    name = displayName;
                }

                if (!string.IsNullOrWhiteSpace(shortDescription))
                {
                    description = shortDescription;
                }
            }

            return (name, description);
        }
        catch
        {
            return (string.Empty, string.Empty);
        }
    }

    private static string ReadJsonString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? string.Empty
            : string.Empty;
    }

    private static string CleanManifestValue(string value)
    {
        return value.Trim().Trim('"', '\'');
    }

    private static string GetRelativePathSafe(string rootDirectory, string path)
    {
        try
        {
            return Path.GetRelativePath(rootDirectory, path);
        }
        catch
        {
            return path;
        }
    }

    private static string SanitizeExtensionId(string value)
    {
        var chars = value
            .Select(character => char.IsLetterOrDigit(character) ? char.ToLowerInvariant(character) : '-')
            .ToArray();
        return new string(chars).Trim('-');
    }

    private bool MergeFastDetectedExtensions()
    {
        var changed = false;

        changed |= MergeDetectedCodexSkills(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "skills"),
            "detected-skill");
        changed |= MergeDetectedCodexPlugins(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "plugins"),
            "detected-plugin");
        changed |= MergeDetectedCodexConfig(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex", "config.toml"));
        changed |= MergeDetectedOpenCodeConfig();
        changed |= MergeDetectedLmStudioConfig();

        return changed;
    }

    private void RemoveObsoleteDetectedToolEntries()
    {
        var obsoleteIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "detected-tool-opencode",
            "detected-tool-lm-studio",
            "detected-tool-ollama"
        };

        foreach (var item in AiExtensions.Where(item => obsoleteIds.Contains(item.Id)).ToList())
        {
            AiExtensions.Remove(item);
        }
    }

    private bool RemoveMissingDetectedExtensions()
    {
        var missingItems = AiExtensions
            .Where(item => item.IsDetected && !TryGetExtensionFileSystemTarget(item, forDelete: false, out _, out _))
            .ToList();

        foreach (var item in missingItems)
        {
            AiExtensions.Remove(item);
        }

        return missingItems.Count > 0;
    }

    private void SetExtensionStatus(string foreground, string text)
    {
        ExtensionStatusForeground = foreground;
        ExtensionStatusText = text;
    }

    private static AiExtensionKind ParseExtensionKind(string? value)
    {
        if (string.Equals(value, "MCP", StringComparison.OrdinalIgnoreCase))
        {
            return AiExtensionKind.Mcp;
        }

        if (string.Equals(value, "Skill", StringComparison.OrdinalIgnoreCase))
        {
            return AiExtensionKind.Skill;
        }

        return AiExtensionKind.Plugin;
    }

    private static string FormatExtensionKindValue(AiExtensionKind kind)
    {
        return kind switch
        {
            AiExtensionKind.Mcp => "MCP",
            AiExtensionKind.Skill => "Skill",
            _ => "Plugin"
        };
    }

    private static string NormalizeExtensionTargetAppValue(string? value)
    {
        if (string.Equals(value, "OpenCode", StringComparison.OrdinalIgnoreCase))
        {
            return "OpenCode";
        }

        if (string.Equals(value, "LM Studio", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, "LmStudio", StringComparison.OrdinalIgnoreCase))
        {
            return "LmStudio";
        }

        return "Codex";
    }

    private void LocalizeExtensionItem(AiExtensionItem item)
    {
        item.TargetApp = NormalizeExtensionTargetAppValue(item.TargetApp);
        item.KindDisplayLabel = item.Kind switch
        {
            AiExtensionKind.Mcp => Strings["ExtensionsKindMcp"],
            AiExtensionKind.Skill => Strings["ExtensionsKindSkill"],
            _ => Strings["ExtensionsKindPlugin"]
        };
        item.TargetAppDisplayLabel = item.TargetApp switch
        {
            "OpenCode" => Strings["ExtensionsTargetOpenCodeTab"],
            "LmStudio" => Strings["ExtensionsTargetLmStudioTab"],
            _ => Strings["ExtensionsTargetCodexTab"]
        };
        item.SourceDisplayLabel = item.IsDetected
            ? Strings["ExtensionsSourceDetected"]
            : item.IsPreset
                ? Strings["ExtensionsSourcePreset"]
                : Strings["ExtensionsSourceCustom"];
        item.InstallStateLabel = item.IsBusy
            ? Strings["ExtensionsInstallStateWorking"]
            : item.HasVerificationError
                ? Strings["ExtensionsInstallStateError"]
                : item.IsVerified
                    ? Strings["ExtensionsInstallStateVerified"]
                    : item.IsCustom
                        ? Strings["ExtensionsInstallStateSavedRecord"]
                        : item.IsInstalled
                            ? Strings["ExtensionsInstallStateConfigured"]
                            : Strings["ExtensionsInstallStateNotInstalled"];
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

    private void RefreshOllamaQuickChecks(CodexEnvironmentSnapshot? snapshot = null)
    {
        var environment = snapshot ?? _lastEnvironmentSnapshot;
        OllamaQuickChecks.Clear();

        if (environment is null)
        {
            return;
        }

        OllamaQuickChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOllamaAppTitle"],
                environment.OllamaAppAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                environment.OllamaAppAvailable
                    ? environment.OllamaAppDetail
                    : Strings["SetupOllamaAppDetailMissing"],
                environment.OllamaAppAvailable));
        OllamaQuickChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOllamaRuntimeTitle"],
                environment.OllamaAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                environment.OllamaAvailable
                    ? environment.OllamaExecutablePath
                    : Strings["SetupOllamaRuntimeDetailMissing"],
                environment.OllamaAvailable));
        OllamaQuickChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOllamaPathTitle"],
                environment.OllamaCommandVisible ? Strings["SetupBadgeReady"] : Strings["SetupBadgeRefresh"],
                environment.OllamaAvailable
                    ? environment.OllamaCommandVisible
                        ? Strings["SetupOllamaPathDetailReady"]
                        : Strings["SetupOllamaPathDetailRestart"]
                    : Strings["SetupOllamaPathDetailMissing"],
                environment.OllamaCommandVisible,
                isWarning: environment.OllamaAvailable && !environment.OllamaCommandVisible));
        OllamaQuickChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOllamaServerTitle"],
                environment.OllamaServerRunning ? Strings["SetupBadgeReady"] : Strings["SetupBadgeStart"],
                environment.OllamaAvailable
                    ? environment.OllamaServerRunning
                        ? Strings["SetupOllamaServerDetailRunning"]
                        : environment.OllamaTrayRunning || environment.OllamaAppAvailable
                            ? Strings["SetupOllamaServerDetailTrayOnly"]
                            : Strings["SetupOllamaServerDetailStopped"]
                    : Strings["SetupOllamaServerDetailMissing"],
                environment.OllamaServerRunning,
                isWarning: environment.OllamaAvailable && !environment.OllamaServerRunning));
        OllamaQuickChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOllamaModelsTitle"],
                environment.OllamaModelCount > 0 ? Strings["SetupBadgeReady"] : Strings["SetupBadgeDownload"],
                environment.OllamaAvailable
                    ? environment.OllamaModelCount > 0
                        ? Strings.Format("SetupOllamaModelsDetailCount", environment.OllamaModelCount, environment.OllamaModelsSummary)
                        : Strings["SetupOllamaModelsDetailEmpty"]
                    : Strings["SetupOllamaModelsDetailMissing"],
                environment.OllamaModelCount > 0,
                isWarning: environment.OllamaAvailable && environment.OllamaModelCount == 0));
    }

    private string BuildOllamaQuickGuidanceText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupOllamaGuidanceNoData"];
        }

        if (!snapshot.OllamaAvailable)
        {
            return Strings["SetupOllamaGuidanceInstallFirst"];
        }

        if (!snapshot.OllamaCommandVisible)
        {
            return Strings["SetupOllamaGuidanceRefreshPath"];
        }

        if (!snapshot.OllamaServerRunning)
        {
            return Strings["SetupOllamaGuidanceStartServer"];
        }

        if (snapshot.OllamaModelCount == 0)
        {
            return Strings["SetupOllamaGuidanceDownloadModel"];
        }

        return Strings["SetupOllamaGuidanceReady"];
    }

    private void RefreshLocalAiModelOptions(CodexEnvironmentSnapshot? snapshot = null)
    {
        snapshot ??= _lastEnvironmentSnapshot;
        var localModels = snapshot?.InstalledOllamaModels ??
                          new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var specs = new[]
        {
            new LocalModelSpec(
                "Qwen 3.5 0.8B",
                "qwen3.5:0.8b",
                "0.8B · 1.0 GB",
                Strings["SetupLocalAiModelLightDescription"],
                1.0,
                4,
                0),
            new LocalModelSpec(
                "Qwen 3.5 4B",
                "qwen3.5:4b",
                "4B · 3.4 GB",
                Strings["SetupLocalAiModelBalancedDescription"],
                3.4,
                8,
                4),
            new LocalModelSpec(
                "Qwen 3.5 9B",
                "qwen3.5:9b",
                "9B · 6.6 GB",
                Strings["SetupLocalAiModelStrongDescription"],
                6.6,
                16,
                8),
            new LocalModelSpec(
                "Qwen3 Coder 30B",
                "qwen3-coder:30b",
                "30B · 19 GB",
                Strings["SetupLocalAiModelCoderDescription"],
                19,
                32,
                20)
        };

        var recommendedTag = GetRecommendedLocalModelTag(snapshot);
        LocalAiModelOptions.Clear();

        foreach (var spec in specs)
        {
            var downloadBytes = Gibibytes(spec.DownloadSizeGiB);
            var minimumRamBytes = Gibibytes(spec.MinimumRamGiB);
            var recommendedVramBytes = Gibibytes(spec.RecommendedVramGiB);
            var hasKnownRam = snapshot?.TotalPhysicalMemoryBytes > 0;
            var hasKnownDisk = snapshot?.SystemDriveFreeBytes > 0;
            var ramFits = !hasKnownRam || snapshot!.TotalPhysicalMemoryBytes >= minimumRamBytes;
            var diskFits = !hasKnownDisk || snapshot!.SystemDriveFreeBytes >= downloadBytes + Gibibytes(1);
            var vramFits = recommendedVramBytes <= 0 ||
                           snapshot?.GpuMemoryBytes >= recommendedVramBytes;
            var isRecommended = string.Equals(spec.ModelTag, recommendedTag, StringComparison.OrdinalIgnoreCase);
            var fitStatusText = !diskFits
                ? Strings["SetupLocalAiFitNoDisk"]
                : !ramFits
                    ? Strings["SetupLocalAiFitNoRam"]
                    : vramFits
                        ? Strings["SetupLocalAiFitFast"]
                        : Strings["SetupLocalAiFitRam"];
            var fitStatusBrush = !diskFits || !ramFits
                ? "#B42318"
                : vramFits
                    ? "#1F7A52"
                    : "#B86E10";
            var recommendationText = isRecommended
                ? Strings["SetupLocalAiRecommendedForPc"]
                : Strings.Format(
                    "SetupLocalAiRequirementsFormat",
                    FormatByteSize(minimumRamBytes),
                    recommendedVramBytes > 0
                        ? FormatByteSize(recommendedVramBytes)
                        : Strings["SetupLocalAiGpuOptional"]);
            var isInstalled = localModels.TryGetValue(spec.ModelTag, out var installedSize);

            LocalAiModelOptions.Add(
                new LocalAiModelOption
                {
                    Name = spec.Name,
                    ModelTag = spec.ModelTag,
                    SizeLabel = spec.SizeLabel,
                    Description = spec.Description,
                    DownloadSizeBytes = downloadBytes,
                    MinimumRamBytes = minimumRamBytes,
                    RecommendedVramBytes = recommendedVramBytes,
                    IsRecommended = isRecommended,
                    CanInstall = ramFits && diskFits,
                    FitStatusText = fitStatusText,
                    FitStatusBrush = fitStatusBrush,
                    RecommendationText = recommendationText,
                    IsInstalled = isInstalled,
                    InstalledStatusText = isInstalled
                        ? Strings["SetupLocalAiInstalled"]
                        : Strings["SetupLocalAiNotInstalled"],
                    InstalledStatusBrush = isInstalled ? "#1F7A52" : "#5E6C76",
                    InstalledSizeText = isInstalled
                        ? Strings.Format("SetupLocalAiInstalledSize", installedSize ?? string.Empty)
                        : Strings["SetupLocalAiMissingSize"]
                });
        }
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
                    ? BuildOpenClawAgentDetailText(environment)
                    : Strings["SetupAiAgentMissingDetail"]
            });
    }

    private string BuildOpenClawAgentDetailText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return "OpenClaw";
        }

        var lines = new List<string>();

        if (!string.IsNullOrWhiteSpace(snapshot.OpenClawPrimaryModel))
        {
            lines.Add(Strings.Format("SetupOpenClawDetectedModelLine", snapshot.OpenClawPrimaryModel));
        }

        if (!string.IsNullOrWhiteSpace(snapshot.OpenClawToolProfile))
        {
            lines.Add(Strings.Format("SetupOpenClawDetectedProfileLine", snapshot.OpenClawToolProfile));
        }

        lines.Add(snapshot.OpenClawNodeInstalled
            ? Strings["SetupOpenClawNodeReadyShort"]
            : Strings["SetupOpenClawNodeMissingShort"]);
        lines.Add(snapshot.OpenClawBrowserReady
            ? Strings["SetupOpenClawBrowserReadyShort"]
            : Strings["SetupOpenClawBrowserNeedsWorkShort"]);

        return string.Join(Environment.NewLine, lines);
    }

    private void RefreshOpenClawSetupModes(CodexEnvironmentSnapshot? snapshot = null)
    {
        var environment = snapshot ?? _lastEnvironmentSnapshot;
        OpenClawSetupModes.Clear();

        var quickIsCurrent =
            environment?.OpenClawAvailable == true &&
            (string.IsNullOrWhiteSpace(environment.OpenClawToolProfile) ||
             environment.OpenClawToolProfile.Equals("minimal", StringComparison.OrdinalIgnoreCase));
        var advancedModelReady = HasStrongOpenClawModel(environment);
        var advancedIsCurrent =
            environment?.OpenClawAvailable == true &&
            environment.OpenClawToolProfile.Equals("coding", StringComparison.OrdinalIgnoreCase) &&
            advancedModelReady;
        var fullAssistantReady =
            environment?.OpenClawNodeInstalled == true &&
            environment.OpenClawBrowserReady;

        OpenClawSetupModes.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawModeQuickTitle"],
                quickIsCurrent
                    ? Strings["SetupOpenClawModeStatusCurrent"]
                    : Strings["SetupOpenClawModeStatusStable"],
                Strings["SetupOpenClawModeQuickDescription"],
                quickIsCurrent,
                isWarning: environment?.OpenClawAvailable == true && !quickIsCurrent));
        OpenClawSetupModes.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawModeAdvancedTitle"],
                advancedIsCurrent
                    ? Strings["SetupOpenClawModeStatusCurrent"]
                    : advancedModelReady
                        ? Strings["SetupOpenClawModeStatusRecommended"]
                        : Strings["SetupOpenClawModeStatusNeedsModel"],
                advancedModelReady
                    ? Strings["SetupOpenClawModeAdvancedDescription"]
                    : Strings["SetupOpenClawModeAdvancedNeedsModelDescription"],
                advancedIsCurrent,
                isWarning: !advancedIsCurrent && advancedModelReady));
        OpenClawSetupModes.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawModeFullTitle"],
                fullAssistantReady
                    ? Strings["SetupOpenClawModeStatusReady"]
                    : Strings["SetupOpenClawModeStatusRequiresSetup"],
                fullAssistantReady
                    ? Strings["SetupOpenClawModeFullDescription"]
                    : Strings["SetupOpenClawModeFullNeedsSetupDescription"],
                fullAssistantReady,
                isWarning: !fullAssistantReady));
    }

    private void RefreshOpenClawCapabilityChecks(CodexEnvironmentSnapshot? snapshot = null)
    {
        var environment = snapshot ?? _lastEnvironmentSnapshot;
        OpenClawCapabilityChecks.Clear();

        if (environment is null)
        {
            return;
        }

        var localChatReady = environment.OpenClawAvailable &&
                             environment.OpenClawConfigExists &&
                             !string.IsNullOrWhiteSpace(environment.OpenClawPrimaryModel) &&
                             (!environment.OpenClawPrimaryModel.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase) ||
                              environment.OllamaAvailable);

        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityLocalChat"],
                localChatReady ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                localChatReady
                    ? Strings.Format("SetupOpenClawCapabilityLocalChatReadyDetail", environment.OpenClawPrimaryModel)
                    : Strings["SetupOpenClawCapabilityLocalChatMissingDetail"],
                localChatReady,
                isWarning: environment.OpenClawAvailable && !localChatReady));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityWebSearch"],
                environment.OpenClawWebSearchEnabled ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                environment.OpenClawWebSearchEnabled
                    ? Strings["SetupOpenClawCapabilityWebSearchReadyDetail"]
                    : Strings["SetupOpenClawCapabilityWebSearchMissingDetail"],
                environment.OpenClawWebSearchEnabled,
                isWarning: environment.OpenClawAvailable && !environment.OpenClawWebSearchEnabled));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityBrowser"],
                environment.OpenClawBrowserReady ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                environment.OpenClawBrowserReady
                    ? Strings["SetupOpenClawCapabilityBrowserReadyDetail"]
                    : Strings.Format("SetupOpenClawCapabilityBrowserMissingDetail", environment.OpenClawBrowserDetail),
                environment.OpenClawBrowserReady,
                isWarning: environment.OpenClawBrowserCliAvailable && !environment.OpenClawBrowserReady));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityDesktopScreenshot"],
                Strings["SetupBadgeMissing"],
                Strings["SetupOpenClawCapabilityDesktopScreenshotDetail"],
                false,
                isWarning: environment.OpenClawBrowserReady));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityGuiAutomation"],
                Strings["SetupBadgeMissing"],
                Strings["SetupOpenClawCapabilityGuiAutomationDetail"],
                false,
                isWarning: environment.OpenClawNodeInstalled || environment.OpenClawBrowserReady));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityTelegram"],
                environment.OpenClawTelegramConfigured ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                environment.OpenClawTelegramConfigured
                    ? Strings["SetupOpenClawCapabilityTelegramReadyDetail"]
                    : Strings["SetupOpenClawCapabilityTelegramMissingDetail"],
                environment.OpenClawTelegramConfigured,
                isWarning: environment.OpenClawAvailable && !environment.OpenClawTelegramConfigured));
        OpenClawCapabilityChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupOpenClawCapabilityAdmin"],
                Strings["SetupBadgeMissing"],
                Strings["SetupOpenClawCapabilityAdminDetail"],
                false,
                isWarning: environment.OpenClawAvailable));
    }

    private string BuildOpenClawDetectedConfigText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupOpenClawDetectedConfigNotChecked"];
        }

        var model = string.IsNullOrWhiteSpace(snapshot.OpenClawPrimaryModel)
            ? Strings["SetupOpenClawValueNotSet"]
            : snapshot.OpenClawPrimaryModel;
        var profile = string.IsNullOrWhiteSpace(snapshot.OpenClawToolProfile)
            ? Strings["SetupOpenClawValueNotSet"]
            : snapshot.OpenClawToolProfile;
        var configPath = string.IsNullOrWhiteSpace(snapshot.OpenClawConfigPath)
            ? _environmentService.OpenClawConfigFilePath
            : snapshot.OpenClawConfigPath;

        return Strings.Format(
            "SetupOpenClawDetectedConfigFormat",
            configPath,
            model,
            profile,
            snapshot.OpenClawWebSearchEnabled ? Strings["Yes"] : Strings["No"],
            snapshot.OpenClawTelegramConfigured ? Strings["Yes"] : Strings["No"]);
    }

    private string BuildOpenClawRecommendationText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupOpenClawRecommendationNoData"];
        }

        if (!snapshot.OpenClawAvailable)
        {
            return Strings["SetupOpenClawRecommendationInstallFirst"];
        }

        if (!snapshot.OpenClawConfigExists)
        {
            return Strings["SetupOpenClawRecommendationCreateConfig"];
        }

        if (IsLightweightOpenClawModel(snapshot.OpenClawPrimaryModel) &&
            !snapshot.OpenClawToolProfile.Equals("minimal", StringComparison.OrdinalIgnoreCase))
        {
            return Strings["SetupOpenClawRecommendationDowngradeTools"];
        }

        if (IsLightweightOpenClawModel(snapshot.OpenClawPrimaryModel))
        {
            return Strings["SetupOpenClawRecommendationSmallModel"];
        }

        if (HasStrongOpenClawModel(snapshot) &&
            !snapshot.OpenClawPrimaryModel.Contains("qwen3.5", StringComparison.OrdinalIgnoreCase))
        {
            return Strings["SetupOpenClawRecommendationUpgradePrimaryModel"];
        }

        if (!snapshot.OpenClawNodeInstalled)
        {
            return Strings["SetupOpenClawRecommendationInstallNode"];
        }

        if (!snapshot.OpenClawBrowserReady)
        {
            return Strings["SetupOpenClawRecommendationFixBrowser"];
        }

        return Strings["SetupOpenClawRecommendationAdvancedReady"];
    }

    private static bool HasStrongOpenClawModel(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return false;
        }

        return snapshot.OpenClawPrimaryModel.Contains("qwen3.5", StringComparison.OrdinalIgnoreCase) ||
               snapshot.InstalledOllamaModels.Keys.Any(model =>
                   model.Contains("qwen3.5", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsLightweightOpenClawModel(string modelName)
    {
        if (string.IsNullOrWhiteSpace(modelName))
        {
            return false;
        }

        return modelName.Contains("qwen2.5:3b", StringComparison.OrdinalIgnoreCase) ||
               modelName.Contains("3b", StringComparison.OrdinalIgnoreCase) ||
               modelName.Contains("mini", StringComparison.OrdinalIgnoreCase);
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

    private bool ShouldUseDangerousBypassForNewSession()
    {
        return string.Equals(SelectedSandboxMode, "danger-full-access", StringComparison.OrdinalIgnoreCase) &&
               string.Equals(SelectedApprovalPolicy, "never", StringComparison.OrdinalIgnoreCase);
    }

    private NewSessionAccessLevel GetNewSessionAccessLevel()
    {
        if (ShouldUseDangerousBypassForNewSession())
        {
            return NewSessionAccessLevel.Critical;
        }

        if (string.Equals(SelectedSandboxMode, "danger-full-access", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(SelectedApprovalPolicy, "never", StringComparison.OrdinalIgnoreCase))
        {
            return NewSessionAccessLevel.Caution;
        }

        return NewSessionAccessLevel.Safe;
    }

    private void NotifyNewSessionAccessSummaryChanged()
    {
        OnPropertyChanged(nameof(NewSessionAccessSummaryTitle));
        OnPropertyChanged(nameof(NewSessionAccessSummaryText));
        OnPropertyChanged(nameof(NewSessionAccessSummaryBackground));
        OnPropertyChanged(nameof(NewSessionAccessSummaryForeground));
        OnPropertyChanged(nameof(NewSessionAccessSummaryBorder));
    }

    private SessionHealthLevel GetSelectedSessionHealthLevel()
    {
        var session = SelectedSession;
        if (session is null)
        {
            return SessionHealthLevel.Stable;
        }

        var fileSize = GetFileSizeSafely(session.FilePath);
        if (session.TotalMessageCount >= 300 ||
            session.ToolCallCount >= 100 ||
            fileSize >= 16L * 1024 * 1024)
        {
            return SessionHealthLevel.Overloaded;
        }

        if (session.TotalMessageCount >= 150 ||
            session.ToolCallCount >= 50 ||
            fileSize >= 8L * 1024 * 1024)
        {
            return SessionHealthLevel.Long;
        }

        return SessionHealthLevel.Stable;
    }

    private static long GetFileSizeSafely(string path)
    {
        try
        {
            return File.Exists(path) ? new FileInfo(path).Length : 0;
        }
        catch
        {
            return 0;
        }
    }

    private bool ConfirmDangerousNewSessionLaunch()
    {
        if (!ShouldUseDangerousBypassForNewSession())
        {
            return true;
        }

        return MessageBox.Show(
                   this,
                   Strings["NewSessionDangerousLaunchWarningMessage"],
                   Strings["NewSessionDangerousLaunchWarningTitle"],
                   MessageBoxButton.YesNo,
                   MessageBoxImage.Warning) == MessageBoxResult.Yes;
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

    private void SetHomeLaunchStatus(string background, string foreground, string text)
    {
        HomeLaunchStatusBackground = background;
        HomeLaunchStatusForeground = foreground;
        HomeLaunchStatusText = text;
    }

    private void SetSetupStatus(string foreground, string text)
    {
        SetupStatusForeground = foreground switch
        {
            "#F8E7D6" => "#1F6F4A",
            "#FFD6D6" => "#B42318",
            _ => foreground
        };
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
            ExportSessionsFeedSafe();

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

    private NewSessionLaunchOptions BuildNewSessionLaunchOptions(
        IReadOnlyList<string>? imagePaths = null,
        string? workingDirectoryOverride = null)
    {
        var useDangerousBypass = ShouldUseDangerousBypassForNewSession();

        return new NewSessionLaunchOptions
        {
            Prompt = NewSessionPrompt,
            WorkingDirectory = workingDirectoryOverride ?? GetNormalizedNewSessionWorkingDirectory(),
            ImagePaths = imagePaths ?? [],
            Model = NewSessionModel,
            Profile = NewSessionProfile,
            SandboxMode = SelectedSandboxMode,
            ApprovalPolicy = SelectedApprovalPolicy,
            LocalProvider = SelectedLocalProvider,
            UseSearch = NewSessionUseSearch,
            UseOss = NewSessionUseOss,
            UseDangerousBypass = useDangerousBypass
        };
    }

    private string? GetLaunchImagePath()
    {
        var clipboardPath = TryGetImagePathFromClipboard();

        if (!string.IsNullOrWhiteSpace(clipboardPath))
        {
            return clipboardPath;
        }

        var dialog = new OpenFileDialog
        {
            Filter = "Image files (*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            Title = Strings["ImagePickerTitle"]
        };

        return dialog.ShowDialog(this) == true ? dialog.FileName : null;
    }

    private string? TryGetImagePathFromClipboard()
    {
        try
        {
            if (Clipboard.ContainsFileDropList())
            {
                var file = Clipboard.GetFileDropList()
                    .Cast<string>()
                    .FirstOrDefault(IsSupportedImageFile);

                if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
                {
                    return file;
                }
            }

            if (!Clipboard.ContainsImage())
            {
                return null;
            }

            var image = Clipboard.GetImage();

            if (image is null)
            {
                return null;
            }

            var directory = Path.Combine(Path.GetTempPath(), "AIHelper", "clipboard-images");
            Directory.CreateDirectory(directory);

            var imagePath = Path.Combine(directory, $"clipboard-{DateTime.Now:yyyyMMdd-HHmmssfff}.png");
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(image));

            using var stream = File.Create(imagePath);
            encoder.Save(stream);
            return imagePath;
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(MainWindow), "Clipboard image export failed.", exception);
            return null;
        }
    }

    private static bool IsSupportedImageFile(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".png", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".webp", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".bmp", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".gif", StringComparison.OrdinalIgnoreCase);
    }

    private string GetNormalizedNewSessionWorkingDirectory()
    {
        return NewSessionWorkingDirectory.Trim();
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

    private async Task RefreshSetupSectionAsync(bool preserveDnsStatus, bool forceRefresh = false)
    {
        var setupRefreshed = await RefreshSetupStatusAsync(forceRefresh);
        var shouldRefreshDns =
            forceRefresh ||
            !preserveDnsStatus ||
            DnsAdapters.Count == 0 ||
            (SelectedAppSection == AppSection.Setup && IsSetupDnsSectionExpanded);

        if (setupRefreshed && shouldRefreshDns)
        {
            await RefreshDnsAdaptersAsync(preserveStatus: preserveDnsStatus);
        }
    }

    private async Task RefreshSettingsSectionAsync(bool forceRefresh = false)
    {
        await RefreshUpdateStatusAsync(forceRefresh);
    }

    private async Task RefreshUpdateStatusAsync(bool forceRefresh = false)
    {
        if (IsUpdateBusy)
        {
            return;
        }

        if (!forceRefresh &&
            _lastAppUpdateSnapshot is not null &&
            DateTime.UtcNow - _lastUpdateRefreshCompletedUtc < UpdateRefreshCacheDuration)
        {
            ApplyUpdateSnapshot(_lastAppUpdateSnapshot);
            return;
        }

        IsUpdateBusy = true;
        SetUpdateStatus("#F8E7D6", "UpdateStatusChecking");

        try
        {
            var snapshot = await _updateService.GetLatestReleaseAsync();
            _lastAppUpdateSnapshot = snapshot;
            _lastUpdateRefreshCompletedUtc = DateTime.UtcNow;
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
            else if (snapshot.IsCurrentVersionNewerThanLatest)
            {
                SetUpdateStatus(
                    "#F8E7D6",
                    "UpdateStatusAheadOfRelease",
                    snapshot.CurrentVersionDisplay,
                    snapshot.LatestVersionDisplay);
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

    private async Task<bool> RefreshSetupStatusAsync(bool forceRefresh = false)
    {
        if (IsSetupBusy)
        {
            return false;
        }

        if (!forceRefresh)
        {
            var fallbackInterval = IsSetupRefreshBoostActive
                ? SetupRefreshBusyInterval
                : SetupRefreshFallbackInterval;

            if (!_setupRefreshPending &&
                _lastSetupRefreshCompletedUtc != DateTime.MinValue &&
                DateTime.UtcNow - _lastSetupRefreshCompletedUtc < fallbackInterval)
            {
                return false;
            }
        }

        IsSetupBusy = true;
        SetSetupStatus("#F8E7D6", Strings["SetupStatusChecking"]);
        var stopwatch = Stopwatch.StartNew();

        try
        {
            if (forceRefresh)
            {
                _environmentService.InvalidateSnapshotCaches();
            }

            var snapshot = await Task.Run(_environmentService.GetEnvironmentSnapshot);
            _lastEnvironmentSnapshot = snapshot;
            _lastSetupRefreshCompletedUtc = DateTime.UtcNow;
            _setupRefreshPending = false;
            ApplySetupSnapshot(snapshot);
            SetSetupStatus(
                "#F8E7D6",
                IsSetupRefreshBoostActive
                    ? Strings["SetupStatusCheckedWatching"]
                    : Strings["SetupStatusChecked"]);
            _logService.Info(
                nameof(MainWindow),
                $"Setup status refreshed in {stopwatch.ElapsedMilliseconds} ms. Force={forceRefresh}; Boost={IsSetupRefreshBoostActive}.");
            return true;
        }
        catch (Exception exception)
        {
            SetSetupStatus("#FFD6D6", Strings.Format("SetupStatusFailed", exception.Message));
            MarkSetupRefreshPending();
            return false;
        }
        finally
        {
            IsSetupBusy = false;
            OnPropertyChanged(nameof(CanInstallBaseComponents));
            OnPropertyChanged(nameof(CanRepairWinget));
            OnPropertyChanged(nameof(CanLaunchNewSession));
            OnPropertyChanged(nameof(CanStartHomeSession));
            OnPropertyChanged(nameof(CanInstallCodexDesktopApp));
            OnPropertyChanged(nameof(CanOpenCodexDesktopStorePage));
            OnPropertyChanged(nameof(CanLaunchCodexLogin));
            OnPropertyChanged(nameof(CanInstallLocalAiTools));
            OnPropertyChanged(nameof(CanInstallLocalAiModels));
            OnPropertyChanged(nameof(CanManageCreativeAiTools));
            OnPropertyChanged(nameof(CanManageAiAgents));
            OnPropertyChanged(nameof(CanApplyOpenClawModes));
            OnPropertyChanged(nameof(CanInspectOpenClawStatus));
            OnPropertyChanged(nameof(CanInstallOpenClawNode));
            OnPropertyChanged(nameof(CanInspectOpenClawNode));
            OnPropertyChanged(nameof(CanInspectOpenClawBrowser));
            OnPropertyChanged(nameof(CanOpenOpenClawConfig));
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
        RefreshLocalAiModelOptions(snapshot);
        RefreshCreativeAiToolOptions(snapshot);
        RefreshAiAgentToolOptions(snapshot);
        RefreshOpenClawSetupModes(snapshot);
        RefreshOpenClawCapabilityChecks(snapshot);
        RefreshOllamaQuickChecks(snapshot);
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
                Strings["SetupCheckCodexDesktop"],
                snapshot.CodexDesktopAppAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.CodexDesktopAppAvailable
                    ? snapshot.CodexDesktopAppDetail
                    : Strings["SetupDetailCodexDesktopMissing"],
                snapshot.CodexDesktopAppAvailable));
        SetupCodexChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckCodex"],
                snapshot.CodexAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.CodexAvailable ? snapshot.CodexVersion : Strings["SetupDetailCodexMissing"],
                snapshot.CodexAvailable));
        SetupCodexChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOpenCode"],
                snapshot.OpenCodeAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.OpenCodeAvailable ? snapshot.OpenCodeDetail : Strings["SetupDetailOpenCodeMissing"],
                snapshot.OpenCodeAvailable));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOllamaApp"],
                snapshot.OllamaAppAvailable ? Strings["SetupBadgeInstalled"] : Strings["SetupBadgeMissing"],
                snapshot.OllamaAppAvailable
                    ? snapshot.OllamaAppDetail
                    : Strings["SetupDetailOllamaAppMissing"],
                snapshot.OllamaAppAvailable));
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
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOpenClawNode"],
                snapshot.OpenClawNodeInstalled ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                snapshot.OpenClawNodeDetail,
                snapshot.OpenClawNodeInstalled,
                isWarning: snapshot.OpenClawAvailable && !snapshot.OpenClawNodeInstalled));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOpenClawBrowser"],
                snapshot.OpenClawBrowserReady ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                snapshot.OpenClawBrowserDetail,
                snapshot.OpenClawBrowserReady,
                isWarning: snapshot.OpenClawBrowserCliAvailable && !snapshot.OpenClawBrowserReady));
        SetupLocalAiChecks.Add(
            CreateSetupCheckItem(
                Strings["SetupCheckOpenClawTelegram"],
                snapshot.OpenClawTelegramConfigured ? Strings["SetupBadgeReady"] : Strings["SetupBadgeMissing"],
                snapshot.OpenClawTelegramConfigured
                    ? Strings["SetupOpenClawCapabilityTelegramReadyDetail"]
                    : Strings["SetupOpenClawCapabilityTelegramMissingDetail"],
                snapshot.OpenClawTelegramConfigured,
                isWarning: snapshot.OpenClawAvailable && !snapshot.OpenClawTelegramConfigured));
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

        OnPropertyChanged(nameof(CanInstallCodexDesktopApp));
        OnPropertyChanged(nameof(CanOpenCodexDesktopStorePage));
        OnPropertyChanged(nameof(CanInstallLocalAiModels));
        OnPropertyChanged(nameof(CanLaunchOllamaApp));
        OnPropertyChanged(nameof(CanStartOllamaServer));
        OnPropertyChanged(nameof(CanStopOllamaProcesses));
        OnPropertyChanged(nameof(CanInstallStarterOllamaModel));
        OnPropertyChanged(nameof(OllamaQuickGuidanceText));
        OnPropertyChanged(nameof(CanManageCreativeAiTools));
        OnPropertyChanged(nameof(CanManageAiAgents));
        OnPropertyChanged(nameof(CanApplyOpenClawModes));
        OnPropertyChanged(nameof(CanInspectOpenClawStatus));
        OnPropertyChanged(nameof(CanInstallOpenClawNode));
        OnPropertyChanged(nameof(CanInspectOpenClawNode));
        OnPropertyChanged(nameof(CanInspectOpenClawBrowser));
        OnPropertyChanged(nameof(CanOpenOpenClawConfig));
        OnPropertyChanged(nameof(OpenClawDetectedConfigText));
        OnPropertyChanged(nameof(OpenClawRecommendationText));
        OnPropertyChanged(nameof(CanInstallOpenCode));
        OnPropertyChanged(nameof(CanLaunchOpenCode));
        OnPropertyChanged(nameof(CanUninstallOpenCode));
        OnPropertyChanged(nameof(OpenCodeSetupDetailText));
        OnPropertyChanged(nameof(CanUninstallOllama));
        OnPropertyChanged(nameof(CanUninstallLmStudio));
        RefreshSetupOverviewBindings();

        if (_beginnerOnboardingInProgress && IsHomeEnvironmentReady)
        {
            CompleteBeginnerOnboarding();
        }
    }

    private void RefreshSetupOverviewBindings()
    {
        OnPropertyChanged(nameof(SetupLiveStatusHintText));
        OnPropertyChanged(nameof(SetupRecommendedNextStepText));
        OnPropertyChanged(nameof(SetupCoreProgressText));
        OnPropertyChanged(nameof(SetupCodexProgressText));
        OnPropertyChanged(nameof(SetupLocalAiProgressText));
        OnPropertyChanged(nameof(SetupCoreNextStepText));
        OnPropertyChanged(nameof(SetupCodexNextStepText));
        OnPropertyChanged(nameof(SetupLocalAiNextStepText));
        OnPropertyChanged(nameof(SetupCoreSummaryBrush));
        OnPropertyChanged(nameof(SetupCodexSummaryBrush));
        OnPropertyChanged(nameof(SetupLocalAiSummaryBrush));
        OnPropertyChanged(nameof(IsHomeEnvironmentReady));
        OnPropertyChanged(nameof(CanStartHomeSession));
        OnPropertyChanged(nameof(HomeReadinessText));
        OnPropertyChanged(nameof(HomeReadinessBrush));
        OnPropertyChanged(nameof(HomeStartHelpText));
        OnPropertyChanged(nameof(BeginnerSetupLocalAiStatusText));
        OnPropertyChanged(nameof(BeginnerSetupLocalAiStatusBrush));
        OnPropertyChanged(nameof(HardwareOverviewText));
        OnPropertyChanged(nameof(HardwareRecommendationText));
        OnPropertyChanged(nameof(HardwareStatusBrush));
        OnPropertyChanged(nameof(LocalAiStorageSummaryText));
    }

    private string GetHardwareRecommendationText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupHardwarePending"];
        }

        if (snapshot.SystemDriveFreeBytes > 0 && snapshot.SystemDriveFreeBytes < Gibibytes(5))
        {
            return Strings["SetupHardwareRecommendationDiskLow"];
        }

        if (snapshot.TotalPhysicalMemoryBytes > 0 && snapshot.TotalPhysicalMemoryBytes < Gibibytes(8))
        {
            return Strings["SetupHardwareRecommendationLight"];
        }

        if (snapshot.GpuMemoryBytes >= Gibibytes(8))
        {
            return Strings["SetupHardwareRecommendationStrong"];
        }

        if (snapshot.GpuMemoryBytes >= Gibibytes(4))
        {
            return Strings["SetupHardwareRecommendationBalanced"];
        }

        return Strings["SetupHardwareRecommendationCpu"];
    }

    private static string GetHardwareStatusBrush(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return "#2D5366";
        }

        if ((snapshot.SystemDriveFreeBytes > 0 && snapshot.SystemDriveFreeBytes < Gibibytes(5)) ||
            (snapshot.TotalPhysicalMemoryBytes > 0 && snapshot.TotalPhysicalMemoryBytes < Gibibytes(8)))
        {
            return "#B42318";
        }

        return snapshot.GpuMemoryBytes >= Gibibytes(4) ? "#1F7A52" : "#B86E10";
    }

    private static string GetRecommendedLocalModelTag(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return "qwen3.5:4b";
        }

        if ((snapshot.TotalPhysicalMemoryBytes > 0 && snapshot.TotalPhysicalMemoryBytes < Gibibytes(8)) ||
            (snapshot.SystemDriveFreeBytes > 0 && snapshot.SystemDriveFreeBytes < Gibibytes(5)))
        {
            return "qwen3.5:0.8b";
        }

        if (snapshot.GpuMemoryBytes >= Gibibytes(8) &&
            snapshot.TotalPhysicalMemoryBytes >= Gibibytes(16) &&
            snapshot.SystemDriveFreeBytes >= Gibibytes(8))
        {
            return "qwen3.5:9b";
        }

        return "qwen3.5:4b";
    }

    private static long Gibibytes(double value)
    {
        var bytes = value * 1024d * 1024d * 1024d;
        return bytes >= long.MaxValue ? long.MaxValue : (long)bytes;
    }

    private static string FormatByteSize(long bytes)
    {
        if (bytes <= 0)
        {
            return "—";
        }

        var gibibytes = bytes / (1024d * 1024d * 1024d);
        return gibibytes >= 0.1
            ? $"{gibibytes:0.#} GB"
            : $"{bytes / (1024d * 1024d):0.#} MB";
    }

    private string GetSetupSectionProgressText(int readyCount, int totalCount)
    {
        if (_lastEnvironmentSnapshot is null)
        {
            return Strings["SetupOverviewUnavailable"];
        }

        return Strings.Format("SetupOverviewReadyFormat", readyCount, totalCount);
    }

    private static int GetCoreReadyCount(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return 0;
        }

        var ready = 0;
        ready += snapshot.WingetAvailable ? 1 : 0;
        ready += snapshot.NodeAvailable ? 1 : 0;
        ready += snapshot.NpmAvailable ? 1 : 0;
        ready += snapshot.GitAvailable ? 1 : 0;
        return ready;
    }

    private static int GetCodexReadyCount(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return 0;
        }

        var ready = 0;
        ready += snapshot.CodexDesktopAppAvailable || snapshot.CodexAvailable ? 1 : 0;
        ready += snapshot.OpenCodeAvailable ? 1 : 0;
        ready += snapshot.LoggedIn ? 1 : 0;
        ready += snapshot.SessionsFolderExists ? 1 : 0;
        return ready;
    }

    private static int GetLocalAiReadyCount(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return 0;
        }

        var ready = 0;
        ready += snapshot.OllamaAppAvailable ? 1 : 0;
        ready += snapshot.OllamaServerRunning ? 1 : 0;
        ready += snapshot.OllamaModelCount > 0 ? 1 : 0;
        ready += snapshot.OpenClawAvailable && snapshot.OpenClawNodeInstalled ? 1 : 0;
        return ready;
    }

    private static string GetSetupSummaryBrush(int readyCount, int totalCount)
    {
        if (readyCount >= totalCount)
        {
            return "#1F7A52";
        }

        if (readyCount > 0)
        {
            return "#B86E10";
        }

        return "#B42318";
    }

    private string GetSetupRecommendedNextStepText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupRecommendedNextStepPending"];
        }

        var coreNext = GetSetupCoreNextStepText(snapshot);

        if (!string.Equals(coreNext, Strings["SetupNextCoreDone"], StringComparison.Ordinal))
        {
            return coreNext;
        }

        var codexNext = GetSetupCodexNextStepText(snapshot);

        if (!string.Equals(codexNext, Strings["SetupNextCodexDone"], StringComparison.Ordinal))
        {
            return codexNext;
        }

        return GetSetupLocalAiNextStepText(snapshot);
    }

    private string GetSetupCoreNextStepText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupRecommendedNextStepPending"];
        }

        if (!snapshot.WingetAvailable)
        {
            return Strings["SetupNextCoreRepairWinget"];
        }

        if (!snapshot.NodeAvailable)
        {
            return Strings["SetupNextCoreInstallNode"];
        }

        if (!snapshot.NpmAvailable)
        {
            return Strings["SetupNextCoreRepairNpm"];
        }

        if (!snapshot.GitAvailable)
        {
            return Strings["SetupNextCoreInstallGit"];
        }

        return Strings["SetupNextCoreDone"];
    }

    private string GetSetupCodexNextStepText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupRecommendedNextStepPending"];
        }

        if (!snapshot.CodexDesktopAppAvailable && !snapshot.CodexAvailable)
        {
            return Strings["SetupNextCodexInstallAny"];
        }

        if (!snapshot.CodexAvailable)
        {
            return Strings["SetupNextCodexInstallCli"];
        }

        if (!snapshot.LoggedIn)
        {
            return Strings["SetupNextCodexLogin"];
        }

        if (!snapshot.SessionsFolderExists)
        {
            return Strings["SetupNextCodexOpenFirstSession"];
        }

        if (!snapshot.OpenCodeAvailable)
        {
            return Strings["SetupNextCodexInstallOpenCode"];
        }

        return Strings["SetupNextCodexDone"];
    }

    private string GetSetupLocalAiNextStepText(CodexEnvironmentSnapshot? snapshot)
    {
        if (snapshot is null)
        {
            return Strings["SetupRecommendedNextStepPending"];
        }

        if (!snapshot.OllamaAppAvailable)
        {
            return Strings["SetupNextLocalInstallOllama"];
        }

        if (!snapshot.OllamaServerRunning)
        {
            return Strings["SetupNextLocalStartOllama"];
        }

        if (snapshot.OllamaModelCount == 0)
        {
            return Strings["SetupNextLocalInstallStarterModel"];
        }

        if (!snapshot.OpenClawAvailable)
        {
            return Strings["SetupNextLocalInstallOpenClaw"];
        }

        if (!snapshot.OpenClawNodeInstalled)
        {
            return Strings["SetupNextLocalInstallOpenClawNode"];
        }

        return Strings["SetupNextLocalDone"];
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
                ? "#8A4B08"
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

    private async Task RefreshSessionsAsync(bool isAutomaticRefresh, bool forceRefresh = false)
    {
        EnsureSessionWatchersInitialized();

        if (_isRefreshing)
        {
            return;
        }

        if (!forceRefresh &&
            !_sessionRefreshPending &&
            _lastSessionRefreshCompletedUtc != DateTime.MinValue &&
            DateTime.UtcNow - _lastSessionRefreshCompletedUtc < SessionRefreshFallbackInterval)
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
        var stopwatch = Stopwatch.StartNew();

        try
        {
            var currentLanguage = Strings.CurrentLanguage;
            var refreshedSessions = await Task.Run(() => _sessionService.GetSessions(currentLanguage));
            ApplySessions(refreshedSessions);

            _lastSessionRefreshCompletedUtc = DateTime.UtcNow;
            _sessionRefreshPending = false;
            _lastUpdatedAtLocal = DateTime.Now;
            LastUpdatedText = Strings.Format("LastUpdated", _lastUpdatedAtLocal.Value.ToString("dd.MM.yyyy HH:mm:ss"));
            SetStatus(
                "#F8E7D6",
                refreshedSessions.Count == 0 ? "StatusNoSessions" : "StatusLoadedCount",
                refreshedSessions.Count);
            _logService.Info(
                nameof(MainWindow),
                $"Sessions refreshed in {stopwatch.ElapsedMilliseconds} ms. Force={forceRefresh}; Auto={isAutomaticRefresh}; Count={refreshedSessions.Count}.");
        }
        catch (Exception exception)
        {
            SetStatus("#FFD6D6", "StatusError", exception.Message);
            MarkSessionsRefreshPending();
        }
        finally
        {
            IsLoading = false;
            _isRefreshing = false;
        }
    }

    private void UpdateRefreshTimer()
    {
        if (AutoRefreshEnabled && SelectedAppSection == AppSection.Sessions && IsActive)
        {
            _refreshTimer.Start();
        }
        else
        {
            _refreshTimer.Stop();
        }
    }

    private void UpdateSetupRefreshTimer()
    {
        ExpireSetupRefreshBoostIfNeeded();
        _setupRefreshTimer.Interval = IsSetupRefreshBoostActive ? SetupRefreshBusyInterval : SetupRefreshNormalInterval;

        if ((SelectedAppSection == AppSection.Setup || IsSetupRefreshBoostActive) && IsActive)
        {
            _setupRefreshTimer.Start();
        }
        else
        {
            _setupRefreshTimer.Stop();
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

    private bool IsCompactWindowLayout =>
        ActualWidth > 0 && ActualWidth < 1320;

    private bool IsWideWindowLayout =>
        WindowState == WindowState.Maximized || ActualWidth >= 1720;

    private void ScheduleAdaptiveLayoutRefresh()
    {
        if (!IsLoaded)
        {
            RefreshAdaptiveLayoutBindings();
            return;
        }

        _layoutRefreshTimer.Stop();
        _layoutRefreshTimer.Start();
    }

    private void RefreshAdaptiveLayoutBindings()
    {
        OnPropertyChanged(nameof(AppOuterMargin));
        OnPropertyChanged(nameof(AppContentMaxWidth));
        OnPropertyChanged(nameof(AppContentWidth));
        OnPropertyChanged(nameof(ShellSidebarColumnWidth));
        OnPropertyChanged(nameof(ShellMainGapColumnWidth));
        OnPropertyChanged(nameof(SectionRailGapColumnWidth));
        OnPropertyChanged(nameof(SessionsHeaderSearchColumnWidth));
        OnPropertyChanged(nameof(SessionsActionButtonColumnWidth));
        OnPropertyChanged(nameof(SessionsDetailColumnWidth));
        OnPropertyChanged(nameof(NewSessionAsideColumnWidth));
        OnPropertyChanged(nameof(SettingsAsideColumnWidth));
        OnPropertyChanged(nameof(HomeTitleFontSize));
        OnPropertyChanged(nameof(HomePromptHeight));
        OnPropertyChanged(nameof(HomeSafetyColumnWidth));
    }

    private readonly record struct LocalModelSpec(
        string Name,
        string ModelTag,
        string SizeLabel,
        string Description,
        double DownloadSizeGiB,
        double MinimumRamGiB,
        double RecommendedVramGiB);

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

    private enum SessionHealthLevel
    {
        Stable,
        Long,
        Overloaded
    }

    private enum NewSessionAccessLevel
    {
        Safe,
        Caution,
        Critical
    }
}






