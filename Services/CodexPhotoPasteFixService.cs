using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace LaptopSessionViewer.Services;

public sealed class CodexPhotoPasteFixService : IDisposable
{
    private const int WhKeyboardLl = 13;
    private const int WmKeyDown = 0x0100;
    private const int WmKeyUp = 0x0101;
    private const int WmSysKeyDown = 0x0104;
    private const int WmSysKeyUp = 0x0105;
    private const int WmInputLangChangeRequest = 0x0050;
    private const uint SmtoAbortIfHung = 0x0002;
    private const uint KlfActivate = 0x00000001;
    private const uint KlfSubstituteOk = 0x00000002;
    private const int VkControl = 0x11;
    private const int VkLControl = 0xA2;
    private const int VkRControl = 0xA3;
    private const int VkMenu = 0x12;
    private const int VkLMenu = 0xA4;
    private const int VkRMenu = 0xA5;
    private const int VkV = 0x56;
    private const uint CfBitmap = 2;
    private const uint CfDib = 8;
    private const uint CfText = 1;
    private const uint CfHdrop = 15;
    private const uint CfDibV5 = 17;
    private const uint CfUnicodeText = 13;
    private const uint MapvkVkToVsc = 0;
    private const uint KeyeventfExtendedKey = 0x0001;
    private const uint KeyeventfKeyUp = 0x0002;
    private const uint LlkhfInjected = 0x00000010;
    private const uint LlkhfAltdown = 0x00000020;
    private const string EnglishUnitedStatesLayoutId = "00000409";

    private static readonly uint VScanCode = MapVirtualKey(VkV, MapvkVkToVsc);
    private static readonly TimeSpan ClipboardImageWaitTimeout = TimeSpan.FromMilliseconds(650);
    private static readonly TimeSpan ClipboardImagePollInterval = TimeSpan.FromMilliseconds(25);
    private static readonly TimeSpan LayoutSettleDelay = TimeSpan.FromMilliseconds(120);
    private static readonly TimeSpan PostPasteDelay = TimeSpan.FromMilliseconds(180);
    private static readonly TimeSpan LayoutRestoreDelay = TimeSpan.FromMilliseconds(90);
    private static readonly TimeSpan LayoutRestoreRetryDelay = TimeSpan.FromMilliseconds(140);
    private static readonly TimeSpan AltPasteDispatchDelay = TimeSpan.FromMilliseconds(35);
    private static readonly TimeSpan CtrlPasteReleaseDelay = TimeSpan.FromMilliseconds(45);
    private static readonly TimeSpan PasteActionCooldown = TimeSpan.FromSeconds(1);
    private static readonly TimeSpan PendingStateTimeout = TimeSpan.FromSeconds(2);
    private static readonly HashSet<string> SupportedTerminalProcesses =
    [
        "windowsterminal",
        "wt",
        "cmd",
        "powershell",
        "pwsh",
        "conhost",
        "openconsole",
        "codex",
        "node"
    ];
    private static readonly HashSet<string> SupportedWindowClasses =
    [
        "ConsoleWindowClass",
        "CASCADIA_HOSTING_WINDOW_CLASS"
    ];

    private readonly AppLogService _logService;
    private readonly LowLevelKeyboardProc _hookCallback;
    private IntPtr _englishLayoutHandle;
    private IntPtr _hookHandle;
    private bool _isEnabled;
    private bool _suppressNextVKeyUp;
    private CodexWindowContext? _pendingCtrlPasteContext;
    private DateTime _pendingCtrlPasteStartedUtc = DateTime.MinValue;
    private DateTime _lastInterceptionLogUtc = DateTime.MinValue;
    private DateTime _lastPasteActionUtc = DateTime.MinValue;

    public CodexPhotoPasteFixService(AppLogService logService)
    {
        _logService = logService;
        _hookCallback = HookCallback;
    }

    public bool IsEnabled => _isEnabled;

    public void UpdateConfiguration(bool enabled)
    {
        if (enabled)
        {
            InstallHook();
            _isEnabled = true;
            _logService.Info(nameof(CodexPhotoPasteFixService), "Photo paste fix enabled.");
            return;
        }

        _isEnabled = false;
        RemoveHook();
        _logService.Info(nameof(CodexPhotoPasteFixService), "Photo paste fix disabled.");
    }

    public void Dispose()
    {
        RemoveHook();
        GC.SuppressFinalize(this);
    }

    private void InstallHook()
    {
        if (_hookHandle != IntPtr.Zero)
        {
            return;
        }

        _englishLayoutHandle = LoadKeyboardLayout(EnglishUnitedStatesLayoutId, KlfActivate | KlfSubstituteOk);

        if (_englishLayoutHandle == IntPtr.Zero)
        {
            throw new InvalidOperationException("English keyboard layout 00000409 could not be loaded.");
        }

        using var currentProcess = Process.GetCurrentProcess();
        using var currentModule = currentProcess.MainModule;
        var moduleHandle = GetModuleHandle(currentModule?.ModuleName);

        _hookHandle = SetWindowsHookEx(WhKeyboardLl, _hookCallback, moduleHandle, 0);

        if (_hookHandle == IntPtr.Zero)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to install the global keyboard hook.");
        }
    }

    private void RemoveHook()
    {
        if (_hookHandle == IntPtr.Zero)
        {
            return;
        }

        UnhookWindowsHookEx(_hookHandle);
        _hookHandle = IntPtr.Zero;
        _suppressNextVKeyUp = false;
        _pendingCtrlPasteContext = null;
        _pendingCtrlPasteStartedUtc = DateTime.MinValue;
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code < 0 || !_isEnabled)
        {
            return CallNextHookEx(_hookHandle, code, wParam, lParam);
        }

        try
        {
            var message = unchecked((int)wParam.ToInt64());
            var keyInfo = Marshal.PtrToStructure<KbdllHookStruct>(lParam);

            if ((keyInfo.flags & LlkhfInjected) != 0)
            {
                return CallNextHookEx(_hookHandle, code, wParam, lParam);
            }

            ExpirePendingStatesIfNeeded();
            HandlePendingCtrlPasteRelease(message, keyInfo);

            if (_suppressNextVKeyUp && IsHandledVKeyUp(message, keyInfo))
            {
                _suppressNextVKeyUp = false;
                return (IntPtr)1;
            }

            if (!IsVKeyDown(message, keyInfo))
            {
                return CallNextHookEx(_hookHandle, code, wParam, lParam);
            }

            if (!TryGetCodexTerminalContext(out var context))
            {
                return CallNextHookEx(_hookHandle, code, wParam, lParam);
            }

            var altDown = (keyInfo.flags & LlkhfAltdown) != 0 || IsKeyPressed(VkMenu);
            var controlState = GetControlStateSnapshot();

            if (altDown && !controlState.AnyPressed)
            {
                if (IsEnglishLayout(context.KeyboardLayout))
                {
                    return CallNextHookEx(_hookHandle, code, wParam, lParam);
                }

                if (!ClipboardHasImagePayload())
                {
                    return CallNextHookEx(_hookHandle, code, wParam, lParam);
                }

                if (!TryEnterPasteCooldown())
                {
                    _suppressNextVKeyUp = true;
                    return (IntPtr)1;
                }

                _suppressNextVKeyUp = true;
                QueueOnUiThread(async () =>
                {
                    try
                    {
                        await Task.Delay(AltPasteDispatchDelay);

                        if (GetForegroundWindow() != context.WindowHandle)
                        {
                            _logService.Info(
                                nameof(CodexPhotoPasteFixService),
                                "Skipped deferred Alt+V image paste because the target terminal lost focus.");
                            return;
                        }

                        var altStillPressed = IsKeyPressed(VkMenu) || IsKeyPressed(VkLMenu) || IsKeyPressed(VkRMenu);
                        TrySendEnglishAltV(context, explicitAltInjection: !altStillPressed);
                    }
                    catch (Exception exception)
                    {
                        _logService.Error(nameof(CodexPhotoPasteFixService), "Deferred Alt+V image paste failed.", exception);
                    }
                });
                return (IntPtr)1;
            }

            if (controlState.AnyPressed && !altDown && ShouldTranslateCtrlVToAltV())
            {
                if (!TryEnterPasteCooldown())
                {
                    _suppressNextVKeyUp = true;
                    return (IntPtr)1;
                }

                _suppressNextVKeyUp = true;
                _pendingCtrlPasteContext = context;
                _pendingCtrlPasteStartedUtc = DateTime.UtcNow;
                return (IntPtr)1;
            }
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(CodexPhotoPasteFixService), "Photo paste hook failed.", exception);
        }

        return CallNextHookEx(_hookHandle, code, wParam, lParam);
    }

    private void ExpirePendingStatesIfNeeded()
    {
        var now = DateTime.UtcNow;


        if (_pendingCtrlPasteContext is not null &&
            now - _pendingCtrlPasteStartedUtc > PendingStateTimeout)
        {
            _pendingCtrlPasteContext = null;
            _pendingCtrlPasteStartedUtc = DateTime.MinValue;
        }
    }


    private void HandlePendingCtrlPasteRelease(int message, KbdllHookStruct keyInfo)
    {
        if (_pendingCtrlPasteContext is null || !IsControlKeyUp(message, keyInfo))
        {
            return;
        }

        var context = _pendingCtrlPasteContext.Value;
        _pendingCtrlPasteContext = null;
        _pendingCtrlPasteStartedUtc = DateTime.MinValue;

        QueueOnUiThread(async () =>
        {
            try
            {
                await Task.Delay(CtrlPasteReleaseDelay);

                if (GetForegroundWindow() != context.WindowHandle)
                {
                    _logService.Info(
                        nameof(CodexPhotoPasteFixService),
                        "Skipped deferred Ctrl+V image paste because the target terminal lost focus.");
                    return;
                }

                TrySendEnglishAltV(context, explicitAltInjection: true);
            }
            catch (Exception exception)
            {
                _logService.Error(nameof(CodexPhotoPasteFixService), "Deferred Ctrl+V image paste failed.", exception);
            }
        });
    }

    private void QueueOnUiThread(Func<Task> action)
    {
        var dispatcher = Application.Current?.Dispatcher;

        if (dispatcher is null)
        {
            _ = Task.Run(action);
            return;
        }

        if (dispatcher.CheckAccess())
        {
            _ = action();
            return;
        }

        _ = dispatcher.BeginInvoke(async () => await action(), DispatcherPriority.Background);
    }

    private bool TryEnterPasteCooldown()
    {
        var now = DateTime.UtcNow;

        if (now - _lastPasteActionUtc < PasteActionCooldown)
        {
            return false;
        }

        _lastPasteActionUtc = now;
        return true;
    }

    private bool TrySendEnglishAltV(CodexWindowContext context, bool explicitAltInjection)
    {
        try
        {
            var needsLayoutRestore = context.KeyboardLayout != IntPtr.Zero &&
                                     _englishLayoutHandle != IntPtr.Zero &&
                                     !HasSameKeyboardLayout(context.KeyboardLayout, _englishLayoutHandle);

            if (needsLayoutRestore &&
                !TryApplyKeyboardLayout(
                    context,
                    _englishLayoutHandle,
                    explicitAltInjection ? "PrepareInjectedAltV" : "PrepareHeldAltV",
                    attemptCount: 2))
            {
                return false;
            }

            var restoreOperation = explicitAltInjection ? "RestoreAfterInjectedAltV" : "RestoreAfterHeldAltV";

            try
            {
                SendInput(BuildPasteSequence(explicitAltInjection));
                Thread.Sleep(PostPasteDelay);
            }
            finally
            {
                if (needsLayoutRestore)
                {
                    Thread.Sleep(LayoutRestoreDelay);

                    if (!TryRestoreKeyboardLayout(context, restoreOperation))
                    {
                        Thread.Sleep(LayoutRestoreRetryDelay);
                        TryRestoreKeyboardLayout(context, restoreOperation + "Retry");
                    }
                }
            }

            LogInterception(context, explicitAltInjection);
            return true;
        }
        catch (Exception exception)
        {
            _logService.Error(nameof(CodexPhotoPasteFixService), "Failed to inject Alt+V for Codex.", exception);
            return false;
        }
    }

    private static IReadOnlyList<Input> BuildPasteSequence(bool explicitAltInjection)
    {
        var inputs = new List<Input>();

        if (explicitAltInjection)
        {
            inputs.Add(CreateVirtualKeyInput(VkMenu, keyUp: false));
        }

        inputs.Add(CreateVirtualKeyInput(VkV, keyUp: false));
        inputs.Add(CreateVirtualKeyInput(VkV, keyUp: true));

        if (explicitAltInjection)
        {
            inputs.Add(CreateVirtualKeyInput(VkMenu, keyUp: true));
        }

        return inputs;
    }

    private static Input CreateVirtualKeyInput(ushort virtualKey, bool keyUp, bool extended = false)
    {
        var flags = keyUp ? KeyeventfKeyUp : 0u;

        if (extended)
        {
            flags |= KeyeventfExtendedKey;
        }

        return new Input
        {
            type = 1,
            union = new InputUnion
            {
                ki = new KeybdInput
                {
                    wVk = virtualKey,
                    wScan = 0,
                    dwFlags = flags,
                    dwExtraInfo = IntPtr.Zero,
                    time = 0
                }
            }
        };
    }

    private static bool ClipboardHasImagePayload()
    {
        if (IsClipboardFormatAvailable(CfBitmap) ||
            IsClipboardFormatAvailable(CfDib) ||
            IsClipboardFormatAvailable(CfDibV5))
        {
            return true;
        }

        if (!IsClipboardFormatAvailable(CfHdrop))
        {
            return false;
        }

        if (!OpenClipboard(IntPtr.Zero))
        {
            return false;
        }

        try
        {
            var hDrop = GetClipboardData(CfHdrop);

            if (hDrop == IntPtr.Zero)
            {
                return false;
            }

            var fileCount = DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);

            for (uint index = 0; index < fileCount; index++)
            {
                var pathLength = DragQueryFile(hDrop, index, null, 0);
                var builder = new StringBuilder((int)pathLength + 1);
                DragQueryFile(hDrop, index, builder, (uint)builder.Capacity);

                if (IsSupportedImageFile(builder.ToString()))
                {
                    return true;
                }
            }

            return false;
        }
        finally
        {
            CloseClipboard();
        }
    }

    private static bool ClipboardContainsTextPayload()
    {
        return IsClipboardFormatAvailable(CfUnicodeText) || IsClipboardFormatAvailable(CfText);
    }

    private static bool ShouldTranslateCtrlVToAltV()
    {
        if (ClipboardHasImagePayload())
        {
            return true;
        }

        if (ClipboardContainsTextPayload())
        {
            return false;
        }

        var start = Environment.TickCount64;

        while (Environment.TickCount64 - start < ClipboardImageWaitTimeout.TotalMilliseconds)
        {
            Thread.Sleep(ClipboardImagePollInterval);

            if (ClipboardHasImagePayload())
            {
                return true;
            }

            if (ClipboardContainsTextPayload())
            {
                return false;
            }
        }

        return false;
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

    private static bool TryGetCodexTerminalContext(out CodexWindowContext context)
    {
        context = default;

        var foregroundWindow = GetForegroundWindow();

        if (foregroundWindow == IntPtr.Zero)
        {
            return false;
        }

        var title = GetWindowTitle(foregroundWindow);
        var windowClass = GetWindowClassName(foregroundWindow);
        var windowThreadId = GetWindowThreadProcessId(foregroundWindow, out var processId);

        if (windowThreadId == 0 || processId == 0)
        {
            return false;
        }

        string processName;

        try
        {
            using var process = Process.GetProcessById((int)processId);
            processName = process.ProcessName.Trim();
        }
        catch
        {
            return false;
        }

        var isSupportedProcess = SupportedTerminalProcesses.Contains(processName, StringComparer.OrdinalIgnoreCase);
        var isSupportedWindowClass = SupportedWindowClasses.Contains(windowClass, StringComparer.OrdinalIgnoreCase);

        if (!isSupportedProcess && !isSupportedWindowClass)
        {
            return false;
        }

        context = new CodexWindowContext
        {
            WindowHandle = foregroundWindow,
            ThreadId = windowThreadId,
            KeyboardLayout = GetKeyboardLayout(windowThreadId),
            ProcessName = processName,
            WindowTitle = title,
            WindowClass = windowClass
        };

        return true;
    }

    private static string GetWindowTitle(IntPtr windowHandle)
    {
        var length = GetWindowTextLength(windowHandle);

        if (length <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(length + 1);
        GetWindowText(windowHandle, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetWindowClassName(IntPtr windowHandle)
    {
        var builder = new StringBuilder(256);
        return GetClassName(windowHandle, builder, builder.Capacity) > 0
            ? builder.ToString()
            : string.Empty;
    }

    private void LogInterception(CodexWindowContext context, bool explicitAltInjection)
    {
        var now = DateTime.UtcNow;

        if (now - _lastInterceptionLogUtc < TimeSpan.FromSeconds(3))
        {
            return;
        }

        _lastInterceptionLogUtc = now;
        var mode = explicitAltInjection ? "Ctrl+V => temporary English Alt+V" : "Alt+V on temporary English layout";
        _logService.Info(
            nameof(CodexPhotoPasteFixService),
            $"Applied image paste fix via {mode}. Process={context.ProcessName}; Class={context.WindowClass}; Title={context.WindowTitle}");
    }

    private bool TryApplyKeyboardLayout(CodexWindowContext context, IntPtr targetLayout, string operation, int attemptCount)
    {
        if (targetLayout == IntPtr.Zero)
        {
            _logService.Info(nameof(CodexPhotoPasteFixService), $"{operation}: skipped because target layout is empty.");
            return false;
        }

        for (var attempt = 1; attempt <= attemptCount; attempt++)
        {
            SwitchKeyboardLayout(context, targetLayout);
            Thread.Sleep(LayoutSettleDelay);

            var actualLayout = GetKeyboardLayout(context.ThreadId);
            var exactMatch = HasSameKeyboardLayout(actualLayout, targetLayout);
            var languageMatch = HasSameKeyboardLanguage(actualLayout, targetLayout);

            _logService.Info(
                nameof(CodexPhotoPasteFixService),
                $"{operation}: attempt={attempt}; target={FormatKeyboardLayout(targetLayout)}; actual={FormatKeyboardLayout(actualLayout)}; exactMatch={exactMatch}; languageMatch={languageMatch}; window={context.WindowHandle}; thread={context.ThreadId}");

            if (exactMatch || languageMatch)
            {
                return true;
            }

            Thread.Sleep(LayoutRestoreRetryDelay);
        }

        return false;
    }

    private bool TryRestoreKeyboardLayout(CodexWindowContext context, string operation)
    {
        return TryApplyKeyboardLayout(context, context.KeyboardLayout, operation, attemptCount: 3);
    }

    private static void SwitchKeyboardLayout(CodexWindowContext context, IntPtr keyboardLayout)
    {
        var resolvedKeyboardLayout = ResolveKeyboardLayoutHandle(keyboardLayout);

        if (resolvedKeyboardLayout == IntPtr.Zero)
        {
            return;
        }

        var currentThreadId = GetCurrentThreadId();
        var attached = false;

        if (currentThreadId != context.ThreadId)
        {
            attached = AttachThreadInput(currentThreadId, context.ThreadId, true);
        }

        try
        {
            ActivateKeyboardLayout(resolvedKeyboardLayout, 0);
            SendMessageTimeout(
                context.WindowHandle,
                WmInputLangChangeRequest,
                IntPtr.Zero,
                resolvedKeyboardLayout,
                SmtoAbortIfHung,
                100,
                out _);
            WaitForKeyboardLayout(context.ThreadId, resolvedKeyboardLayout);
        }
        finally
        {
            if (attached)
            {
                AttachThreadInput(currentThreadId, context.ThreadId, false);
            }
        }
    }

    private static void WaitForKeyboardLayout(uint threadId, IntPtr expectedLayout)
    {
        if (threadId == 0 || expectedLayout == IntPtr.Zero)
        {
            return;
        }

        var start = Environment.TickCount64;

        while (Environment.TickCount64 - start < 120)
        {
            if (HasSameKeyboardLayout(GetKeyboardLayout(threadId), expectedLayout))
            {
                return;
            }

            Thread.Sleep(5);
        }
    }

    private static IntPtr ResolveKeyboardLayoutHandle(IntPtr keyboardLayout)
    {
        if (keyboardLayout == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var resolvedLayout = LoadKeyboardLayout(GetKeyboardLayoutIdentifier(keyboardLayout).ToString("X8"), KlfActivate | KlfSubstituteOk);

        if (resolvedLayout != IntPtr.Zero)
        {
            return resolvedLayout;
        }

        var languageId = GetKeyboardLanguageId(keyboardLayout);
        resolvedLayout = LoadKeyboardLayout($"0000{languageId:X4}", KlfActivate | KlfSubstituteOk);

        return resolvedLayout != IntPtr.Zero
            ? resolvedLayout
            : keyboardLayout;
    }

    private static void SendInput(IReadOnlyList<Input> inputs)
    {
        if (inputs.Count == 0)
        {
            return;
        }

        var inputSize = Marshal.SizeOf<Input>();
        var sent = SendInput((uint)inputs.Count, [.. inputs], inputSize);

        if (sent != inputs.Count)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                $"Failed to inject keyboard input. Sent={sent}; Expected={inputs.Count}; InputSize={inputSize}.");
        }
    }

    private static ControlStateSnapshot GetControlStateSnapshot()
    {
        return new ControlStateSnapshot(
            IsKeyPressed(VkLControl),
            IsKeyPressed(VkRControl));
    }

    private static bool IsHandledVKeyUp(int message, KbdllHookStruct keyInfo)
    {
        return (message == WmKeyUp || message == WmSysKeyUp) && IsVKey(keyInfo);
    }

    private static bool IsAltKeyUp(int message, KbdllHookStruct keyInfo)
    {
        if (message != WmKeyUp && message != WmSysKeyUp)
        {
            return false;
        }

        return keyInfo.vkCode == VkMenu ||
               keyInfo.vkCode == VkLMenu ||
               keyInfo.vkCode == VkRMenu;
    }

    private static bool IsControlKeyUp(int message, KbdllHookStruct keyInfo)
    {
        if (message != WmKeyUp && message != WmSysKeyUp)
        {
            return false;
        }

        return keyInfo.vkCode == VkControl ||
               keyInfo.vkCode == VkLControl ||
               keyInfo.vkCode == VkRControl;
    }

    private static bool IsVKeyDown(int message, KbdllHookStruct keyInfo)
    {
        return (message == WmKeyDown || message == WmSysKeyDown) && IsVKey(keyInfo);
    }

    private static bool IsVKey(KbdllHookStruct keyInfo)
    {
        return keyInfo.vkCode == VkV || keyInfo.scanCode == VScanCode;
    }

    private static bool IsKeyPressed(int virtualKey)
    {
        return (GetAsyncKeyState(virtualKey) & 0x8000) != 0;
    }

    private static bool HasSameKeyboardLayout(IntPtr left, IntPtr right)
    {
        if (left == IntPtr.Zero || right == IntPtr.Zero)
        {
            return left == right;
        }

        return GetKeyboardLayoutIdentifier(left) == GetKeyboardLayoutIdentifier(right);
    }

    private static bool HasSameKeyboardLanguage(IntPtr left, IntPtr right)
    {
        if (left == IntPtr.Zero || right == IntPtr.Zero)
        {
            return left == right;
        }

        return GetKeyboardLanguageId(left) == GetKeyboardLanguageId(right);
    }

    private static uint GetKeyboardLayoutIdentifier(IntPtr keyboardLayout)
    {
        return unchecked((uint)keyboardLayout.ToInt64());
    }

    private static ushort GetKeyboardLanguageId(IntPtr keyboardLayout)
    {
        return unchecked((ushort)(GetKeyboardLayoutIdentifier(keyboardLayout) & 0xFFFF));
    }

    private static bool IsEnglishLayout(IntPtr keyboardLayout)
    {
        return GetKeyboardLanguageId(keyboardLayout) == 0x0409;
    }

    private static string FormatKeyboardLayout(IntPtr keyboardLayout)
    {
        if (keyboardLayout == IntPtr.Zero)
        {
            return "0x00000000/lang=0000";
        }

        return $"0x{GetKeyboardLayoutIdentifier(keyboardLayout):X8}/lang={GetKeyboardLanguageId(keyboardLayout):X4}";
    }

    private readonly record struct CodexWindowContext
    {
        public required IntPtr WindowHandle { get; init; }

        public required uint ThreadId { get; init; }

        public required IntPtr KeyboardLayout { get; init; }

        public required string ProcessName { get; init; }

        public required string WindowTitle { get; init; }

        public required string WindowClass { get; init; }
    }

    private readonly record struct ControlStateSnapshot(bool LeftPressed, bool RightPressed)
    {
        public bool AnyPressed => LeftPressed || RightPressed;
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct KbdllHookStruct
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public uint type;
        public InputUnion union;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInput mi;

        [FieldOffset(0)]
        public KeybdInput ki;

        [FieldOffset(0)]
        public HardwareInput hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeybdInput
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HardwareInput
    {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(
        int idHook,
        LowLevelKeyboardProc lpfn,
        IntPtr hMod,
        uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint idThread);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint flags);

    [DllImport("user32.dll")]
    private static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, uint flags);

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        int msg,
        IntPtr wParam,
        IntPtr lParam,
        uint fuFlags,
        uint uTimeout,
        out IntPtr lpdwResult);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll")]
    private static extern bool IsClipboardFormatAvailable(uint format);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll")]
    private static extern IntPtr GetClipboardData(uint uFormat);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint DragQueryFile(IntPtr hDrop, uint iFile, StringBuilder? lpszFile, uint cch);
}

