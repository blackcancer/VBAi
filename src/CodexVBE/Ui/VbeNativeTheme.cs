using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace CodexVBE
{
    /// <summary>Applies an opt-in native dark-mode experiment to the VBE window tree.</summary>
    internal static class VbeNativeTheme
    {
        internal const string ExperimentVariable = "CODEXVBE_NATIVE_DARK_EXPERIMENT";
        internal const string LocalRefreshExperimentVariable = "CODEXVBE_NATIVE_LOCAL_REFRESH_EXPERIMENT";
        private const uint WmThemeChanged = 0x031A;
        private const uint WmRefreshChrome = 0x8000 + 0x564;
        private const uint WmRefreshCaption = 0x8000 + 0x565;
        // Read-only counters for the disposable native-renderer probe. Disabled
        // outside its process environment gate; no pointer crosses processes.
        private const uint WmQueryPropertyTabPaint = 0x8000 + 0x56A;
        private const uint WmQueryToolbarPaint = 0x8000 + 0x56B;
        private const uint WmEraseBackground = 0x0014;
        private const uint WmNcDestroy = 0x0082;
        private const uint WmCtlColorEdit = 0x0133;
        private const uint WmCtlColorListBox = 0x0134;
        private const uint WmCtlColorButton = 0x0135;
        private const uint WmCtlColorDialog = 0x0136;
        private const uint WmCtlColorScrollBar = 0x0137;
        private const uint WmCtlColorStatic = 0x0138;
        private const uint TreeSetBackground = 0x111D;
        private const uint TreeSetText = 0x111E;
        private const uint ListSetBackground = 0x1001;
        private const uint ListSetText = 0x1024;
        private const uint ListSetTextBackground = 0x1026;
        private const uint RichEditSetBackground = 0x0443;
        private const uint RedrawFlags = 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400;
        private const uint EventObjectCreate = 0x8000;
        private const uint EventObjectDestroy = 0x8001;
        private const uint EventObjectShow = 0x8002;
        private const uint WinEventOutOfContext = 0x0000;
        private const int ObjectIdWindow = 0;
        private const uint GetWindowOwner = 4;
        private const uint GetAncestorRoot = 2;
        private const int Background = 0x002B2420;
        private const int Foreground = 0x00F0E8E2;

        internal delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate int SetPreferredAppModeDelegate(int mode);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate bool AllowDarkModeForWindowDelegate(IntPtr window, bool allow);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate void VoidThemeDelegate();
        internal delegate IntPtr SubclassCallback(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, UIntPtr subclassId, IntPtr referenceData);
        internal delegate void WinEventCallback(IntPtr hook, uint eventType, IntPtr window, int objectId, int childId, uint eventThread, uint eventTime);

        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern int SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowStyle(IntPtr window, int index);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out NativeRect rectangle);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string caption);
        [DllImport("user32.dll")] private static extern int FillRect(IntPtr deviceContext, ref NativeRect rectangle, IntPtr brush);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
        [DllImport("gdi32.dll")] private static extern int SaveDC(IntPtr deviceContext);
        [DllImport("gdi32.dll")] private static extern bool RestoreDC(IntPtr deviceContext, int saved);
        [DllImport("gdi32.dll")] private static extern int SetTextColor(IntPtr deviceContext, int color);
        [DllImport("gdi32.dll")] private static extern int SetBkColor(IntPtr deviceContext, int color);
        [DllImport("user32.dll")] private static extern bool RedrawWindow(IntPtr window, IntPtr update, IntPtr region, uint flags);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr module, WinEventCallback callback, uint processId, uint threadId, uint flags);
        [DllImport("user32.dll")] private static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("comctl32.dll")] private static extern bool SetWindowSubclass(IntPtr window, SubclassCallback callback, UIntPtr subclassId, IntPtr referenceData);
        [DllImport("comctl32.dll")] private static extern bool RemoveWindowSubclass(IntPtr window, SubclassCallback callback, UIntPtr subclassId);
        [DllImport("comctl32.dll")] private static extern IntPtr DefSubclassProc(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] private static extern int SetWindowTheme(IntPtr window, string appName, string idList);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string moduleName);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)] private static extern IntPtr GetProcAddress(IntPtr module, IntPtr ordinal);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NativeOsVersion
        {
            internal uint Size, Major, Minor, Build, Platform;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] internal string ServicePack;
        }
        [DllImport("ntdll.dll", ExactSpelling = true)] private static extern int RtlGetVersion(ref NativeOsVersion version);

        internal static Version ReadNativeWindowsVersion()
        {
            var version = new NativeOsVersion { Size = (uint)Marshal.SizeOf(typeof(NativeOsVersion)), ServicePack = string.Empty };
            if (RtlGetVersion(ref version) != 0)
                throw new PlatformNotSupportedException("Cannot verify the Windows dark-mode ABI.");
            return new Version((int)version.Major, (int)version.Minor, (int)version.Build);
        }

        internal static bool SupportsPreferredAppMode(Version version)
        {
            // Ordinal 135 used a different signature in Windows 10 1809.
            // Future major versions must be qualified before calling private APIs.
            return version != null && version.Major == 10 && version.Minor == 0 && version.Build >= 18362;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct NativeRect { internal int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeDrawItem
        {
            internal uint ControlType, ControlId, ItemId, Action, State;
            internal IntPtr Window, DeviceContext;
            internal NativeRect Bounds;
            internal UIntPtr Data;
        }

        private static readonly UIntPtr SubclassId = new UIntPtr(0x56424544);
        private static readonly SubclassCallback Subclass = ThemeWindowProcedure;
        private static readonly WinEventCallback WindowEvent = WindowEventReceived;
        private static readonly object Sync = new object();
        private static readonly List<IntPtr> themedWindows = new List<IntPtr>();
        private static readonly List<IntPtr> subclassedWindows = new List<IntPtr>();
        private static readonly HashSet<IntPtr> pendingChrome = new HashSet<IntPtr>();
        private static readonly HashSet<IntPtr> pendingCaptions = new HashSet<IntPtr>();
        private static int toolbarPaintCount, toolbarDeferredPaintCount;
        private static bool localChromeRefresh;
        private static readonly Dictionary<IntPtr, VbeNativePropertyTabs> propertyTabs = new Dictionary<IntPtr, VbeNativePropertyTabs>();
        private sealed class ControlPalette
        {
            internal int Background, Foreground, TextBackground;
        }
        private static readonly Dictionary<IntPtr, ControlPalette> originalControlColors = new Dictionary<IntPtr, ControlPalette>();
        private static AllowDarkModeForWindowDelegate allowDarkMode;
        private static SetPreferredAppModeDelegate setPreferredMode;
        private static VoidThemeDelegate flushMenuThemes;
        private static IntPtr backgroundBrush;
        private static IntPtr editorWindow;
        private static int previousPreferredMode;
        private static IntPtr windowEventHook;
        private static bool preferredModeChanged;
        private static VbeNativePalette nativePalette;
        private static IntPtr immediateWindow;
        private static string immediateCaption;

        /// <summary>Returns whether the explicitly gated experiment is enabled for this host process.</summary>
        internal static bool ExperimentEnabled()
        {
            return string.Equals(Environment.GetEnvironmentVariable(ExperimentVariable), "1", StringComparison.Ordinal);
        }

        internal static Func<IntPtr, int> ApplyNativeTheme = Apply;
        internal static Func<object, IntPtr, VbeNativePalette> CreatePalette = (vbe, editor) => new VbeNativePalette(vbe, editor);

        /// <summary>Stores the VBE owner and applies the persisted or explicitly gated preference.</summary>
        internal static void Initialize(IntPtr editor, bool enabled, object vbe = null)
        {
            if (editor == IntPtr.Zero) throw new ArgumentException("The VBE main window handle is required.", nameof(editor));
            editorWindow = editor;
            immediateWindow = IntPtr.Zero;
            immediateCaption = null;
            if (vbe != null)
            {
                try
                {
                    foreach (dynamic pane in ((dynamic)vbe).Windows)
                        if ((int)pane.Type == 5)
                        {
                            immediateWindow = new IntPtr(Convert.ToInt64(pane.HWnd));
                            immediateCaption = Convert.ToString(pane.Caption);
                            break;
                        }
                }
                catch (Exception error) { LoadLog.Write("Native Immediate pane identification failed: " + error.Message); }
            }
            nativePalette?.Dispose();
            // Disposable probes drive palette recovery explicitly in their own
            // artifact directory; they must not create a production recovery file.
            nativePalette = vbe != null && !ExperimentEnabled() ? CreatePalette(vbe, editor) : null;
            SetEnabled(enabled || ExperimentEnabled());
        }

        /// <summary>Applies or removes native dark styling in the current VBE process.</summary>
        internal static void SetEnabled(bool enabled)
        {
            if (!enabled)
            {
                if (!Reset()) throw new InvalidOperationException("Native VBE renderer could not stop; deactivate it on the VBE UI thread.");
                nativePalette?.Request(false);
                return;
            }
            if (editorWindow == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not initialized.");
            if (themedWindows.Count == 0)
            {
                try { ApplyNativeTheme(editorWindow); }
                catch { Reset(); throw; }
            }
            nativePalette?.Request(true);
        }

        /// <summary>Applies the native dark-mode experiment to the VBE main window and existing native children.</summary>
        /// <param name="editor">VBE main window handle.</param>
        /// <returns>Number of native windows submitted to the dark-mode APIs.</returns>
        internal static int Apply(IntPtr editor)
        {
            if (editor == IntPtr.Zero) throw new ArgumentException("The VBE main window handle is required.", nameof(editor));
            if (!IsCurrentWindowThread(editor))
                throw new InvalidOperationException("The native theme must be initialized on the VBE window thread.");
            localChromeRefresh = string.Equals(Environment.GetEnvironmentVariable(LocalRefreshExperimentVariable), "1", StringComparison.Ordinal);
            Version windowsVersion = ReadNativeWindowsVersion();
            if (!SupportsPreferredAppMode(windowsVersion))
                throw new PlatformNotSupportedException("Native VBE dark mode requires the Windows 10 1903 or newer dark-mode ABI. Detected: " + windowsVersion);
            IntPtr themeModule = GetModuleHandle("uxtheme.dll");
            setPreferredMode = Resolve<SetPreferredAppModeDelegate>(themeModule, 135);
            allowDarkMode = Resolve<AllowDarkModeForWindowDelegate>(themeModule, 133);
            var refreshPolicy = Resolve<VoidThemeDelegate>(themeModule, 104);
            flushMenuThemes = Resolve<VoidThemeDelegate>(themeModule, 136);
            if (setPreferredMode == null || allowDarkMode == null)
                throw new PlatformNotSupportedException("The Windows native dark-mode entry points are unavailable.");

            // ForceDark is process-wide and deliberately restricted to the explicit experiment gate.
            if (ExperimentEnabled())
            {
                previousPreferredMode = setPreferredMode(2);
                preferredModeChanged = true;
                refreshPolicy?.Invoke();
            }
            VbeNativeRenderer.Start(editor);
            var windows = new List<IntPtr> { editor };
            EnumChildWindows(editor, (window, parameter) => { windows.Add(window); return true; }, IntPtr.Zero);
            foreach (IntPtr window in windows) ApplyWindow(window);
            int enabled = 1;
            int titleResult = DwmSetWindowAttribute(editor, 20, ref enabled, sizeof(int));
            if (titleResult != 0) DwmSetWindowAttribute(editor, 19, ref enabled, sizeof(int));
            flushMenuThemes?.Invoke();
            editorWindow = editor;
            windowEventHook = SetWinEventHook(EventObjectCreate, EventObjectShow, IntPtr.Zero, WindowEvent,
                (uint)System.Diagnostics.Process.GetCurrentProcess().Id, 0, WinEventOutOfContext);
            RedrawWindow(editor, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
            return themedWindows.Count;
        }

        /// <summary>Removes the experimental subclasses and restores the host process theme preference.</summary>
        internal static bool Reset()
        {
            lock (Sync)
            {
                if (!VbeNativeRenderer.Stop()) return false;
                if (windowEventHook != IntPtr.Zero) { UnhookWinEvent(windowEventHook); windowEventHook = IntPtr.Zero; }
                foreach (VbeNativePropertyTabs tabs in propertyTabs.Values) tabs.Dispose();
                propertyTabs.Clear();
                foreach (IntPtr window in subclassedWindows) RemoveWindowSubclass(window, Subclass, SubclassId);
                subclassedWindows.Clear();
                pendingChrome.Clear();
                pendingCaptions.Clear();
                toolbarPaintCount = toolbarDeferredPaintCount = 0;
                foreach (IntPtr window in themedWindows.ToArray())
                {
                    allowDarkMode?.Invoke(window, false);
                    if (WindowClass(window) == "#32770")
                    {
                        int disabled = 0;
                        DwmSetWindowAttribute(window, 20, ref disabled, sizeof(int));
                        DwmSetWindowAttribute(window, 19, ref disabled, sizeof(int));
                    }
                    SetWindowTheme(window, null, null);
                    RestoreControlPalette(window, WindowClass(window));
                    SendMessage(window, WmThemeChanged, IntPtr.Zero, IntPtr.Zero);
                }
                themedWindows.Clear();
                originalControlColors.Clear();
                if (backgroundBrush != IntPtr.Zero) { DeleteObject(backgroundBrush); backgroundBrush = IntPtr.Zero; }
            }
            if (preferredModeChanged)
            {
                setPreferredMode?.Invoke(previousPreferredMode);
                preferredModeChanged = false;
            }
            flushMenuThemes?.Invoke();
            if (editorWindow != IntPtr.Zero)
            {
                int disabled = 0;
                DwmSetWindowAttribute(editorWindow, 20, ref disabled, sizeof(int));
                DwmSetWindowAttribute(editorWindow, 19, ref disabled, sizeof(int));
                RedrawWindow(editorWindow, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
            }
            return true;
        }

        /// <summary>Restores native styling and releases the VBE handle when the add-in disconnects.</summary>
        internal static bool Disconnect()
        {
            if (!Reset()) return false;
            nativePalette?.Dispose();
            nativePalette = null;
            editorWindow = IntPtr.Zero;
            immediateWindow = IntPtr.Zero;
            immediateCaption = null;
            return true;
        }

        internal static void ApplyControlPalette(IntPtr window, string className)
        {
            if (originalControlColors.ContainsKey(window)) return;
            IntPtr background = new IntPtr(Background);
            var original = new ControlPalette();
            switch (className)
            {
                case "SysTreeView32":
                    original.Background = SendMessage(window, TreeSetBackground, IntPtr.Zero, background);
                    original.Foreground = SendMessage(window, TreeSetText, IntPtr.Zero, new IntPtr(Foreground));
                    break;
                case "SysListView32":
                    original.Background = SendMessage(window, 0x1000, IntPtr.Zero, IntPtr.Zero);
                    original.Foreground = SendMessage(window, 0x1023, IntPtr.Zero, IntPtr.Zero);
                    original.TextBackground = SendMessage(window, 0x1025, IntPtr.Zero, IntPtr.Zero);
                    SendMessage(window, ListSetBackground, IntPtr.Zero, background);
                    SendMessage(window, ListSetText, IntPtr.Zero, new IntPtr(Foreground));
                    SendMessage(window, ListSetTextBackground, IntPtr.Zero, background);
                    break;
                case "RichEdit20A":
                case "RICHEDIT60W":
                    original.Background = SendMessage(window, RichEditSetBackground, IntPtr.Zero, background);
                    break;
                default: return;
            }
            originalControlColors.Add(window, original);
        }

        internal static void RestoreControlPalette(IntPtr window, string className)
        {
            ControlPalette original;
            if (!originalControlColors.TryGetValue(window, out original)) return;
            IntPtr background = new IntPtr(original.Background);
            IntPtr foreground = new IntPtr(original.Foreground);
            switch (className)
            {
                case "SysTreeView32":
                    SendMessage(window, TreeSetBackground, IntPtr.Zero, background);
                    SendMessage(window, TreeSetText, IntPtr.Zero, foreground);
                    break;
                case "SysListView32":
                    SendMessage(window, ListSetBackground, IntPtr.Zero, background);
                    SendMessage(window, ListSetText, IntPtr.Zero, foreground);
                    SendMessage(window, ListSetTextBackground, IntPtr.Zero, new IntPtr(original.TextBackground));
                    break;
                case "RichEdit20A":
                case "RICHEDIT60W":
                    SendMessage(window, RichEditSetBackground, IntPtr.Zero, background);
                    break;
            }
            originalControlColors.Remove(window);
        }

        private static void ApplyWindow(IntPtr window)
        {
            lock (Sync)
            {
                if (window == IntPtr.Zero || themedWindows.Contains(window)) return;
                // SetWindowSubclass is only supported on the owning thread. An
                // out-of-context WinEvent may also describe another host thread.
                // Do not mark those windows as themed or prevent a later retry.
                if (!IsCurrentWindowThread(window)) return;
                string className = WindowClass(window);
                if (className == "MsoCommandBar") VbeNativeRenderer.Register(window);
                if (className == "VbaWindow" && immediateWindow == IntPtr.Zero && !string.IsNullOrEmpty(immediateCaption))
                {
                    var caption = new StringBuilder(256);
                    GetWindowText(window, caption, caption.Capacity);
                    if (caption.ToString() == immediateCaption) immediateWindow = window;
                }
                if (string.IsNullOrEmpty(className) || (IsManagedAddInWindow(className) && className != "GenericPane")) return;
                bool propertiesControl = WindowClass(GetAncestorParent(window)) == "wndclass_pbrs";
                bool dialogControl = WindowClass(GetAncestor(window, GetAncestorRoot)) == "#32770";
                bool scopedControl = className == "ListBox" ? propertiesControl :
                    className != "SysTabControl32" || propertiesControl || dialogControl;
                if (scopedControl && ShouldSubclass(className))
                {
                    if (!SetWindowSubclass(window, Subclass, SubclassId, IntPtr.Zero)) return;
                    subclassedWindows.Add(window);
                }
                if (className == "SysTabControl32" && propertiesControl)
                    TryAttachPropertyTabs(window);
                allowDarkMode(window, true);
                if (className == "#32770")
                {
                    int enabled = 1;
                    if (DwmSetWindowAttribute(window, 20, ref enabled, sizeof(int)) != 0)
                        DwmSetWindowAttribute(window, 19, ref enabled, sizeof(int));
                }
                SetWindowTheme(window, "DarkMode_Explorer", null);
                // Keep the standard tab interaction and geometry. Supported
                // Properties tabs supply their final colors before drawing text;
                // other tabs retain the existing classic-renderer fallback.
                // The themed checkbox/group-box renderer ignores the dialog text
                // color. Classic controls honor WM_CTLCOLOR from their parent.
                if ((className == "SysTabControl32" && (propertiesControl || dialogControl)) ||
                    (className == "Button" && dialogControl && IsDialogLabelButton(window)))
                    SetWindowTheme(window, "", "");
                ApplyControlPalette(window, className);
                if (className == "GenericPane")
                    LoadLog.Write("Native theme GenericPane: " + window + ", subclass=" + subclassedWindows.Contains(window));
                themedWindows.Add(window);
                SendMessage(window, WmThemeChanged, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private static void WindowEventReceived(IntPtr hook, uint eventType, IntPtr window, int objectId, int childId, uint eventThread, uint eventTime)
        {
            if (objectId != ObjectIdWindow || childId != 0) return;
            // Destruction can arrive after the parent relationship has disappeared.
            // Forget the handle so a later window reusing it receives its own theme.
            if (eventType == EventObjectDestroy)
            {
                ForgetWindow(window);
                return;
            }
            if (!BelongsToEditor(window)) return;
            if (IsCurrentWindowThread(window)) VbeNativeRenderer.Refresh();
            ApplyWindow(window);
            if (eventType == EventObjectShow)
            {
                // Children may be created before their popup or pane is attached
                // to the editor. Revisit them once the owner becomes visible.
                EnumChildWindows(window, (child, parameter) =>
                {
                    ApplyWindow(child);
                    return true;
                }, IntPtr.Zero);
            }
            RedrawWindow(window, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
        }

        private static bool BelongsToEditor(IntPtr window)
        {
            if (editorWindow == IntPtr.Zero || window == IntPtr.Zero) return false;
            if (window == editorWindow || IsChild(editorWindow, window)) return true;
            IntPtr root = GetAncestor(window, GetAncestorRoot);
            IntPtr owner = GetWindow(root == IntPtr.Zero ? window : root, GetWindowOwner);
            for (int depth = 0; owner != IntPtr.Zero && depth < 16; depth++)
            {
                if (owner == editorWindow || IsChild(editorWindow, owner)) return true;
                owner = GetWindow(owner, GetWindowOwner);
            }
            return false;
        }

        internal static bool ShouldSubclass(string className)
        {
            return className == "#32770" || className == "wndclass_desked_gsk" || className == "MDIClient" ||
                className == "VbaWindow" || className == "PROJECT" || className == "wndclass_pbrs" ||
                className == "ToolsPalette" || className == "VBSlider" || className == "GenericPane" ||
                className == "MsoCommandBar" || className == "MsoCommandBarPopup" || className == "MsoCommandBarDock" || className == "SysTabControl32" ||
                className == "ListBox" || className == "SysTreeView32" || className == "ComboBox" || className == "Edit";
        }

        private static void ForgetWindow(IntPtr window)
        {
            lock (Sync)
            {
                themedWindows.Remove(window);
                subclassedWindows.Remove(window);
                pendingChrome.Remove(window);
                pendingCaptions.Remove(window);
                originalControlColors.Remove(window);
                VbeNativePropertyTabs tabs;
                if (propertyTabs.TryGetValue(window, out tabs))
                {
                    tabs.Dispose();
                    propertyTabs.Remove(window);
                }
                if (window == immediateWindow) immediateWindow = IntPtr.Zero;
            }
        }

        private static IntPtr ThemeWindowProcedure(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, UIntPtr subclassId, IntPtr referenceData)
        {
            if (message == WmQueryToolbarPaint && ExperimentEnabled())
                return new IntPtr(wParam.ToInt64() == 1 ? toolbarDeferredPaintCount : toolbarPaintCount);
            if (message == WmRefreshCaption)
            {
                if (!pendingCaptions.Remove(window)) return IntPtr.Zero;
                if (subclassedWindows.Contains(window))
                {
                    VbeNativeChrome.Paint(window, false, IntPtr.Zero, WindowClass(window) == "GenericPane");
                    VbeNativeChrome.PaintBorder(window);
                }
                return IntPtr.Zero;
            }
            if (message == WmRefreshChrome)
            {
                if (!pendingChrome.Remove(window)) return IntPtr.Zero;
                if (!subclassedWindows.Contains(window)) return IntPtr.Zero;
                if (propertyTabs.ContainsKey(window)) return IntPtr.Zero;
                string refreshedClass = WindowClass(window);
                if (!localChromeRefresh && (refreshedClass == "PROJECT" || refreshedClass == "wndclass_pbrs" || refreshedClass == "VbaWindow"))
                    VbeNativeChrome.Paint(window, false, IntPtr.Zero);
                if (refreshedClass == "MsoCommandBar") toolbarDeferredPaintCount++;
                PaintChrome(window, refreshedClass, IntPtr.Zero);
                return IntPtr.Zero;
            }
            if (message == WmNcDestroy)
            {
                RemoveWindowSubclass(window, Subclass, SubclassId);
                ForgetWindow(window);
                return DefSubclassProc(window, message, wParam, lParam);
            }
            VbeNativePropertyTabs directTabs;
            if (propertyTabs.TryGetValue(window, out directTabs))
            {
                if (message == WmQueryPropertyTabPaint && ExperimentEnabled())
                    return new IntPtr(wParam.ToInt64() == 2 ? (VbeNativePropertyTabs.CanRender(window) ? 1 : -1) :
                        wParam.ToInt64() == 1 ? directTabs.PrintCount : directTabs.PaintCount);
                IntPtr handled;
                if (directTabs.TryHandleMessage(message, wParam, lParam, out handled)) return handled;
                IntPtr nativeResult = DefSubclassProc(window, message, wParam, lParam);
                directTabs.AfterNativeMessage(message, wParam, lParam);
                return nativeResult;
            }
            string className = message == WmEraseBackground ? WindowClass(window) : null;
            if (message == WmEraseBackground && (className == "#32770" || className == "MDIClient" || className == "VBSlider"))
            {
                NativeRect rectangle;
                if (GetClientRect(window, out rectangle))
                {
                    EnsureBackgroundBrush();
                    FillRect(wParam, ref rectangle, backgroundBrush);
                    return new IntPtr(1);
                }
            }
            if (message == WmCtlColorEdit || message == WmCtlColorListBox || message == WmCtlColorButton ||
                message == WmCtlColorDialog || message == WmCtlColorScrollBar || message == WmCtlColorStatic)
            {
                EnsureBackgroundBrush();
                SetTextColor(wParam, Foreground);
                SetBkColor(wParam, Background);
                return backgroundBrush;
            }
            if (message == 0x002B && lParam != IntPtr.Zero && WindowClass(window) == "wndclass_pbrs")
            {
                var item = (NativeDrawItem)Marshal.PtrToStructure(lParam, typeof(NativeDrawItem));
                if (item.ControlType == 2 && WindowClass(item.Window) == "ListBox" &&
                    item.DeviceContext != IntPtr.Zero && (item.Action & 3) != 0)
                {
                    int saved = SaveDC(item.DeviceContext);
                    if (saved != 0)
                    {
                        try
                        {
                            // Render the native row on its original light palette,
                            // then translate that fresh row exactly once. Changing
                            // the whole list later would retouch ClearType pixels.
                            SetTextColor(item.DeviceContext, 0);
                            SetBkColor(item.DeviceContext, 0x00ffffff);
                            IntPtr drawn = DefSubclassProc(window, message, wParam, lParam);
                            VbeNativeChrome.PaintPropertyRow(item.DeviceContext, item.Bounds);
                            return drawn;
                        }
                        finally { RestoreDC(item.DeviceContext, saved); }
                    }
                }
            }
            IntPtr result = DefSubclassProc(window, message, wParam, lParam);
            // Creation notifications can precede insertion of the native items.
            // Retry after native initialization instead of fixing the fallback
            // renderer for the complete lifetime of that HWND.
            if ((message == 0x0018 || message == 0x0047 || message == 0x1307 || message == 0x133e) &&
                WindowClass(window) == "SysTabControl32" && WindowClass(GetAncestorParent(window)) == "wndclass_pbrs" &&
                TryAttachPropertyTabs(window))
            {
                propertyTabs[window].AfterNativeMessage(0x031a, IntPtr.Zero, IntPtr.Zero);
                return result;
            }
            if ((message == 0x0201 || message == 0x0202 || message == 0x014F || message == 0x000A) && WindowClass(window) == "ComboBox")
                VbeNativeChrome.PaintComboButton(window);
            if (message == 0x0085 || message == 0x0086 || message == 0x000C || message == 0x0317)
            {
                if (window != editorWindow) VbeNativeChrome.Paint(window, false, message == 0x0317 ? wParam : IntPtr.Zero);
            }
            if (message == 0x000F || message == 0x0317 || message == 0x0318 || message == 0x0200 || message == 0x02A3)
            {
                string paintClass = WindowClass(window);
                // Moving over a document without selecting text does not repaint
                // its content; avoid copying its entire surface for every mouse move.
                if (paintClass != "VbaWindow" || message != 0x0200 || (wParam.ToInt64() & 1) != 0)
                    PaintChrome(window, paintClass, message == 0x0317 || message == 0x0318 || message == 0x000F ? wParam : IntPtr.Zero);
            }
            // The legacy editor can draw text or scroll pixels directly without a
            // subsequent WM_PAINT. Coalesce one refresh after those operations.
            if ((message == 0x0102 || message == 0x0101 || message == 0x0202 || message == 0x0114 || message == 0x0115 || message == 0x020A) &&
                WindowClass(window) == "VbaWindow" && pendingChrome.Add(window))
            {
                if (!PostMessage(window, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero)) pendingChrome.Remove(window);
            }
            // Caption work is local and never triggers a code-surface conversion.
            // A Properties notification must not repaint unrelated Office toolbars.
            if (localChromeRefresh && (message == 0x0086 || message == 0x000c || message == 0x0047 || message == 0x0222))
            {
                string captionClass = WindowClass(window);
                if ((captionClass == "PROJECT" || captionClass == "wndclass_pbrs" || captionClass == "VbaWindow" || captionClass == "GenericPane") &&
                    pendingCaptions.Add(window) && !PostMessage(window, WmRefreshCaption, IntPtr.Zero, IntPtr.Zero))
                    pendingCaptions.Remove(window);
            }
            // The measured VBE updates its Standard toolbar directly during code
            // activation and caret actions (outside that toolbar's WM_PAINT).
            // Restrict the legacy catch-up to those operations and the editor's
            // own toolbar dock, rather than every paint/notification in the IDE.
            if (localChromeRefresh && (message == 0x0222 || message == 0x0102 || message == 0x0101 || message == 0x0202 ||
                message == 0x0114 || message == 0x0115 || message == 0x020a) && IsCodeSurface(window, WindowClass(window)))
                QueueContainerChromeRefresh(editorWindow);
            if (localChromeRefresh && (message == 0x000f || message == 0x0085 || message == 0x0222))
            {
                string updatedClass = WindowClass(window);
                // Office updates shared toolbar state and hosted captions while
                // painting/activating a document, including its form designer.
                // Properties/list/tab paints deliberately have no such fan-out.
                if (window == editorWindow || updatedClass == "MDIClient" || updatedClass == "VbaWindow")
                    QueueContainerChromeRefresh(editorWindow);
                else if (updatedClass == "PROJECT")
                    QueueContainerChromeRefresh(window);
            }
            if (localChromeRefresh && message == 0x004e && WindowClass(window) == "PROJECT")
                QueueContainerChromeRefresh(window);
            // The scoped-refresh pilot leaves some Office button faces light after
            // view activation. Keep its old recovery path until a renderer supplies
            // final colors before drawing. The pilot is a diagnostic-only opt-in.
            if (!localChromeRefresh && (message == 0x000f || message == 0x0085 || message == 0x004e || message == 0x0047))
            {
                foreach (IntPtr candidate in subclassedWindows.ToArray())
                {
                    if (propertyTabs.ContainsKey(candidate)) continue;
                    string candidateClass = WindowClass(candidate);
                    if ((candidateClass == "MsoCommandBar" || candidateClass == "MsoCommandBarPopup" || candidateClass == "MsoCommandBarDock" ||
                         candidateClass == "GenericPane" || candidateClass == "SysTabControl32" || candidateClass == "ListBox" || candidateClass == "VbaWindow" ||
                         candidateClass == "PROJECT" || candidateClass == "wndclass_pbrs" || candidateClass == "SysTreeView32" || candidateClass == "ComboBox" || candidateClass == "Edit") && pendingChrome.Add(candidate))
                        if (!PostMessage(candidate, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero)) pendingChrome.Remove(candidate);
                }
            }
            return result;
        }

        private static void PaintChrome(IntPtr window, string className, IntPtr dc)
        {
            // A direct renderer must never be followed by image recoloring or a
            // screen-DC border pass, including already queued legacy refreshes.
            if (propertyTabs.ContainsKey(window)) return;
            if (className == "MsoCommandBar" && VbeNativeRenderer.Active &&
                IsChild(editorWindow, window)) return;
            if (className == "MsoCommandBar") toolbarPaintCount++;
            if (className == "ComboBox" && dc == IntPtr.Zero) VbeNativeChrome.PaintComboButton(window);
            if (className == "GenericPane") VbeNativeChrome.Paint(window, false, dc, true);
            bool codeSurface = IsCodeSurface(window, className);
            if (codeSurface) VbeNativeChrome.Paint(window, true, dc, false, false, true);
            if (className == "MsoCommandBar" || className == "MsoCommandBarPopup" || className == "MsoCommandBarDock" || className == "SysTabControl32" ||
                (className == "VbaWindow" && !codeSurface))
                VbeNativeChrome.Paint(window, true, dc, false, className == "VbaWindow");
            if (dc == IntPtr.Zero && window != editorWindow) VbeNativeChrome.PaintBorder(window);
        }

        private static bool IsDialogLabelButton(IntPtr window)
        {
            int buttonType = GetWindowStyle(window, -16) & 0x0f;
            // Checkbox, three-state, radio and group-box labels need the parent
            // text color. Push buttons keep their native dark visual style.
            return (buttonType >= 2 && buttonType <= 7) || buttonType == 9;
        }

        private static IntPtr GetAncestorParent(IntPtr window) { return GetAncestor(window, 1); }

        private static bool IsCodeSurface(IntPtr window, string className)
        {
            return className == "VbaWindow" && (window == immediateWindow ||
                FindWindowEx(window, IntPtr.Zero, "ObtbarWndClass", null) != IntPtr.Zero);
        }

        private static void QueueContainerChromeRefresh(IntPtr container)
        {
            foreach (IntPtr candidate in subclassedWindows.ToArray())
            {
                IntPtr dock = GetAncestorParent(candidate);
                if (WindowClass(candidate) == "MsoCommandBar" && WindowClass(dock) == "MsoCommandBarDock" &&
                    GetAncestorParent(dock) == container && pendingChrome.Add(candidate) &&
                    !PostMessage(candidate, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero))
                    pendingChrome.Remove(candidate);
                if (dock == container && WindowClass(candidate) == "GenericPane" && pendingCaptions.Add(candidate) &&
                    !PostMessage(candidate, WmRefreshCaption, IntPtr.Zero, IntPtr.Zero))
                    pendingCaptions.Remove(candidate);
            }
        }

        private static bool TryAttachPropertyTabs(IntPtr window)
        {
            if (propertyTabs.ContainsKey(window)) return true;
            VbeNativePropertyTabs renderer;
            if (!VbeNativePropertyTabs.TryCreate(window, out renderer)) return false;
            propertyTabs.Add(window, renderer);
            pendingChrome.Remove(window);
            LoadLog.Write("Native Properties tabs: direct text renderer attached to " + window);
            return true;
        }

        private static bool IsCurrentWindowThread(IntPtr window)
        {
            uint processId;
            return GetWindowThreadProcessId(window, out processId) == GetCurrentThreadId() &&
                processId == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        private static void EnsureBackgroundBrush()
        {
            if (backgroundBrush == IntPtr.Zero) backgroundBrush = CreateSolidBrush(Background);
        }

        private static T Resolve<T>(IntPtr module, int ordinal) where T : class
        {
            if (module == IntPtr.Zero) return null;
            IntPtr address = GetProcAddress(module, new IntPtr(ordinal));
            return address == IntPtr.Zero ? null : Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
        }

        private static string WindowClass(IntPtr window)
        {
            var text = new StringBuilder(256);
            GetClassName(window, text, text.Capacity);
            return text.ToString();
        }

        internal static bool IsManagedAddInWindow(string className)
        {
            return className.StartsWith("WindowsForms10.", StringComparison.Ordinal) ||
                className.StartsWith("HwndWrapper[", StringComparison.Ordinal) ||
                className == "GenericPane";
        }
    }
}
