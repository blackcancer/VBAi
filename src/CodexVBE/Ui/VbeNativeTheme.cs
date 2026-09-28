using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace CodexVBE
{
    /// <summary>Applies an opt-in native dark-mode experiment to the VBE window tree.</summary>
    internal static class VbeNativeTheme
    {
        /// <summary>Process environment variable that explicitly enables the native dark-mode experiment.</summary>
        internal const string ExperimentVariable = "CODEXVBE_NATIVE_DARK_EXPERIMENT";
        /// <summary>Process environment variable that enables the experimental local chrome-refresh path.</summary>
        internal const string LocalRefreshExperimentVariable = "CODEXVBE_NATIVE_LOCAL_REFRESH_EXPERIMENT";
        /// <summary>Native message sent when a window's theme changes.</summary>
        private const uint WmThemeChanged = 0x031A;
        /// <summary>Private message requesting a deferred chrome repaint.</summary>
        private const uint WmRefreshChrome = 0x8000 + 0x564;
        /// <summary>Private message requesting a deferred caption repaint.</summary>
        private const uint WmRefreshCaption = 0x8000 + 0x565;
        // Read-only counters for the disposable native-renderer probe. Disabled
        // outside its process environment gate; no pointer crosses processes.
        /// <summary>Probe-only message that reads the direct property-tab paint counter.</summary>
        private const uint WmQueryPropertyTabPaint = 0x8000 + 0x56A;
        /// <summary>Probe-only message that reads the native toolbar paint counter.</summary>
        private const uint WmQueryToolbarPaint = 0x8000 + 0x56B;
        /// <summary>Windows message sent before a control erases its background.</summary>
        private const uint WmEraseBackground = 0x0014;
        /// <summary>Windows message sent when a window is destroyed.</summary>
        private const uint WmNcDestroy = 0x0082;
        /// <summary>Windows color request for edit controls.</summary>
        private const uint WmCtlColorEdit = 0x0133;
        /// <summary>Windows color request for list boxes.</summary>
        private const uint WmCtlColorListBox = 0x0134;
        /// <summary>Windows color request for buttons.</summary>
        private const uint WmCtlColorButton = 0x0135;
        /// <summary>Windows color request for dialog backgrounds.</summary>
        private const uint WmCtlColorDialog = 0x0136;
        /// <summary>Windows color request for scroll bars.</summary>
        private const uint WmCtlColorScrollBar = 0x0137;
        /// <summary>Windows color request for static controls.</summary>
        private const uint WmCtlColorStatic = 0x0138;
        /// <summary>Tree-view message that sets its background color.</summary>
        private const uint TreeSetBackground = 0x111D;
        /// <summary>Tree-view message that sets its text color.</summary>
        private const uint TreeSetText = 0x111E;
        /// <summary>List-view message that sets its background color.</summary>
        private const uint ListSetBackground = 0x1001;
        /// <summary>List-view message that sets its text color.</summary>
        private const uint ListSetText = 0x1024;
        /// <summary>List-view message that sets its text background color.</summary>
        private const uint ListSetTextBackground = 0x1026;
        /// <summary>Rich Edit message that sets the editor background color.</summary>
        private const uint RichEditSetBackground = 0x0443;
        /// <summary>RedrawWindow flags used to invalidate and update themed editor windows.</summary>
        private const uint RedrawFlags = 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400;
        /// <summary>WinEvent identifier for a newly created accessible object.</summary>
        private const uint EventObjectCreate = 0x8000;
        /// <summary>WinEvent identifier for a destroyed accessible object.</summary>
        private const uint EventObjectDestroy = 0x8001;
        /// <summary>WinEvent identifier for an accessible object becoming visible.</summary>
        private const uint EventObjectShow = 0x8002;
        /// <summary>WinEvent hook mode that delivers callbacks out of context.</summary>
        private const uint WinEventOutOfContext = 0x0000;
        /// <summary>Accessible object identifier representing the window itself.</summary>
        private const int ObjectIdWindow = 0;
        /// <summary>GetWindow selector for the owner window.</summary>
        private const uint GetWindowOwner = 4;
        /// <summary>GetAncestor selector for the root ancestor.</summary>
        private const uint GetAncestorRoot = 2;
        /// <summary>Dark background RGB color used for native controls.</summary>
        private const int Background = 0x002B2420;
        /// <summary>Light foreground RGB color used for native controls.</summary>
        private const int Foreground = 0x00F0E8E2;

        /// <summary>Callback used to enumerate native child windows.</summary>
        /// <param name="window">Child window handle.</param><param name="parameter">Caller-supplied context.</param>
        /// <returns><see langword="true"/> to continue enumeration.</returns>
        internal delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
        /// <summary>Private Windows theme API for setting the process preferred app mode.</summary>
        /// <param name="mode">Native preferred-mode value.</param><returns>Previous mode value.</returns>
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate int SetPreferredAppModeDelegate(int mode);
        /// <summary>Private Windows theme API for allowing dark mode on one window.</summary>
        /// <param name="window">Target window handle.</param><param name="allow">Whether dark mode is allowed.</param>
        /// <returns>Whether Windows accepted the setting.</returns>
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate bool AllowDarkModeForWindowDelegate(IntPtr window, bool allow);
        /// <summary>Private no-argument theme refresh API.</summary>
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate void VoidThemeDelegate();
        /// <summary>Native comctl32 subclass procedure callback.</summary>
        /// <param name="window">Window receiving the message.</param><param name="message">Windows message identifier.</param>
        /// <param name="wParam">Message-specific first value.</param><param name="lParam">Message-specific second value.</param>
        /// <param name="subclassId">Identifier supplied when the subclass was installed.</param><param name="referenceData">Subclass caller data.</param>
        /// <returns>Window-procedure result.</returns>
        internal delegate IntPtr SubclassCallback(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, UIntPtr subclassId, IntPtr referenceData);
        /// <summary>Callback invoked for Windows accessibility events watched by this theme.</summary>
        /// <param name="hook">Installed event hook.</param><param name="eventType">Event identifier.</param><param name="window">Associated window.</param>
        /// <param name="objectId">Accessible object identifier.</param><param name="childId">Accessible child identifier.</param>
        /// <param name="eventThread">Thread that raised the event.</param><param name="eventTime">Event timestamp.</param>
        internal delegate void WinEventCallback(IntPtr hook, uint eventType, IntPtr window, int objectId, int childId, uint eventThread, uint eventTime);

        /// <summary>Enumerates child windows owned by a native parent.</summary><param name="parent">Parent handle.</param><param name="callback">Enumeration callback.</param><param name="parameter">Callback context.</param><returns>Native enumeration result.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        /// <summary>Retrieves a native window's class name.</summary><param name="window">Window handle.</param><param name="text">Output text buffer.</param><param name="capacity">Buffer capacity.</param><returns>Characters copied.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        /// <summary>Retrieves a native window's caption text.</summary><param name="window">Window handle.</param><param name="text">Output text buffer.</param><param name="capacity">Buffer capacity.</param><returns>Characters copied.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        /// <summary>Sends a Windows message and waits for its window-procedure result.</summary><param name="window">Target handle.</param><param name="message">Message identifier.</param><param name="wParam">First message value.</param><param name="lParam">Second message value.</param><returns>Window-procedure result.</returns>
        [DllImport("user32.dll")] private static extern int SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        /// <summary>Posts a Windows message to the target window's queue.</summary><param name="window">Target handle.</param><param name="message">Message identifier.</param><param name="wParam">First message value.</param><param name="lParam">Second message value.</param><returns>Whether the message was posted.</returns>
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        /// <summary>Reads a native window style or extended style.</summary><param name="window">Window handle.</param><param name="index">Style selector.</param><returns>Style bits.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowStyle(IntPtr window, int index);
        /// <summary>Gets a window's client rectangle in client coordinates.</summary><param name="window">Window handle.</param><param name="rectangle">Receives the rectangle.</param><returns>Whether the rectangle was read.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out NativeRect rectangle);
        /// <summary>Finds a child window by class and optional caption.</summary><param name="parent">Parent handle.</param><param name="after">Previous child handle used to continue the search.</param><param name="className">Class name, or null.</param><param name="caption">Caption, or null.</param><returns>Matching child handle, or zero.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string caption);
        /// <summary>Fills a rectangle in a device context with a brush.</summary><param name="deviceContext">Target device context.</param><param name="rectangle">Rectangle to fill.</param><param name="brush">Brush handle.</param><returns>Native operation result.</returns>
        [DllImport("user32.dll")] private static extern int FillRect(IntPtr deviceContext, ref NativeRect rectangle, IntPtr brush);
        /// <summary>Creates a solid GDI brush from a COLORREF value.</summary><param name="color">RGB color value.</param><returns>Brush handle, or zero on failure.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);
        /// <summary>Deletes a GDI object previously created by the add-in.</summary><param name="value">GDI object handle.</param><returns>Whether it was deleted.</returns>
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
        /// <summary>Saves the current state of a device context.</summary><param name="deviceContext">Device context to save.</param><returns>Saved state identifier, or zero on failure.</returns>
        [DllImport("gdi32.dll")] private static extern int SaveDC(IntPtr deviceContext);
        /// <summary>Restores a device context to a saved state.</summary><param name="deviceContext">Device context to restore.</param><param name="saved">Saved state identifier.</param><returns>Whether restoration succeeded.</returns>
        [DllImport("gdi32.dll")] private static extern bool RestoreDC(IntPtr deviceContext, int saved);
        /// <summary>Sets the text color for subsequent GDI drawing.</summary><param name="deviceContext">Target device context.</param><param name="color">COLORREF value.</param><returns>Previous text color.</returns>
        [DllImport("gdi32.dll")] private static extern int SetTextColor(IntPtr deviceContext, int color);
        /// <summary>Sets the text background color for subsequent GDI drawing.</summary><param name="deviceContext">Target device context.</param><param name="color">COLORREF value.</param><returns>Previous background color.</returns>
        [DllImport("gdi32.dll")] private static extern int SetBkColor(IntPtr deviceContext, int color);
        /// <summary>Invalidates and optionally redraws a window or selected region.</summary><param name="window">Window handle.</param><param name="update">Update rectangle, or zero.</param><param name="region">Update region, or zero.</param><param name="flags">Redraw behavior flags.</param><returns>Whether the redraw request succeeded.</returns>
        [DllImport("user32.dll")] private static extern bool RedrawWindow(IntPtr window, IntPtr update, IntPtr region, uint flags);
        /// <summary>Checks whether one window is a descendant of another.</summary><param name="parent">Candidate ancestor handle.</param><param name="window">Candidate child handle.</param><returns>Whether the relationship holds.</returns>
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr window);
        /// <summary>Gets the process and thread that created a window.</summary><param name="window">Window handle.</param><param name="processId">Receives the process identifier.</param><returns>Thread identifier.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        /// <summary>Gets the current native thread identifier.</summary><returns>Calling thread identifier.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        /// <summary>Gets a related window according to a Windows relationship selector.</summary><param name="window">Starting window.</param><param name="command">Relationship selector.</param><returns>Related handle or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        /// <summary>Gets a specified ancestor window.</summary><param name="window">Starting window.</param><param name="flags">Ancestor selector.</param><returns>Ancestor handle or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        /// <summary>Installs a hook for a range of accessibility events.</summary>
        /// <param name="eventMin">First event identifier.</param><param name="eventMax">Last event identifier.</param><param name="module">Callback module, or zero.</param>
        /// <param name="callback">Event callback.</param><param name="processId">Process filter, or zero.</param><param name="threadId">Thread filter, or zero.</param>
        /// <param name="flags">Hook behavior flags.</param><returns>Hook handle, or zero on failure.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr module, WinEventCallback callback, uint processId, uint threadId, uint flags);
        /// <summary>Removes an accessibility event hook.</summary><param name="hook">Hook handle.</param><returns>Whether it was removed.</returns>
        [DllImport("user32.dll")] private static extern bool UnhookWinEvent(IntPtr hook);
        /// <summary>Installs a subclass procedure on a window.</summary><param name="window">Window handle.</param><param name="callback">Subclass callback.</param><param name="subclassId">Unique identifier for this subclass.</param><param name="referenceData">Callback context.</param><returns>Whether installation succeeded.</returns>
        [DllImport("comctl32.dll")] private static extern bool SetWindowSubclass(IntPtr window, SubclassCallback callback, UIntPtr subclassId, IntPtr referenceData);
        /// <summary>Removes a previously installed subclass procedure.</summary><param name="window">Window handle.</param><param name="callback">Subclass callback used during installation.</param><param name="subclassId">Subclass identifier.</param><returns>Whether removal succeeded.</returns>
        [DllImport("comctl32.dll")] private static extern bool RemoveWindowSubclass(IntPtr window, SubclassCallback callback, UIntPtr subclassId);
        /// <summary>Calls the next procedure in the window subclass chain.</summary><param name="window">Window receiving the message.</param><param name="message">Message identifier.</param><param name="wParam">First message value.</param><param name="lParam">Second message value.</param><returns>Window-procedure result.</returns>
        [DllImport("comctl32.dll")] private static extern IntPtr DefSubclassProc(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        /// <summary>Sets or clears the visual style theme for a native window.</summary><param name="window">Window handle.</param><param name="appName">Theme application name, or null.</param><param name="idList">Theme class list, or null.</param><returns>HRESULT from uxtheme.</returns>
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] private static extern int SetWindowTheme(IntPtr window, string appName, string idList);
        /// <summary>Sets a Desktop Window Manager attribute for a native window.</summary><param name="window">Window handle.</param><param name="attribute">DWM attribute identifier.</param><param name="value">Attribute value.</param><param name="size">Value size in bytes.</param><returns>HRESULT from DWM.</returns>
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        /// <summary>Gets a loaded module handle without loading a new module.</summary><param name="moduleName">Module name, or null for the current process.</param><returns>Module handle, or zero.</returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string moduleName);
        /// <summary>Resolves a function exported by ordinal from a native module.</summary><param name="module">Loaded module handle.</param><param name="ordinal">Export ordinal represented as a pointer-sized integer.</param><returns>Function pointer, or zero.</returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, ExactSpelling = true)] private static extern IntPtr GetProcAddress(IntPtr module, IntPtr ordinal);

        /// <summary>Native OSVERSIONINFOEX-compatible structure returned by RtlGetVersion.</summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct NativeOsVersion
        {
            /// <summary>Structure size, Windows version tuple, and platform identifier.</summary>
            internal uint Size, Major, Minor, Build, Platform;
            /// <summary>Service-pack string embedded in the native structure.</summary>
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] internal string ServicePack;
        }
        /// <summary>Reads the actual Windows version without compatibility-manifest version virtualization.</summary>
        /// <param name="version">Receives the native OS version structure.</param><returns>NTSTATUS result code.</returns>
        [DllImport("ntdll.dll", ExactSpelling = true)] private static extern int RtlGetVersion(ref NativeOsVersion version);

        /// <summary>Defines the read client callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="rectangle">The rectangle used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate bool ReadClient(IntPtr window, out NativeRect rectangle);
        /// <summary>Defines the read window thread callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="process">The process used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate uint ReadWindowThread(IntPtr window, out uint process);
        /// <summary>Defines the window attribute callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="attribute">The attribute used by this operation.</param>
/// <param name="value">The value used by this operation.</param>
/// <param name="size">The size used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate int WindowAttribute(IntPtr window, int attribute, ref int value, int size);
        /// <summary>Defines the read os version callback.</summary>
/// <param name="version">The version used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate int ReadOsVersion(ref NativeOsVersion version);
        /// <summary>Defines the paint region callback.</summary>
/// <param name="dc">The dc used by this operation.</param>
/// <param name="rectangle">The rectangle used by this operation.</param>
/// <param name="brush">The brush used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate int PaintRegion(IntPtr dc, ref NativeRect rectangle, IntPtr brush);
        /// <summary>Defines the paint surface callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="client">Indicates whether client is enabled.</param>
/// <param name="dc">The dc used by this operation.</param>
/// <param name="hosted">Indicates whether hosted is enabled.</param>
/// <param name="preserve">Indicates whether preserve is enabled.</param>
/// <param name="code">Indicates whether code is enabled.</param>
        internal delegate void PaintSurface(IntPtr window, bool client, IntPtr dc, bool hosted, bool preserve, bool code);
        /// <summary>Stores the enumerate children used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, EnumWindowCallback, IntPtr, bool> EnumerateChildren = EnumChildWindows;
        /// <summary>Stores the read class name,read window text used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, StringBuilder, int, int> ReadClassName = GetClassName, ReadWindowText = GetWindowText;
        /// <summary>Stores the native send used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, uint, IntPtr, IntPtr, int> NativeSend = SendMessage;
        /// <summary>Stores the native post used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, uint, IntPtr, IntPtr, bool> NativePost = PostMessage;
        /// <summary>Stores the read style used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, int, int> ReadStyle = GetWindowStyle;
        /// <summary>Stores the read client bounds used by VbeNativeTheme.</summary>
        internal static ReadClient ReadClientBounds = GetClientRect;
        /// <summary>Stores the find child used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, IntPtr, string, string, IntPtr> FindChild = FindWindowEx;
        /// <summary>Stores the paint background used by VbeNativeTheme.</summary>
        internal static PaintRegion PaintBackground = FillRect;
        /// <summary>Stores the new brush used by VbeNativeTheme.</summary>
        internal static Func<int, IntPtr> NewBrush = CreateSolidBrush;
        /// <summary>Stores the release brush used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, bool> ReleaseBrush = DeleteObject;
        /// <summary>Stores the save device context used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, int> SaveDeviceContext = SaveDC;
        /// <summary>Stores the restore device context used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, int, bool> RestoreDeviceContext = RestoreDC;
        /// <summary>Stores the foreground color,background color used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, int, int> ForegroundColor = SetTextColor, BackgroundColor = SetBkColor;
        /// <summary>Stores the redraw used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, IntPtr, IntPtr, uint, bool> Redraw = RedrawWindow;
        /// <summary>Stores the child relation used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, IntPtr, bool> ChildRelation = IsChild;
        /// <summary>Stores the window thread used by VbeNativeTheme.</summary>
        internal static ReadWindowThread WindowThread = GetWindowThreadProcessId;
        /// <summary>Stores the current thread used by VbeNativeTheme.</summary>
        internal static Func<uint> CurrentThread = GetCurrentThreadId;
        /// <summary>Stores the window relation,ancestor used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, uint, IntPtr> WindowRelation = GetWindow, Ancestor = GetAncestor;
        /// <summary>Stores the install window hook used by VbeNativeTheme.</summary>
        internal static Func<uint, uint, IntPtr, WinEventCallback, uint, uint, uint, IntPtr> InstallWindowHook = SetWinEventHook;
        /// <summary>Stores the remove window hook used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, bool> RemoveWindowHook = UnhookWinEvent;
        /// <summary>Stores the install subclass used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, SubclassCallback, UIntPtr, IntPtr, bool> InstallSubclass = SetWindowSubclass;
        /// <summary>Stores the remove subclass used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, SubclassCallback, UIntPtr, bool> RemoveSubclass = RemoveWindowSubclass;
        /// <summary>Stores the native procedure used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, uint, IntPtr, IntPtr, IntPtr> NativeProcedure = DefSubclassProc;
        /// <summary>Stores the set native theme used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, string, string, int> SetNativeTheme = SetWindowTheme;
        /// <summary>Stores the set attribute used by VbeNativeTheme.</summary>
        internal static WindowAttribute SetAttribute = DwmSetWindowAttribute;
        /// <summary>Stores the theme module used by VbeNativeTheme.</summary>
        internal static Func<string, IntPtr> ThemeModule = GetModuleHandle;
        /// <summary>Stores the native entry point used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, IntPtr, IntPtr> NativeEntryPoint = GetProcAddress;
        /// <summary>Stores the read version used by VbeNativeTheme.</summary>
        internal static ReadOsVersion ReadVersion = RtlGetVersion;
        /// <summary>Stores the draw chrome used by VbeNativeTheme.</summary>
        internal static PaintSurface DrawChrome = VbeNativeChrome.Paint;
        /// <summary>Stores the draw border,draw combo used by VbeNativeTheme.</summary>
        internal static Action<IntPtr> DrawBorder = VbeNativeChrome.PaintBorder, DrawCombo = VbeNativeChrome.PaintComboButton;
        /// <summary>Stores the draw property row used by VbeNativeTheme.</summary>
        internal static Action<IntPtr, NativeRect> DrawPropertyRow = VbeNativeChrome.PaintPropertyRow;

        /// <summary>Reads the native Windows major, minor, and build version.</summary>
        /// <returns>Native Windows version tuple.</returns>
        /// <exception cref="PlatformNotSupportedException">Windows version information could not be read.</exception>
        internal static Version ReadNativeWindowsVersion()
        {
            var version = new NativeOsVersion { Size = (uint)Marshal.SizeOf(typeof(NativeOsVersion)), ServicePack = string.Empty };
            if (ReadVersion(ref version) != 0)
                throw new PlatformNotSupportedException("Cannot verify the Windows dark-mode ABI.");
            return new Version((int)version.Major, (int)version.Minor, (int)version.Build);
        }

        /// <summary>Checks whether a Windows version is qualified for the preferred-app-mode private API ABI.</summary>
        /// <param name="version">Native Windows version to check.</param><returns>Whether the supported Windows 10 build range is met.</returns>
        internal static bool SupportsPreferredAppMode(Version version)
        {
            // Ordinal 135 used a different signature in Windows 10 1809.
            // Future major versions must be qualified before calling private APIs.
            return version != null && version.Major == 10 && version.Minor == 0 && version.Build >= 18362;
        }

        /// <summary>Native rectangle using left, top, right, and bottom edge coordinates.</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct NativeRect { /// <summary>Bounds in left, top, right, bottom coordinate order.</summary>
internal int Left, Top, Right, Bottom; }

        /// <summary>Layout-compatible WM_DRAWITEM payload for an owner-drawn control.</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct NativeDrawItem
        {
            /// <summary>Control type, identifier, item identifier, draw action, and item state.</summary>
            internal uint ControlType, ControlId, ItemId, Action, State;
            /// <summary>Control window and drawing device context.</summary>
            internal IntPtr Window, DeviceContext;
            /// <summary>Item bounds in the device context.</summary>
            internal NativeRect Bounds;
            /// <summary>Control-defined item data pointer.</summary>
            internal UIntPtr Data;
        }

        /// <summary>Unique identifier and callback instances retained for installed native hooks.</summary>
        private static readonly UIntPtr SubclassId = new UIntPtr(0x56424544);
        /// <summary>Subclass callback installed on eligible VBE windows.</summary>
        private static readonly SubclassCallback Subclass = ThemeWindowProcedure;
        /// <summary>WinEvent callback retained for the lifetime of the event hook.</summary>
        private static readonly WinEventCallback WindowEvent = WindowEventReceived;
        /// <summary>Synchronization object protecting native window tracking collections.</summary>
        private static readonly object Sync = new object();
        /// <summary>Windows currently opted into the native dark theme.</summary>
        private static readonly List<IntPtr> themedWindows = new List<IntPtr>();
        /// <summary>Windows with this class's subclass callback installed.</summary>
        private static readonly List<IntPtr> subclassedWindows = new List<IntPtr>();
        /// <summary>Command bars with a deferred chrome repaint queued.</summary>
        private static readonly HashSet<IntPtr> pendingChrome = new HashSet<IntPtr>();
        /// <summary>Pane captions with a deferred repaint queued.</summary>
        private static readonly HashSet<IntPtr> pendingCaptions = new HashSet<IntPtr>();
        /// <summary>Native toolbar paint and deferred-paint diagnostic counters.</summary>
        private static int toolbarPaintCount, toolbarDeferredPaintCount;
        /// <summary>Whether the opt-in local refresh experiment is enabled for the current process.</summary>
        private static bool localChromeRefresh;
        /// <summary>Direct renderers attached to supported native Properties tab controls.</summary>
        private static readonly Dictionary<IntPtr, VbeNativePropertyTabs> propertyTabs = new Dictionary<IntPtr, VbeNativePropertyTabs>();
        /// <summary>Original text and background colors saved for a native child control.</summary>
        private sealed class ControlPalette
        {
            /// <summary>Background, foreground, and text-background COLORREF values.</summary>
            internal int Background, Foreground, TextBackground;
        }
        /// <summary>Original control colors, keyed by child window handle.</summary>
        private static readonly Dictionary<IntPtr, ControlPalette> originalControlColors = new Dictionary<IntPtr, ControlPalette>();
        /// <summary>Resolved private API that opts one native window into dark mode.</summary>
        private static AllowDarkModeForWindowDelegate allowDarkMode;
        /// <summary>Resolved private API that controls process preferred theme mode.</summary>
        private static SetPreferredAppModeDelegate setPreferredMode;
        /// <summary>Resolved private API that flushes cached native menu themes.</summary>
        private static VoidThemeDelegate flushMenuThemes;
        /// <summary>GDI brush used to erase native control backgrounds.</summary>
        private static IntPtr backgroundBrush;
        /// <summary>Main VBE editor window handle currently managed by this theme.</summary>
        private static IntPtr editorWindow;
        /// <summary>Process theme mode value saved before the experiment changed it.</summary>
        private static int previousPreferredMode;
        /// <summary>Installed WinEvent hook that tracks VBE windows created or shown later.</summary>
        private static IntPtr windowEventHook;
        /// <summary>Whether this type changed the process preferred mode and must restore it.</summary>
        private static bool preferredModeChanged;
        /// <summary>Deferred palette transaction service, when the VBE automation object is available.</summary>
        private static VbeNativePalette nativePalette;
        /// <summary>Handle of the Immediate window captured during initialization.</summary>
        private static IntPtr immediateWindow;
        /// <summary>Caption captured with the Immediate window handle.</summary>
        private static string immediateCaption;

                /// <summary>Returns whether the explicitly gated experiment is enabled for this host process.</summary>
        /// <returns><see langword="true"/> when the process environment variable equals <c>1</c>.</returns>
        internal static bool ExperimentEnabled()
        {
            return string.Equals(Environment.GetEnvironmentVariable(ExperimentVariable), "1", StringComparison.Ordinal);
        }

        /// <summary>Stores the apply native theme used by VbeNativeTheme.</summary>
        internal static Func<IntPtr, int> ApplyNativeTheme = Apply;
        /// <summary>Stores the create palette used by VbeNativeTheme.</summary>
        internal static Func<object, IntPtr, VbeNativePalette> CreatePalette = (vbe, editor) => new VbeNativePalette(vbe, editor);

                /// <summary>Stores the VBE owner and applies the persisted or explicitly gated preference.</summary>
        /// <param name="editor">VBE main-window handle.</param><param name="enabled">Persisted preference for native styling.</param>
        /// <param name="vbe">Optional VBE automation object used for Immediate-window detection and palette recovery.</param>
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
        /// <param name="enabled">Whether to apply the dark style or restore the tracked native state.</param>
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
                throw new PlatformNotSupportedException("Native VBE dark mode requires the Windows 10 1903 or newer dark-mode ABI. Detected: " + windowsVersion.ToString());
            IntPtr themeModule = ThemeModule("uxtheme.dll");
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
            EnumerateChildren(editor, (window, parameter) => { windows.Add(window); return true; }, IntPtr.Zero);
            foreach (IntPtr window in windows) ApplyWindow(window);
            int enabled = 1;
            int titleResult = SetAttribute(editor, 20, ref enabled, sizeof(int));
            if (titleResult != 0) SetAttribute(editor, 19, ref enabled, sizeof(int));
            flushMenuThemes?.Invoke();
            editorWindow = editor;
            windowEventHook = InstallWindowHook(EventObjectCreate, EventObjectShow, IntPtr.Zero, WindowEvent,
                (uint)System.Diagnostics.Process.GetCurrentProcess().Id, 0, WinEventOutOfContext);
            Redraw(editor, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
            return themedWindows.Count;
        }

                /// <summary>Removes installed native hooks and restores tracked windows and process theme preferences.</summary>
        /// <returns><see langword="true"/> when the renderer stopped and managed state was reset; otherwise <see langword="false"/>.</returns>
        internal static bool Reset()
        {
            lock (Sync)
            {
                if (!VbeNativeRenderer.Stop()) return false;
                if (windowEventHook != IntPtr.Zero) { RemoveWindowHook(windowEventHook); windowEventHook = IntPtr.Zero; }
                foreach (VbeNativePropertyTabs tabs in propertyTabs.Values) tabs.Dispose();
                propertyTabs.Clear();
                foreach (IntPtr window in subclassedWindows) RemoveSubclass(window, Subclass, SubclassId);
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
                        SetAttribute(window, 20, ref disabled, sizeof(int));
                        SetAttribute(window, 19, ref disabled, sizeof(int));
                    }
                    SetNativeTheme(window, null, null);
                    RestoreControlPalette(window, WindowClass(window));
                    NativeSend(window, WmThemeChanged, IntPtr.Zero, IntPtr.Zero);
                }
                themedWindows.Clear();
                originalControlColors.Clear();
                if (backgroundBrush != IntPtr.Zero) { ReleaseBrush(backgroundBrush); backgroundBrush = IntPtr.Zero; }
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
                SetAttribute(editorWindow, 20, ref disabled, sizeof(int));
                SetAttribute(editorWindow, 19, ref disabled, sizeof(int));
                Redraw(editorWindow, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
            }
            return true;
        }

                /// <summary>Disconnects the experiment while preserving state if native hooks cannot be stopped safely.</summary>
        /// <returns><see langword="true"/> when disconnection completed; otherwise <see langword="false"/>.</returns>
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

        /// <summary>Applies dark control colors to supported child controls and saves their prior values.</summary>
        /// <param name="window">Child control handle.</param>
        /// <param name="className">Native window class used to select control-specific color messages.</param>
        internal static void ApplyControlPalette(IntPtr window, string className)
        {
            if (originalControlColors.ContainsKey(window)) return;
            IntPtr background = new IntPtr(Background);
            var original = new ControlPalette();
            switch (className)
            {
                case "SysTreeView32":
                    original.Background = NativeSend(window, TreeSetBackground, IntPtr.Zero, background);
                    original.Foreground = NativeSend(window, TreeSetText, IntPtr.Zero, new IntPtr(Foreground));
                    break;
                case "SysListView32":
                    original.Background = NativeSend(window, 0x1000, IntPtr.Zero, IntPtr.Zero);
                    original.Foreground = NativeSend(window, 0x1023, IntPtr.Zero, IntPtr.Zero);
                    original.TextBackground = NativeSend(window, 0x1025, IntPtr.Zero, IntPtr.Zero);
                    NativeSend(window, ListSetBackground, IntPtr.Zero, background);
                    NativeSend(window, ListSetText, IntPtr.Zero, new IntPtr(Foreground));
                    NativeSend(window, ListSetTextBackground, IntPtr.Zero, background);
                    break;
                case "RichEdit20A":
                case "RICHEDIT60W":
                    original.Background = NativeSend(window, RichEditSetBackground, IntPtr.Zero, background);
                    break;
                default: return;
            }
            originalControlColors.Add(window, original);
        }

        /// <summary>Restores a tracked control's original colors after the experimental theme is removed.</summary>
        /// <param name="window">Child control handle.</param>
        /// <param name="className">Native window class used to restore control-specific color messages.</param>
        internal static void RestoreControlPalette(IntPtr window, string className)
        {
            ControlPalette original;
            if (!originalControlColors.TryGetValue(window, out original)) return;
            IntPtr background = new IntPtr(original.Background);
            IntPtr foreground = new IntPtr(original.Foreground);
            switch (className)
            {
                case "SysTreeView32":
                    NativeSend(window, TreeSetBackground, IntPtr.Zero, background);
                    NativeSend(window, TreeSetText, IntPtr.Zero, foreground);
                    break;
                case "SysListView32":
                    NativeSend(window, ListSetBackground, IntPtr.Zero, background);
                    NativeSend(window, ListSetText, IntPtr.Zero, foreground);
                    NativeSend(window, ListSetTextBackground, IntPtr.Zero, new IntPtr(original.TextBackground));
                    break;
                case "RichEdit20A":
                case "RICHEDIT60W":
                    NativeSend(window, RichEditSetBackground, IntPtr.Zero, background);
                    break;
            }
            originalControlColors.Remove(window);
        }

        /// <summary>Registers one VBE-owned window with the native theme and its required redraw handlers.</summary>
        /// <param name="window">Window to theme.</param>
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
                    ReadWindowText(window, caption, caption.Capacity);
                    if (caption.ToString() == immediateCaption) immediateWindow = window;
                }
                if (string.IsNullOrEmpty(className) || (IsManagedAddInWindow(className) && className != "GenericPane")) return;
                bool propertiesControl = WindowClass(GetAncestorParent(window)) == "wndclass_pbrs";
                bool dialogControl = WindowClass(Ancestor(window, GetAncestorRoot)) == "#32770";
                bool scopedControl = className == "ListBox" ? propertiesControl :
                    className != "SysTabControl32" || propertiesControl || dialogControl;
                if (scopedControl && ShouldSubclass(className))
                {
                    if (!InstallSubclass(window, Subclass, SubclassId, IntPtr.Zero)) return;
                    subclassedWindows.Add(window);
                }
                if (className == "SysTabControl32" && propertiesControl)
                    TryAttachPropertyTabs(window);
                allowDarkMode(window, true);
                if (className == "#32770")
                {
                    int enabled = 1;
                    if (SetAttribute(window, 20, ref enabled, sizeof(int)) != 0)
                        SetAttribute(window, 19, ref enabled, sizeof(int));
                }
                SetNativeTheme(window, "DarkMode_Explorer", null);
                // Keep the standard tab interaction and geometry. Supported
                // Properties tabs supply their final colors before drawing text;
                // other tabs retain the existing classic-renderer fallback.
                // The themed checkbox/group-box renderer ignores the dialog text
                // color. Classic controls honor WM_CTLCOLOR from their parent.
                if ((className == "SysTabControl32" && (propertiesControl || dialogControl)) ||
                    (className == "Button" && dialogControl && IsDialogLabelButton(window)))
                    SetNativeTheme(window, "", "");
                ApplyControlPalette(window, className);
                if (className == "GenericPane")
                    LoadLog.Write("Native theme GenericPane: " + window + ", subclass=" + subclassedWindows.Contains(window));
                themedWindows.Add(window);
                NativeSend(window, WmThemeChanged, IntPtr.Zero, IntPtr.Zero);
            }
        }

        /// <summary>Tracks newly shown VBE windows and attaches rendering support when they belong to the editor.</summary>
        /// <param name="hook">WinEvent hook handle.</param><param name="eventType">Received WinEvent identifier.</param>
        /// <param name="window">Window associated with the event.</param><param name="objectId">Accessible object identifier.</param>
        /// <param name="childId">Accessible child identifier.</param><param name="eventThread">Thread that raised the event.</param>
        /// <param name="eventTime">Event timestamp supplied by Windows.</param>
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
                EnumerateChildren(window, (child, parameter) =>
                {
                    ApplyWindow(child);
                    return true;
                }, IntPtr.Zero);
            }
            Redraw(window, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
        }

        /// <summary>Checks whether a window is the VBE editor or a descendant owned by it.</summary>
        /// <param name="window">Window being checked.</param>
        /// <returns><see langword="true"/> when the window belongs to the active VBE editor.</returns>
        private static bool BelongsToEditor(IntPtr window)
        {
            if (editorWindow == IntPtr.Zero || window == IntPtr.Zero) return false;
            if (window == editorWindow || ChildRelation(editorWindow, window)) return true;
            IntPtr root = Ancestor(window, GetAncestorRoot);
            IntPtr owner = WindowRelation(root == IntPtr.Zero ? window : root, GetWindowOwner);
            for (int depth = 0; owner != IntPtr.Zero && depth < 16; depth++)
            {
                if (owner == editorWindow || ChildRelation(editorWindow, owner)) return true;
                owner = WindowRelation(owner, GetWindowOwner);
            }
            return false;
        }

        /// <summary>Identifies native VBE window classes whose messages need experimental handling.</summary>
        /// <param name="className">Native window class name.</param>
        /// <returns><see langword="true"/> for a class handled by the theme subclass.</returns>
        internal static bool ShouldSubclass(string className)
        {
            return className == "#32770" || className == "wndclass_desked_gsk" || className == "MDIClient" ||
                className == "VbaWindow" || className == "PROJECT" || className == "wndclass_pbrs" ||
                className == "ToolsPalette" || className == "VBSlider" || className == "GenericPane" ||
                className == "MsoCommandBar" || className == "MsoCommandBarPopup" || className == "MsoCommandBarDock" || className == "SysTabControl32" ||
                className == "ListBox" || className == "SysTreeView32" || className == "ComboBox" || className == "Edit";
        }

        /// <summary>Removes a destroyed window from all tracking collections without releasing shared process state.</summary>
        /// <param name="window">Destroyed window handle.</param>
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

        /// <summary>Handles experimental theme and repaint messages before delegating to the native window procedure.</summary>
        /// <param name="window">Window receiving the message.</param><param name="message">Windows message identifier.</param>
        /// <param name="wParam">Message-specific first value.</param><param name="lParam">Message-specific second value.</param>
        /// <param name="subclassId">Identifier supplied when the subclass was installed.</param>
        /// <param name="referenceData">Caller data supplied when the subclass was installed.</param>
        /// <returns>Window-procedure result for the handled or delegated message.</returns>
        private static IntPtr ThemeWindowProcedure(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, UIntPtr subclassId, IntPtr referenceData)
        {
            if (message == WmQueryToolbarPaint && ExperimentEnabled())
                return new IntPtr(wParam.ToInt64() == 1 ? toolbarDeferredPaintCount : toolbarPaintCount);
            if (message == WmRefreshCaption)
            {
                if (!pendingCaptions.Remove(window)) return IntPtr.Zero;
                if (subclassedWindows.Contains(window))
                {
                    DrawChrome(window, false, IntPtr.Zero, WindowClass(window) == "GenericPane", false, false);
                    DrawBorder(window);
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
                    DrawChrome(window, false, IntPtr.Zero, false, false, false);
                if (refreshedClass == "MsoCommandBar") toolbarDeferredPaintCount++;
                PaintChrome(window, refreshedClass, IntPtr.Zero);
                return IntPtr.Zero;
            }
            if (message == WmNcDestroy)
            {
                RemoveSubclass(window, Subclass, SubclassId);
                ForgetWindow(window);
                return NativeProcedure(window, message, wParam, lParam);
            }
            VbeNativePropertyTabs directTabs;
            if (propertyTabs.TryGetValue(window, out directTabs))
            {
                if (message == WmQueryPropertyTabPaint && ExperimentEnabled())
                    return new IntPtr(wParam.ToInt64() == 2 ? (VbeNativePropertyTabs.CanRender(window) ? 1 : -1) :
                        wParam.ToInt64() == 1 ? directTabs.PrintCount : directTabs.PaintCount);
                IntPtr handled;
                if (directTabs.TryHandleMessage(message, wParam, lParam, out handled)) return handled;
                IntPtr nativeResult = NativeProcedure(window, message, wParam, lParam);
                directTabs.AfterNativeMessage(message, wParam, lParam);
                return nativeResult;
            }
            string className = message == WmEraseBackground ? WindowClass(window) : null;
            if (message == WmEraseBackground && (className == "#32770" || className == "MDIClient" || className == "VBSlider"))
            {
                NativeRect rectangle;
                if (ReadClientBounds(window, out rectangle))
                {
                    EnsureBackgroundBrush();
                    PaintBackground(wParam, ref rectangle, backgroundBrush);
                    return new IntPtr(1);
                }
            }
            if (message == WmCtlColorEdit || message == WmCtlColorListBox || message == WmCtlColorButton ||
                message == WmCtlColorDialog || message == WmCtlColorScrollBar || message == WmCtlColorStatic)
            {
                EnsureBackgroundBrush();
                ForegroundColor(wParam, Foreground);
                BackgroundColor(wParam, Background);
                return backgroundBrush;
            }
            if (message == 0x002B && lParam != IntPtr.Zero && WindowClass(window) == "wndclass_pbrs")
            {
                var item = (NativeDrawItem)Marshal.PtrToStructure(lParam, typeof(NativeDrawItem));
                if (item.ControlType == 2 && WindowClass(item.Window) == "ListBox" &&
                    item.DeviceContext != IntPtr.Zero && (item.Action & 3) != 0)
                {
                    int saved = SaveDeviceContext(item.DeviceContext);
                    if (saved != 0)
                    {
                        try
                        {
                            // Render the native row on its original light palette,
                            // then translate that fresh row exactly once. Changing
                            // the whole list later would retouch ClearType pixels.
                            ForegroundColor(item.DeviceContext, 0);
                            BackgroundColor(item.DeviceContext, 0x00ffffff);
                            IntPtr drawn = NativeProcedure(window, message, wParam, lParam);
                            DrawPropertyRow(item.DeviceContext, item.Bounds);
                            return drawn;
                        }
                        finally { RestoreDeviceContext(item.DeviceContext, saved); }
                    }
                }
            }
            IntPtr result = NativeProcedure(window, message, wParam, lParam);
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
                DrawCombo(window);
            if (message == 0x0085 || message == 0x0086 || message == 0x000C || message == 0x0317)
            {
                if (window != editorWindow) DrawChrome(window, false, message == 0x0317 ? wParam : IntPtr.Zero, false, false, false);
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
                if (!NativePost(window, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero)) pendingChrome.Remove(window);
            }
            // Caption work is local and never triggers a code-surface conversion.
            // A Properties notification must not repaint unrelated Office toolbars.
            if (localChromeRefresh && (message == 0x0086 || message == 0x000c || message == 0x0047 || message == 0x0222))
            {
                string captionClass = WindowClass(window);
                if ((captionClass == "PROJECT" || captionClass == "wndclass_pbrs" || captionClass == "VbaWindow" || captionClass == "GenericPane") &&
                    pendingCaptions.Add(window) && !NativePost(window, WmRefreshCaption, IntPtr.Zero, IntPtr.Zero))
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
                        if (!NativePost(candidate, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero)) pendingChrome.Remove(candidate);
                }
            }
            return result;
        }

        /// <summary>Applies the selected chrome repaint path for a native VBE window class.</summary>
        /// <param name="window">Window to repaint.</param><param name="className">Native class name.</param>
        /// <param name="dc">Device context supplied by the current paint message, or zero.</param>
        private static void PaintChrome(IntPtr window, string className, IntPtr dc)
        {
            // A direct renderer must never be followed by image recoloring or a
            // screen-DC border pass, including already queued legacy refreshes.
            if (propertyTabs.ContainsKey(window)) return;
            if (className == "MsoCommandBar" && VbeNativeRenderer.Active &&
                ChildRelation(editorWindow, window)) return;
            if (className == "MsoCommandBar") toolbarPaintCount++;
            if (className == "ComboBox" && dc == IntPtr.Zero) DrawCombo(window);
            if (className == "GenericPane") DrawChrome(window, false, dc, true, false, false);
            bool codeSurface = IsCodeSurface(window, className);
            if (codeSurface) DrawChrome(window, true, dc, false, false, true);
            if (className == "MsoCommandBar" || className == "MsoCommandBarPopup" || className == "MsoCommandBarDock" || className == "SysTabControl32" ||
                (className == "VbaWindow" && !codeSurface))
                DrawChrome(window, true, dc, false, className == "VbaWindow", false);
            if (dc == IntPtr.Zero && window != editorWindow) DrawBorder(window);
        }

        /// <summary>Checks whether a dialog button uses a text label that should inherit its parent's text color.</summary>
        /// <param name="window">Native button handle.</param>
        /// <returns><see langword="true"/> for checkbox, radio, three-state, or group-box styles.</returns>
        private static bool IsDialogLabelButton(IntPtr window)
        {
            int buttonType = ReadStyle(window, -16) & 0x0f;
            // Checkbox, three-state, radio and group-box labels need the parent
            // text color. Push buttons keep their native dark visual style.
            return (buttonType >= 2 && buttonType <= 7) || buttonType == 9;
        }

        /// <summary>Gets the immediate parent window through the native ancestor API.</summary>
        /// <param name="window">Window whose parent is requested.</param>
        /// <returns>Parent handle, or zero when none exists.</returns>
        private static IntPtr GetAncestorParent(IntPtr window) { return Ancestor(window, 1); }

        /// <summary>Identifies VBE code panes, including the Immediate window and panes with a code toolbar.</summary>
        /// <param name="window">Candidate VBE window.</param><param name="className">Candidate's native class name.</param>
        /// <returns><see langword="true"/> when the window displays VBA code.</returns>
        private static bool IsCodeSurface(IntPtr window, string className)
        {
            return className == "VbaWindow" && (window == immediateWindow ||
                FindChild(window, IntPtr.Zero, "ObtbarWndClass", null) != IntPtr.Zero);
        }

        /// <summary>Posts deferred repaint messages to tracked command bars and captions inside a container.</summary>
        /// <param name="container">Native container whose chrome should be refreshed.</param>
        private static void QueueContainerChromeRefresh(IntPtr container)
        {
            foreach (IntPtr candidate in subclassedWindows.ToArray())
            {
                IntPtr dock = GetAncestorParent(candidate);
                if (WindowClass(candidate) == "MsoCommandBar" && WindowClass(dock) == "MsoCommandBarDock" &&
                    GetAncestorParent(dock) == container && pendingChrome.Add(candidate) &&
                    !NativePost(candidate, WmRefreshChrome, IntPtr.Zero, IntPtr.Zero))
                    pendingChrome.Remove(candidate);
                if (dock == container && WindowClass(candidate) == "GenericPane" && pendingCaptions.Add(candidate) &&
                    !NativePost(candidate, WmRefreshCaption, IntPtr.Zero, IntPtr.Zero))
                    pendingCaptions.Remove(candidate);
            }
        }

        /// <summary>Attaches the direct text renderer when the native Properties tabs meet its supported constraints.</summary>
        /// <param name="window">Candidate tab-control handle.</param>
        /// <returns><see langword="true"/> when a renderer is already attached or was created successfully.</returns>
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

        /// <summary>Checks that a window belongs to this process and the calling thread.</summary>
        /// <param name="window">Window handle to inspect.</param>
        /// <returns><see langword="true"/> when both process and thread match.</returns>
        private static bool IsCurrentWindowThread(IntPtr window)
        {
            uint processId;
            return WindowThread(window, out processId) == CurrentThread() &&
                processId == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        }

        /// <summary>Creates the shared background brush on first use.</summary>
        private static void EnsureBackgroundBrush()
        {
            if (backgroundBrush == IntPtr.Zero) backgroundBrush = NewBrush(Background);
        }

        /// <summary>Resolves a private theme API by ordinal and marshals it to the requested delegate type.</summary>
        /// <typeparam name="T">Delegate type matching the native ordinal's ABI.</typeparam>
        /// <param name="module">Loaded module that exports the ordinal.</param><param name="ordinal">Export ordinal.</param>
        /// <returns>The marshaled delegate, or <see langword="null"/> if the module or export is unavailable.</returns>
        private static T Resolve<T>(IntPtr module, int ordinal) where T : class
        {
            if (module == IntPtr.Zero) return null;
            IntPtr address = NativeEntryPoint(module, new IntPtr(ordinal));
            return address == IntPtr.Zero ? null : Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
        }

        /// <summary>Reads a native window's class name into a managed string.</summary>
        /// <param name="window">Window handle whose class is queried.</param>
        /// <returns>Class name, or an empty string when Windows returns no name.</returns>
        private static string WindowClass(IntPtr window)
        {
            var text = new StringBuilder(256);
            ReadClassName(window, text, text.Capacity);
            return text.ToString();
        }

        /// <summary>Recognizes managed add-in surface classes that should retain their own rendering.</summary>
        /// <param name="className">Native class name to inspect.</param>
        /// <returns><see langword="true"/> for Windows Forms, WPF wrapper, or GenericPane windows.</returns>
        internal static bool IsManagedAddInWindow(string className)
        {
            return className.StartsWith("WindowsForms10.", StringComparison.Ordinal) ||
                className.StartsWith("HwndWrapper[", StringComparison.Ordinal) ||
                className == "GenericPane";
        }
    }
}
