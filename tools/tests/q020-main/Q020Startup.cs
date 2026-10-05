using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

// Qualification-only definitions. No COM activation, keyboard, focus or desktop switching.
public static class Q020Startup
{
    public const string ToolbarText = "Les informations de la barre d'outils sont inconsistantes, Les réglages par défaut de la barre d'outils vont être utilisés.";
    private delegate bool Visitor(IntPtr window, IntPtr state);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumDesktopWindows(IntPtr desktop, Visitor callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr root, Visitor callback, IntPtr state);
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr OpenDesktop(string name, uint flags, bool inherit, uint access);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool SetThreadDesktop(IntPtr desktop);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool GetUserObjectInformation(IntPtr handle, int index, StringBuilder text, int bytes, out int needed);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll")] private static extern void SetLastError(uint value);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "SendMessageTimeoutW")]
    private static extern IntPtr ReadText(IntPtr window, uint message, IntPtr capacity, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr first, IntPtr second, uint flags, uint timeout, out UIntPtr result);
    public sealed class UiaChild { public int Pid, ControlType; public long NativeHandle; public string Name; public bool Enabled; }
    public sealed class ToolbarUia { public bool Complete; public int ExactTexts, EnabledOkButtons; public long NativeButton; public UiaChild[] Children; }
    public sealed class Window { public long Handle, Root; public uint Pid, Thread; public int Id; public string Class, Caption; public bool Visible, Enabled; public Window[] Children; public ToolbarUia Toolbar; }
    private static string DesktopName(IntPtr handle)
    {
        var name = new StringBuilder(256); int needed;
        if (handle == IntPtr.Zero || !GetUserObjectInformation(handle, 2, name, name.Capacity * 2, out needed))
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Exact private desktop name unproved.");
        return name.ToString();
    }
    private static void RequirePrivateModal(Window modal, string desktop)
    {
        // MAIN uses its configured exact-owner Default guard; private campaigns retain their sentinel guard.
        Type guard = null;
        if (desktop == "Default")
        {
            if (Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME") != "Default")
                throw new InvalidOperationException("Explicit MAIN desktop context is required.");
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                var type = assembly.GetType("Q020MainDesktopGate", false);
                if (type == null) continue;
                if (guard != null) throw new InvalidOperationException("Ambiguous loaded MAIN desktop guard.");
                guard = type;
            }
        }
        else
        {
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                var type = assembly.GetType("VBAi.Tests.Integration.IsolatedTestDesktop", false);
                if (type != null && assembly.GetName().Name == "VBAi.Desktop.Helper") { guard = type; break; }
            }
        }
        if (guard == null) throw new InvalidOperationException("Loaded exact desktop guard unavailable; no UIA observation.");
        var method = guard.GetMethod("RequireOfficeWindowInventory", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
        if (method == null) throw new InvalidOperationException("Exact private/input inventory guard unavailable.");
        method.Invoke(null, new object[] { desktop, modal.Pid, true, new IntPtr(modal.Handle) });
    }
    private static ToolbarUia ObserveToolbarUia(Window modal, DateTime deadline)
    {
        string desktopName = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
        if (string.IsNullOrEmpty(desktopName) || DesktopName(GetThreadDesktop(GetCurrentThreadId())) != desktopName)
            throw new InvalidOperationException("Original caller private desktop unproved.");
        IntPtr desktop = OpenDesktop(desktopName, 0, false, 0x41);
        if (desktop == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        ToolbarUia result = null; Exception failure = null;
        var observer = new Thread(() => {
            try {
                // Fresh MTA thread binds before UIA/COM. It never switches the input desktop.
                if (!SetThreadDesktop(desktop) || DesktopName(GetThreadDesktop(GetCurrentThreadId())) != desktopName)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "UIA observer private desktop binding failed.");
                RequirePrivateModal(modal, desktopName);
                var root = AutomationElement.FromHandle(new IntPtr(modal.Handle));
                var elements = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                if (elements.Count > 128) throw new InvalidOperationException("Owned toolbar UIA inventory limit exceeded.");
                var values = new List<UiaChild>(); int texts = 0, buttons = 0; long button = 0;
                for (int i = 0; i < elements.Count; i++)
                {
                    if (DateTime.UtcNow >= deadline) throw new TimeoutException("Original startup deadline expired during UIA read.");
                    RequirePrivateModal(modal, desktopName);
                    var info = elements[i].Current;
                    if (info.ProcessId != modal.Pid || info.Name == null || info.Name.Length > 2048)
                        throw new InvalidOperationException("Owned toolbar UIA identity/text bound changed.");
                    values.Add(new UiaChild { Pid = info.ProcessId, ControlType = info.ControlType.Id, NativeHandle = info.NativeWindowHandle, Name = info.Name, Enabled = info.IsEnabled });
                    if (info.ControlType == ControlType.Text && info.Name == ToolbarText) texts++;
                    if (info.ControlType == ControlType.Button && info.Name == "OK" && info.IsEnabled) { buttons++; button = info.NativeWindowHandle; }
                }
                RequirePrivateModal(modal, desktopName);
                var fresh = Observe(new IntPtr(modal.Handle), modal.Pid, true, deadline, true);
                if (fresh.Thread != modal.Thread || fresh.Class != "#32770" || fresh.Caption != "SOLIDWORKS" || !fresh.Visible || !fresh.Enabled)
                    throw new InvalidOperationException("Known toolbar modal changed during UIA read.");
                result = new ToolbarUia { Complete = true, ExactTexts = texts, EnabledOkButtons = buttons, NativeButton = button, Children = values.ToArray() };
            } catch (Exception error) { failure = error; }
        });
        observer.SetApartmentState(ApartmentState.MTA);
        bool started = false, joined = false;
        try {
            observer.Start(); started = true;
            joined = observer.Join((int)Math.Min(int.MaxValue, Math.Max(0, (deadline - DateTime.UtcNow).TotalMilliseconds)));
            if (!joined) throw new TimeoutException("Read-only UIA provider pending at original startup deadline; observer/desktop retained, no acknowledgment.");
        } finally {
            // Never close a handle still assigned to a pending observer, and never retry a failed close.
            if (!started || joined) { if (!CloseDesktop(desktop)) throw new Win32Exception(Marshal.GetLastWin32Error(), "UIA observer desktop close failed once."); }
        }
        if (failure != null) throw failure;
        return result;
    }
    private static Window Observe(IntPtr window, uint expectedPid, bool text, DateTime deadline, bool topLevel = false)
    {
        if (DateTime.UtcNow >= deadline) throw new TimeoutException("Original startup deadline expired during native observation.");
        uint pid; uint thread = GetWindowThreadProcessId(window, out pid);
        if (thread == 0 || pid != expectedPid) throw new InvalidOperationException("Owned HWND identity changed.");
        var cls = new StringBuilder(256);
        int length = GetClassName(window, cls, cls.Capacity);
        if (length <= 0 || length >= cls.Capacity - 1) throw new InvalidOperationException("Owned HWND class unproved/truncated.");
        var value = new Window { Handle = window.ToInt64(), Root = GetAncestor(window, 2).ToInt64(), Pid = pid, Thread = thread,
            Id = GetDlgCtrlID(window), Class = cls.ToString(), Visible = IsWindowVisible(window), Enabled = IsWindowEnabled(window), Children = new Window[0] };
        if (text)
        {
            var caption = new StringBuilder(2048);
            if (topLevel)
            {
                // Cross-process top-level caption reads use USER's stored caption,
                // avoiding WM_GETTEXT delivery to hidden/background Office windows.
                if (GetWindowText(window, caption, caption.Capacity) >= caption.Capacity - 1)
                    throw new InvalidOperationException("Owned top-level caption truncated.");
            }
            else
            {
                UIntPtr result; uint remaining = (uint)Math.Min(500, Math.Max(1, (deadline - DateTime.UtcNow).TotalMilliseconds));
                if (ReadText(window, 13, new IntPtr(caption.Capacity), caption, 2, remaining, out result) == IntPtr.Zero || result.ToUInt64() >= (ulong)caption.Capacity - 1)
                    throw new InvalidOperationException("Owned native control text read uncertain/truncated.");
            }
            value.Caption = caption.ToString();
        }
        return value;
    }
    public static Window[] Snapshot(int ownedPid, DateTime deadline)
    {
        var desktop = GetThreadDesktop(GetCurrentThreadId());
        if (desktop == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        var windows = new List<Window>(); Exception failure = null; int count = 0;
        Visitor callback = (window, state) => {
            try {
                if (++count > 8192 || DateTime.UtcNow >= deadline) throw new InvalidOperationException("Private inventory limit/deadline exceeded.");
                uint pid; if (GetWindowThreadProcessId(window, out pid) == 0 || pid == 0) throw new InvalidOperationException("Private inventory HWND identity incomplete.");
                if (pid == (uint)ownedPid) windows.Add(Observe(window, pid, true, deadline, true));
                return true;
            } catch (Exception error) { failure = error; SetLastError(0); return false; }
        };
        SetLastError(0); bool complete = EnumDesktopWindows(desktop, callback, IntPtr.Zero); int errorCode = Marshal.GetLastWin32Error(); GC.KeepAlive(callback);
        if (failure != null) throw failure;
        if (!complete) throw new Win32Exception(errorCode, "Private desktop enumeration incomplete.");
        foreach (var modal in windows)
        {
            if (!IsDialog(modal)) continue;
            var children = new List<Window>(); failure = null; bool stopped = false;
            Visitor child = (window, state) => {
                try {
                    if (children.Count >= 128) throw new InvalidOperationException("Owned modal control inventory limit exceeded.");
                    var value = Observe(window, (uint)ownedPid, false, deadline);
                    if (value.Thread != modal.Thread || value.Root != modal.Handle) throw new InvalidOperationException("Owned modal control ancestry/thread changed.");
                    if (value.Class == "Static" || value.Class == "Button") value = Observe(window, (uint)ownedPid, true, deadline);
                    children.Add(value); return true;
                } catch (Exception error) { failure = error; stopped = true; return false; }
            };
            // EnumChildWindows' return value is unused by its documented API; track our callback stop explicitly.
            EnumChildWindows(new IntPtr(modal.Handle), child, IntPtr.Zero); GC.KeepAlive(child);
            if (stopped || failure != null) throw failure ?? new InvalidOperationException("Modal inventory incomplete.");
            var fresh = Observe(new IntPtr(modal.Handle), (uint)ownedPid, true, deadline, true);
            if (fresh.Thread != modal.Thread || fresh.Class != modal.Class || fresh.Caption != modal.Caption) throw new InvalidOperationException("Modal changed during snapshot.");
            modal.Children = children.ToArray();
            if (modal.Class == "#32770" && modal.Caption == "SOLIDWORKS" && modal.Enabled)
                modal.Toolbar = ObserveToolbarUia(modal, deadline);
        }
        return windows.ToArray();
    }
    public static bool IsDialog(Window window) { return window.Visible && (window.Class == "#32770" || window.Class.IndexOf("Dialog", StringComparison.OrdinalIgnoreCase) >= 0); }
    public static Window KnownToolbarButton(Window[] windows)
    {
        var modals = new List<Window>(); foreach (var window in windows) if (IsDialog(window)) modals.Add(window);
        if (modals.Count == 0) return null;
        if (modals.Count != 1 || modals[0].Class != "#32770" || modals[0].Caption != "SOLIDWORKS" || !modals[0].Enabled)
            throw new InvalidOperationException("Unknown/ambiguous startup alert; no acknowledgment.");
        var modal = modals[0]; int buttons = 0; Window answer = null;
        if (modal.Toolbar == null || !modal.Toolbar.Complete || modal.Toolbar.ExactTexts != 1 || modal.Toolbar.EnabledOkButtons != 1 || modal.Toolbar.NativeButton == 0)
            throw new InvalidOperationException("Exact virtual toolbar text / unique enabled UIA OK unproved.");
        foreach (var child in modal.Children)
        {
            if (child.Class == "Button" && child.Visible) { buttons++; if (child.Id == 0 && child.Caption == "OK" && child.Enabled && child.Handle == modal.Toolbar.NativeButton) answer = child; }
        }
        if (buttons != 1 || answer == null || answer.Pid != modal.Pid || answer.Thread != modal.Thread || answer.Root != modal.Handle)
            throw new InvalidOperationException("Exact toolbar warning / unique owned OK unproved.");
        return answer;
    }
    public static void Acknowledge(Window[] original, string claimPath, DateTime deadline, Action requirePrivate)
    {
        requirePrivate(); var answer = KnownToolbarButton(original);
        if (answer == null) throw new InvalidOperationException("Toolbar warning absent; no click.");
        requirePrivate(); var fresh = KnownToolbarButton(Snapshot((int)answer.Pid, deadline));
        if (fresh == null || fresh.Handle != answer.Handle || fresh.Root != answer.Root || fresh.Thread != answer.Thread || DateTime.UtcNow >= deadline)
            throw new InvalidOperationException("Toolbar warning changed/deadline expired before delivery.");
        string json = "{\"State\":\"ONE_ACKNOWLEDGMENT_CLAIMED\",\"Pid\":" + answer.Pid + ",\"Modal\":" + answer.Root + ",\"Button\":" + answer.Handle + ",\"Invocations\":1,\"Utc\":\"" + DateTime.UtcNow.ToString("o") + "\"}";
        using (var file = new FileStream(claimPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) { byte[] bytes = new UTF8Encoding(false).GetBytes(json); file.Write(bytes, 0, bytes.Length); file.Flush(true); }
        requirePrivate();
        if (DateTime.UtcNow >= deadline) throw new InvalidOperationException("Startup deadline expired after claim; no click.");
        UIntPtr result; uint remaining = (uint)Math.Min(2000, Math.Max(1, (deadline - DateTime.UtcNow).TotalMilliseconds));
        if (SendMessageTimeout(new IntPtr(answer.Handle), 0xF5, IntPtr.Zero, IntPtr.Zero, 2, remaining, out result) == IntPtr.Zero)
            throw new InvalidOperationException("Toolbar acknowledgment uncertain; never retry.");
    }
}
