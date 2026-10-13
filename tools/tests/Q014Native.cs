using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

// Qualification-only window addressing and exact ROT lookup. No global input or activation.
public static class Q014Native
{
    private delegate bool Visitor(IntPtr window, IntPtr state);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumDesktopWindows(IntPtr desktop, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr root, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, IntPtr first, IntPtr second);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr first, string text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")] private static extern IntPtr ReadMessage(IntPtr window, uint message, IntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("ole32.dll")] private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx context);
    public sealed class Window { public long Handle; public uint Pid, Thread; public string Class, Caption; public bool Visible, Enabled; }
    private static Window Observe(IntPtr handle)
    {
        uint pid; uint thread = GetWindowThreadProcessId(handle, out pid);
        var cls = new StringBuilder(256); var caption = new StringBuilder(2048);
        GetClassName(handle, cls, cls.Capacity); GetWindowText(handle, caption, caption.Capacity);
        return new Window { Handle = handle.ToInt64(), Pid = pid, Thread = thread, Class = cls.ToString(), Caption = caption.ToString(), Visible = IsWindowVisible(handle), Enabled = IsWindowEnabled(handle) };
    }

    /// <summary>Requests the system-menu Close command of one owned VBE root once, without synthesizing keyboard input.</summary>
    /// <remarks>Message delivery is not completion. The controller must separately observe any named save prompt and the hidden root.</remarks>
    public sealed class VbeCloseAttempt
    {
        private readonly Window expected;
        private readonly Func<Window[]> inventory;
        private readonly Action requireContext;
        private readonly Func<long, bool> postClose;
        private bool entered;

        /// <summary>Creates a close attempt bound to the previously captured root and current desktop guard.</summary>
        public VbeCloseAttempt(Window root, Action requireContext)
            : this(root, () => new Window[0], requireContext,
                  handle => DeliverVbeSystemClose(handle, (target, message, command) =>
                      PostMessage(new IntPtr(target), message, new IntPtr(command), IntPtr.Zero)))
        {
            inventory = () => Windows(checked((int)expected.Pid));
        }

        /// <summary>Injects native observation and delivery boundaries for complete rejection and one-shot tests.</summary>
        public VbeCloseAttempt(Window root, Func<Window[]> inventory, Action requireContext, Func<long, bool> postClose)
        {
            if (root == null || root.Handle <= 0 || root.Pid == 0 || root.Pid > int.MaxValue || root.Thread == 0 ||
                root.Class != "wndclass_desked_gsk" || string.IsNullOrWhiteSpace(root.Caption) || !root.Visible || !root.Enabled)
                throw new ArgumentException("One captured visible, enabled VBE root is required.", "root");
            if (inventory == null) throw new ArgumentNullException("inventory");
            if (requireContext == null) throw new ArgumentNullException("requireContext");
            if (postClose == null) throw new ArgumentNullException("postClose");
            expected = new Window { Handle = root.Handle, Pid = root.Pid, Thread = root.Thread,
                Class = root.Class, Caption = root.Caption, Visible = root.Visible, Enabled = root.Enabled };
            this.inventory = inventory;
            this.requireContext = requireContext;
            this.postClose = postClose;
        }

        /// <summary>Revalidates desktop membership and exact root identity twice, then delivers at most one system-menu Close command.</summary>
        public void Request()
        {
            if (entered) throw new InvalidOperationException("The VBE close attempt has already entered; no retry.");
            entered = true;
            requireContext(); RequireInventory(inventory());
            requireContext(); RequireInventory(inventory());
            requireContext();
            if (!postClose(expected.Handle))
                throw new InvalidOperationException("VBE close delivery failed; completion is unproven and no retry is permitted.");
        }

        /// <summary>Rejects incomplete ownership, reused roots and pre-existing modal dialogs before message delivery.</summary>
        private void RequireInventory(Window[] windows)
        {
            if (windows == null || windows.Length == 0 || windows.Length > 8192)
                throw new InvalidOperationException("Complete bounded owned window inventory required before VBE close.");
            var seen = new HashSet<long>(); Window root = null;
            foreach (var window in windows)
            {
                if (window == null || window.Handle <= 0 || window.Pid != expected.Pid || window.Thread == 0 ||
                    string.IsNullOrWhiteSpace(window.Class) || !seen.Add(window.Handle))
                    throw new InvalidOperationException("Window inventory ownership or uniqueness changed; no close.");
                if (window.Class == "#32770" && window.Visible)
                    throw new InvalidOperationException("A pre-existing modal prevents ordinary VBE close.");
                if (window.Handle == expected.Handle) root = window;
            }
            if (root == null || root.Thread != expected.Thread || root.Class != expected.Class || root.Caption != expected.Caption ||
                !root.Visible || !root.Enabled)
                throw new InvalidOperationException("The captured VBE root changed before close; no delivery.");
        }
    }
    /// <summary>Delivers WM_SYSCOMMAND/SC_CLOSE so the VBE can handle its system-menu action before default destruction.</summary>
    /// <param name="handle">The positive native handle already admitted by the close-attempt ownership checks.</param>
    /// <param name="post">The single native delivery boundary; message and command remain independently observable.</param>
    /// <returns>Whether posting succeeded; this does not prove VBE closure or continued host lifetime.</returns>
    public static bool DeliverVbeSystemClose(long handle, Func<long, uint, long, bool> post)
    {
        if (handle <= 0) throw new ArgumentOutOfRangeException("handle");
        if (post == null) throw new ArgumentNullException("post");
        return post(handle, 0x0112, 0xF060);
    }

    public static Window[] Windows(int pid)
    {
        var result = new List<Window>(); SetLastError(0);
        bool complete = EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()), (h, s) => { var w = Observe(h); if (w.Pid == pid) result.Add(w); return true; }, IntPtr.Zero);
        int error = Marshal.GetLastWin32Error();
        if (!complete && (result.Count != 0 || error != 0)) throw new Win32Exception(error, "Desktop inventory incomplete");
        return result.ToArray();
    }

    /// <summary>Recognizes only the owned welcome screen when its UI Automation provider explicitly reports a modeless window.</summary>
    /// <remarks>Window class alone does not prove modality. Unknown providers, changed inventories and disabled frames remain blocking.</remarks>
    public static bool IsModelessWelcome(Window dialog, Window frame, Window[] windows, Func<long, bool?> readModal)
    {
        if (dialog == null || frame == null || windows == null || windows.Length == 0 || windows.Length > 8192 || readModal == null ||
            dialog.Handle <= 0 || frame.Handle <= 0 || dialog.Handle == frame.Handle || dialog.Pid == 0 || dialog.Thread == 0 ||
            dialog.Pid != frame.Pid || dialog.Thread != frame.Thread || dialog.Class != "#32770" ||
            !string.IsNullOrEmpty(dialog.Caption) || !dialog.Visible || !dialog.Enabled ||
            frame.Class == null || !frame.Class.StartsWith("Afx:", StringComparison.Ordinal) || !frame.Visible || !frame.Enabled)
            return false;
        var handles = new HashSet<long>(); int dialogs = 0, frames = 0, welcomes = 0, exactDialog = 0, exactFrame = 0;
        foreach (var window in windows)
        {
            if (window == null || window.Handle <= 0 || window.Pid != frame.Pid || window.Thread == 0 ||
                string.IsNullOrWhiteSpace(window.Class) || !handles.Add(window.Handle)) return false;
            if (window.Handle == dialog.Handle)
            {
                if (!SameWindow(window, dialog)) return false;
                exactDialog++;
            }
            if (window.Handle == frame.Handle)
            {
                if (!SameWindow(window, frame)) return false;
                exactFrame++;
            }
            if (!window.Visible) continue;
            if (window.Class == "#32770") dialogs++;
            if (window.Class.StartsWith("Afx:", StringComparison.Ordinal) && window.Caption == frame.Caption) frames++;
            if (window.Class == "swWelcomeScreen")
            {
                if (window.Thread != frame.Thread || !window.Enabled) return false;
                welcomes++;
            }
        }
        if (exactDialog != 1 || exactFrame != 1 || dialogs != 1 || frames != 1 || welcomes != 1) return false;
        // Provider errors propagate: an unavailable modality observation cannot authorize a command.
        return readModal(dialog.Handle) == false;
    }

    private static bool SameWindow(Window first, Window second)
    {
        return first.Handle == second.Handle && first.Pid == second.Pid &&
            first.Thread == second.Thread && first.Class == second.Class && first.Caption == second.Caption &&
            first.Visible == second.Visible && first.Enabled == second.Enabled;
    }

    public static long[] FileControls(long hwnd, int pid)
    {
        var controls = TryFileControls(hwnd, pid);
        if (controls == null) throw new InvalidOperationException("Filename/button controls unavailable or ambiguous");
        return controls;
    }

    private static long[] TryFileControls(long hwnd, int pid)
    {
        var root = new IntPtr(hwnd); var w = Observe(root);
        if (w.Pid != pid || w.Class != "#32770" || !w.Visible) return null;
        var edits = new List<long>(); var buttons = new List<long>();
        EnumChildWindows(root, (h, s) => {
            var c = Observe(h); int id = GetDlgCtrlID(h); var parent = Observe(GetParent(h));
            if (c.Pid == pid && c.Thread == w.Thread && c.Visible && IsWindowEnabled(h)) {
                if (c.Class == "Edit" && parent.Class == "ComboBox" && (id == 1001 || id == 1148)) edits.Add(c.Handle);
                if (c.Class == "Button" && id == 1) buttons.Add(c.Handle);
            } return true;
        }, IntPtr.Zero);
        if (edits.Count != 1 || buttons.Count != 1) return null;
        return new[] { edits[0], buttons[0] };
    }

    /// <summary>Read-only candidate, not an authorization to follow a replacement after delivery.</summary>
    public sealed class FileDialogSnapshot
    {
        public Window Dialog;
        public long Filename, Action;
        public string ActionCaption;
    }

    /// <summary>Reads the one visible native dialog and returns null until its addressed controls are ready.</summary>
    public static FileDialogSnapshot ObserveFileDialog(int pid)
    {
        if (pid <= 0) throw new ArgumentOutOfRangeException("pid");
        var dialog = UniqueFileDialogCandidate(Windows(pid), pid);
        if (dialog == null) return null;
        var controls = TryFileControls(dialog.Handle, pid);
        if (controls == null) return null;
        var caption = Text(controls[1]);
        var after = Observe(new IntPtr(dialog.Handle));
        if (!SameDialog(dialog, after)) return null;
        return new FileDialogSnapshot { Dialog = after, Filename = controls[0], Action = controls[1], ActionCaption = caption };
    }

    /// <summary>Returns no candidate during zero/multiple-dialog transitions; this never permits a native delivery.</summary>
    public static Window UniqueFileDialogCandidate(Window[] windows, int pid)
    {
        if (windows == null) throw new ArgumentNullException("windows");
        if (pid <= 0) throw new ArgumentOutOfRangeException("pid");
        Window candidate = null;
        foreach (var window in windows)
        {
            if (window == null || window.Pid != pid || window.Class != "#32770" || !window.Visible) continue;
            if (candidate != null) return null;
            candidate = window;
        }
        return candidate;
    }

    /// <summary>Admits two matching ready reads before the first mutation; later deliveries must retain this binding.</summary>
    public static bool IsStableFileDialog(FileDialogSnapshot previous, FileDialogSnapshot current, int pid, string action)
    {
        if (pid <= 0) throw new ArgumentOutOfRangeException("pid");
        if (action != "Save" && action != "Open") throw new ArgumentException("Expected Save or Open", "action");
        return ReadySnapshot(previous, pid, action) && ReadySnapshot(current, pid, action) &&
            SameDialog(previous.Dialog, current.Dialog) && previous.Filename == current.Filename &&
            previous.Action == current.Action && previous.ActionCaption == current.ActionCaption;
    }

    private static bool ReadySnapshot(FileDialogSnapshot value, int pid, string action)
    {
        if (value == null || value.Dialog == null || value.Dialog.Handle <= 0 || value.Dialog.Pid != pid ||
            value.Dialog.Thread == 0 || value.Dialog.Class != "#32770" || !value.Dialog.Visible ||
            value.Filename <= 0 || value.Action <= 0 || value.Filename == value.Action ||
            value.Filename == value.Dialog.Handle || value.Action == value.Dialog.Handle || value.ActionCaption == null) return false;
        string caption = value.ActionCaption.Replace("&", "");
        return action == "Save" ? caption == "Save" || caption == "Enregistrer" : caption == "Open" || caption == "Ouvrir";
    }

    private static bool SameDialog(Window first, Window second)
    {
        return first != null && second != null && first.Handle == second.Handle && first.Pid == second.Pid &&
            first.Thread == second.Thread && first.Class == second.Class && first.Caption == second.Caption &&
            first.Visible && second.Visible;
    }
    public static string Text(long hwnd)
    {
        var text = new StringBuilder(4096); UIntPtr result;
        if (ReadMessage(new IntPtr(hwnd), 13, new IntPtr(text.Capacity), text, 2, 3000, out result) == IntPtr.Zero) throw new InvalidOperationException("Text read uncertain");
        return text.ToString();
    }
    public static void Filename(long hwnd, string path)
    {
        UIntPtr result;
        if (SendMessageTimeout(new IntPtr(hwnd), 12, IntPtr.Zero, path, 2, 3000, out result) == IntPtr.Zero || result == UIntPtr.Zero) throw new InvalidOperationException("Filename delivery uncertain");
        if (Text(hwnd) != path) throw new InvalidOperationException("Filename readback differs");
    }
    public static void Click(long hwnd)
    {
        UIntPtr result;
        if (SendMessageTimeout(new IntPtr(hwnd), 0xF5, IntPtr.Zero, null, 2, 10000, out result) == IntPtr.Zero) throw new InvalidOperationException("Button delivery uncertain; never retry");
    }
    public static object ExactRot(string expected)
    {
        IRunningObjectTable table = null; IBindCtx context = null; IEnumMoniker iterator = null;
        try {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out table)); Marshal.ThrowExceptionForHR(CreateBindCtx(0, out context));
            table.EnumRunning(out iterator); var m = new IMoniker[1];
            while (iterator.Next(1, m, IntPtr.Zero) == 0) {
                try { string name; m[0].GetDisplayName(context, null, out name); if (!string.Equals(name, expected, StringComparison.OrdinalIgnoreCase)) continue;
                    object value; table.GetObject(m[0], out value); return value;
                } catch (COMException) { } finally { if (m[0] != null) Marshal.ReleaseComObject(m[0]); }
            }
            return null;
        } finally { if (iterator != null) Marshal.ReleaseComObject(iterator); if (context != null) Marshal.ReleaseComObject(context); if (table != null) Marshal.ReleaseComObject(table); }
    }
}
