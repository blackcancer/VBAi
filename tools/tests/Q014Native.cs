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
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr first, string text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")] private static extern IntPtr ReadMessage(IntPtr window, uint message, IntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("ole32.dll")] private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx context);
    public sealed class Window { public long Handle; public uint Pid, Thread; public string Class, Caption; public bool Visible; }
    private static Window Observe(IntPtr handle)
    {
        uint pid; uint thread = GetWindowThreadProcessId(handle, out pid);
        var cls = new StringBuilder(256); var caption = new StringBuilder(2048);
        GetClassName(handle, cls, cls.Capacity); GetWindowText(handle, caption, caption.Capacity);
        return new Window { Handle = handle.ToInt64(), Pid = pid, Thread = thread, Class = cls.ToString(), Caption = caption.ToString(), Visible = IsWindowVisible(handle) };
    }
    public static Window[] Windows(int pid)
    {
        var result = new List<Window>(); SetLastError(0);
        bool complete = EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()), (h, s) => { var w = Observe(h); if (w.Pid == pid) result.Add(w); return true; }, IntPtr.Zero);
        int error = Marshal.GetLastWin32Error();
        if (!complete && (result.Count != 0 || error != 0)) throw new Win32Exception(error, "Desktop inventory incomplete");
        return result.ToArray();
    }
    public static long[] FileControls(long hwnd, int pid)
    {
        var root = new IntPtr(hwnd); var w = Observe(root);
        if (w.Pid != pid || w.Class != "#32770" || !w.Visible) throw new InvalidOperationException("Exact owned file dialog required");
        var edits = new List<long>(); var buttons = new List<long>();
        EnumChildWindows(root, (h, s) => {
            var c = Observe(h); int id = GetDlgCtrlID(h); var parent = Observe(GetParent(h));
            if (c.Pid == pid && c.Thread == w.Thread && c.Visible && IsWindowEnabled(h)) {
                if (c.Class == "Edit" && parent.Class == "ComboBox" && (id == 1001 || id == 1148)) edits.Add(c.Handle);
                if (c.Class == "Button" && id == 1) buttons.Add(c.Handle);
            } return true;
        }, IntPtr.Zero);
        if (edits.Count != 1 || buttons.Count != 1) throw new InvalidOperationException("Filename/button controls unavailable or ambiguous");
        return new[] { edits[0], buttons[0] };
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
