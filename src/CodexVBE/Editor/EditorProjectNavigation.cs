using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Routes code-module double clicks in the native project tree to Monaco.</summary>
    internal sealed class EditorProjectNavigation : IDisposable
    {
        private readonly object vbe;
        private readonly Control dispatcher;
        private readonly Action<IEditorModule> open;
        private readonly List<TreeHook> hooks = new List<TreeHook>();
        private readonly Timer timer = new Timer { Interval = 1000 };
        internal EditorProjectNavigation(object vbe, Control dispatcher, Action<IEditorModule> open)
        {
            this.vbe = vbe; this.dispatcher = dispatcher; this.open = open;
            timer.Tick += (s, e) => Refresh(); Refresh(); timer.Start();
        }
        private void Refresh()
        {
            hooks.RemoveAll(h => h.Handle == IntPtr.Zero);
            try
            {
                foreach (dynamic window in ((dynamic)vbe).Windows)
                {
                    if ((int)window.Type != 6) continue; // vbext_wt_ProjectWindow
                    foreach (IntPtr handle in FindProjectTrees(System.Diagnostics.Process.GetCurrentProcess().Id, (string)window.Caption))
                    {
                        uint process;
                        if (GetWindowThreadProcessId(handle, out process) == GetCurrentThreadId() &&
                            !hooks.Any(h => h.Handle == handle)) hooks.Add(new TreeHook(handle, Route));
                    }
                }
            }
            catch (Exception error) { LoadLog.Write("Monaco project navigation: " + error.GetType().Name); }
        }
        private bool Route()
        {
            try
            {
                dynamic component = ((dynamic)vbe).SelectedVBComponent;
                if (component == null || !IsCodeComponent((int)component.Type)) return false;
                object project = component.Collection.Parent;
                var module = new EditorVbeModule(vbe, project, component);
                dispatcher.BeginInvoke(new Action(() => open(module)));
                return true;
            }
            catch (Exception error) { LoadLog.Write("Monaco navigation refused: " + error.GetType().Name); return false; }
        }
        internal static bool IsCodeComponent(int type) => type == 1 || type == 2 || type == 100;
        internal static IntPtr[] FindProjectTrees(int processId, string caption)
        {
            var result = new List<IntPtr>();
            EnumWindows((root, ignored) =>
            {
                GetWindowThreadProcessId(root, out uint pid);
                var kind = new StringBuilder(128); GetClassName(root, kind, kind.Capacity);
                if (pid != processId || kind.ToString() != "wndclass_desked_gsk") return true;
                EnumChildWindows(root, (pane, unused) =>
                {
                    var title = new StringBuilder(512); GetWindowText(pane, title, title.Capacity);
                    if (title.ToString() != caption) return true;
                    EnumChildWindows(pane, (child, value) =>
                    {
                        var name = new StringBuilder(128); GetClassName(child, name, name.Capacity);
                        if (name.ToString() == "SysTreeView32" && !result.Contains(child)) result.Add(child);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return result.ToArray();
        }
        public void Dispose()
        { timer.Stop(); timer.Dispose(); foreach (var hook in hooks) if (hook.Handle != IntPtr.Zero) hook.ReleaseHandle(); hooks.Clear(); }
        private sealed class TreeHook : NativeWindow
        {
            private readonly Func<bool> route;
            internal TreeHook(IntPtr handle, Func<bool> route) { this.route = route; AssignHandle(handle); }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x0203) // WM_LBUTTONDBLCLK, delivered only to this native project tree
                {
                    long point = message.LParam.ToInt64();
                    var hit = new HitTest { X = unchecked((short)point), Y = unchecked((short)(point >> 16)) };
                    SendMessage(Handle, 0x1111, IntPtr.Zero, ref hit); // TVM_HITTEST
                    // Leave expansion glyphs, empty space and noncomponent nodes to the tree.
                    if ((hit.Flags & 0x46) != 0 && hit.Item != IntPtr.Zero &&
                        hit.Item == SendMessage(Handle, 0x110A, new IntPtr(9), IntPtr.Zero) && route())
                    {
                        // Keep VBE's backing code window alive. The queued Monaco activation
                        // runs after the native double-click has finished opening it.
                        base.WndProc(ref message); return;
                    }
                }
                base.WndProc(ref message);
            }
        }
        [StructLayout(LayoutKind.Sequential)] private struct HitTest { public int X, Y; public uint Flags; public IntPtr Item; }
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindow callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, StringBuilder value, int count);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder value, int count);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, ref HitTest hit);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
    }
}
