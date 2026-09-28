using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CodexVBE
{
    // A child surface in the VBE document area, independent of any individual CodePane.
    internal sealed class EditorWorkspaceHost : IDisposable
    {
        private readonly ModernEditorWindow editor;
        private readonly object vbe;
        private bool nativeDocument;
        private readonly IntPtr workspace;
        private readonly Timer timer = new Timer { Interval = 200 };
        /// <summary>Changes the native parent; defaults to the Windows API.</summary>
        internal static Func<IntPtr, IntPtr, IntPtr> ChangeParent = SetParent;
        internal EditorWorkspaceHost(object vbe, ModernEditorWindow editor)
        {
            this.editor = editor; this.vbe = vbe;
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(new IntPtr((int)((dynamic)vbe).MainWindow.HWnd), (child, unused) =>
            {
                var name = new StringBuilder(128); GetClassName(child, name, name.Capacity);
                if (name.ToString().IndexOf("MDIClient", StringComparison.OrdinalIgnoreCase) >= 0) { found = child; return false; }
                return true;
            }, IntPtr.Zero);
            if (found == IntPtr.Zero) throw new InvalidOperationException("The VBE document workspace is unavailable.");
            workspace = found;
            editor.WorkspaceHosted = true;
            editor.TopLevel = false;
            editor.MinimumSize = System.Drawing.Size.Empty;
            SetLastError(0);
            if (ChangeParent(editor.Handle, workspace) == IntPtr.Zero && Marshal.GetLastWin32Error() != 0)
                throw new Win32Exception(Marshal.GetLastWin32Error());
            timer.Tick += (s, e) => Resize(); timer.Start();
        }
        internal void Show() { nativeDocument = false; Resize(); editor.Show(); editor.BringToFront(); editor.Browser?.Focus(); }
        private void Resize()
        {
            if (editor.IsDisposed || !IsWindow(workspace)) { timer.Stop(); return; }
            try
            {
                dynamic active = ((dynamic)vbe).ActiveWindow;
                int type = active == null ? -1 : (int)active.Type;
                // Leave native designers and the Object Browser usable in the document area.
                if (type == 1 || type == 2) nativeDocument = true;
                else if (type == 0) nativeDocument = false;
                if (nativeDocument) { editor.Hide(); return; }
                if (!editor.Visible) editor.Show();
            }
            catch (COMException) { return; }
            if (GetClientRect(workspace, out Rect r))
                // Resize only: a timer must never raise Monaco above native panes
                // while their containing frame is being docked or detached.
                SetWindowPos(editor.Handle, IntPtr.Zero, 0, 0, r.Right, r.Bottom, 0x0014);
        }
        public void Dispose() { timer.Stop(); timer.Dispose(); }
        [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);
        [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
    }
}
