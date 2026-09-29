using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace VBAi
{
    // A child surface in the VBE document area, independent of any individual CodePane.
    /// <summary>Hosts the modern editor inside the VBE document workspace and follows the active native document window.</summary>
    internal sealed class EditorWorkspaceHost : IDisposable
    {
        /// <summary>Stores the editor used by EditorWorkspaceHost.</summary>
        private readonly ModernEditorWindow editor;
        /// <summary>Stores the vbe used by EditorWorkspaceHost.</summary>
        private readonly object vbe;
        /// <summary>Stores the native document used by EditorWorkspaceHost.</summary>
        private bool nativeDocument;
        /// <summary>Stores the workspace used by EditorWorkspaceHost.</summary>
        private readonly IntPtr workspace;
        /// <summary>Stores the timer used by EditorWorkspaceHost.</summary>
        private readonly Timer timer = new Timer { Interval = 200 };
        /// <summary>Caches the type of the native document while a docked tool owns focus.</summary>
        private IntPtr documentWindow;
        /// <summary>Detects caption changes or handle reuse before using the cached document type.</summary>
        private string documentCaption;
        /// <summary>Stores the last resolved document type, or minus one when unknown.</summary>
        private int documentType = -1;
        /// <summary>Bounds repeated COM enumeration when a new native document is not yet exposed.</summary>
        private readonly System.Diagnostics.Stopwatch documentRetry = System.Diagnostics.Stopwatch.StartNew();
        /// <summary>Changes the native parent; defaults to the Windows API.</summary>
        internal static Func<IntPtr, IntPtr, IntPtr> ChangeParent = SetParent;
        /// <summary>Initializes a EditorWorkspaceHost instance with the supplied state.</summary>
        /// <param name="vbe">The vbe used by this operation.</param>
        /// <param name="editor">The editor used by this operation.</param>
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
        /// <summary>Shows the editor as a child surface in the VBE document workspace.</summary>
        internal void Show() { nativeDocument = false; Resize(); editor.Show(); editor.BringToFront(); editor.Browser?.Focus(); }
        /// <summary>Matches the hosted editor bounds to the active VBE document workspace and hides it for native designers.</summary>
        private void Resize()
        {
            if (editor.IsDisposed || !IsWindow(workspace)) { timer.Stop(); return; }
            try
            {
                dynamic active = ((dynamic)vbe).ActiveWindow;
                int type = active == null ? -1 : (int)active.Type;
                // ActiveWindow follows keyboard focus, which can already be back in the
                // chat when a newly opened native document raises itself above Monaco.
                if (type != 0 && type != 1 && type != 2) type = ActiveDocumentType();
                // Leave native designers and the Object Browser usable in the document area.
                if (type == 1 || type == 2) nativeDocument = true;
                else if (type == 0) nativeDocument = false;
                if (nativeDocument) { editor.Hide(); return; }
                if (!editor.Visible) editor.Show();
                // A newly opened CodePane can raise itself above the child editor.
                // Reorder only for native code, preserving designers and docked panes.
                if (type == 0 && GetWindow(workspace, 5) != editor.Handle)
                    SetWindowPos(editor.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013);
            }
            catch (COMException) { return; }
            if (GetClientRect(workspace, out Rect r))
                // Resize only: a timer must never raise Monaco above native panes
                // while their containing frame is being docked or detached.
                SetWindowPos(editor.Handle, IntPtr.Zero, 0, 0, r.Right, r.Bottom, 0x0014);
        }
        /// <summary>Identifies the actual MDI document without activating it or moving keyboard focus.</summary>
        /// <returns>The matching VBIDE document type, or minus one when identity is unavailable.</returns>
        private int ActiveDocumentType()
        {
            // WM_MDIGETACTIVE queries the MDI client independently of the focused tool window.
            IntPtr child = SendMessage(workspace, 0x0229, IntPtr.Zero, IntPtr.Zero);
            if (child == IntPtr.Zero || child == editor.Handle || GetParent(child) != workspace) return -1;
            var caption = new StringBuilder(512);
            GetWindowText(child, caption, caption.Capacity);
            string title = caption.ToString();
            if (child == documentWindow && title == documentCaption &&
                (documentType >= 0 || documentRetry.ElapsedMilliseconds < 1000)) return documentType;
            int matched = -1;
            int captionMatches = 0;
            foreach (dynamic window in ((dynamic)vbe).Windows)
            {
                int type = (int)window.Type;
                if (type != 0 && type != 1 && type != 2) continue;
                IntPtr handle = new IntPtr(Convert.ToInt64(window.HWnd));
                if (handle == child) { matched = type; captionMatches = 1; break; }
                // Some VBE code windows report HWnd=0. Require an exact, unique caption
                // for the already verified native MDI child, never the last active pane.
                if (handle == IntPtr.Zero && title.Length != 0 && title == (string)window.Caption)
                { matched = type; captionMatches++; }
            }
            documentWindow = child; documentCaption = title;
            documentType = captionMatches == 1 ? matched : -1;
            documentRetry.Restart();
            return documentType;
        }
        /// <summary>Performs the dispose operation for EditorWorkspaceHost.</summary>
        public void Dispose() { timer.Stop(); timer.Dispose(); }
        /// <summary>Describes the edges of the native editor workspace rectangle.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct Rect { /// <summary>Stores the left,top,right,bottom used by Rect.</summary>
public int Left, Top, Right, Bottom; }
        /// <summary>Defines the enum window callback.</summary>
        /// <param name="handle">The handle used by this operation.</param>
        /// <param name="parameter">The parameter used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);
        /// <summary>Performs the set last error operation for EditorWorkspaceHost.</summary>
        /// <param name="error">The error used by this operation.</param>
        [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);
        /// <summary>Performs the set parent operation for EditorWorkspaceHost.</summary>
        /// <param name="child">The child used by this operation.</param>
        /// <param name="parent">The parent used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        /// <summary>Performs the enum child windows operation for EditorWorkspaceHost.</summary>
        /// <param name="parent">The parent used by this operation.</param>
        /// <param name="callback">The callback used by this operation.</param>
        /// <param name="parameter">The parameter used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
        /// <summary>Performs the get class name operation for EditorWorkspaceHost.</summary>
        /// <param name="window">The window used by this operation.</param>
        /// <param name="name">The name used by this operation.</param>
        /// <param name="count">The count used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
        /// <summary>Performs the get client rect operation for EditorWorkspaceHost.</summary>
        /// <param name="window">The window used by this operation.</param>
        /// <param name="rectangle">The rectangle used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);
        /// <summary>Performs the is window operation for EditorWorkspaceHost.</summary>
        /// <param name="window">The window used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        /// <summary>Performs the set window pos operation for EditorWorkspaceHost.</summary>
        /// <param name="window">The window used by this operation.</param>
        /// <param name="after">The after used by this operation.</param>
        /// <param name="x">The x used by this operation.</param>
        /// <param name="y">The y used by this operation.</param>
        /// <param name="width">The width used by this operation.</param>
        /// <param name="height">The height used by this operation.</param>
        /// <param name="flags">The flags used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
        /// <summary>Reads a related native workspace window.</summary>
        /// <param name="window">The owned workspace window.</param>
        /// <param name="command">The native relationship selector.</param>
        /// <returns>The requested related window, if present.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        /// <summary>Reads the native parent of an owned document.</summary>
        /// <param name="window">The document window.</param>
        /// <returns>The native parent.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        /// <summary>Reads a native document caption without changing focus.</summary>
        /// <param name="window">The document window.</param>
        /// <param name="caption">The caption buffer.</param>
        /// <param name="capacity">The buffer capacity.</param>
        /// <returns>The number of copied characters.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder caption, int capacity);
        /// <summary>Queries the active child of the owned MDI workspace.</summary>
        /// <param name="window">The workspace window.</param>
        /// <param name="message">The query message.</param>
        /// <param name="wParam">The first native parameter.</param>
        /// <param name="lParam">The second native parameter.</param>
        /// <returns>The native query result.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
    }
}
