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

        /// <summary>Modeless editor surface reparented into the VBE's MDI client.</summary>
        private readonly ModernEditorWindow editor;

        /// <summary>VBE automation object used to identify the active native document without activating it.</summary>
        private readonly object vbe;

        /// <summary>Whether the active MDI document is a native designer or Object Browser that must remain visible.</summary>
        private bool nativeDocument;

        /// <summary>Native MDI client handle that owns document surfaces for this VBE instance.</summary>
        private readonly IntPtr workspace;

        /// <summary>UI-thread timer that refreshes visibility and bounds while the VBE workspace changes.</summary>
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

        /// <summary>Finds the VBE MDI client, reparents the editor into it, and starts workspace tracking.</summary>
        /// <param name="vbe">VBE automation object whose main window owns the document workspace.</param>
        /// <param name="editor">Editor form to host as a child surface of that workspace.</param>
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
            // Explicit native inspections and test dispatches retain visibility, ordering and focus.
            if (VbeDebugInspection.IsActive) { ResizeBounds(); return; }
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
                if (nativeDocument) { editor.Hide(); ResizeBounds(); return; }
                if (!editor.Visible) editor.Show();
                // A newly opened CodePane can raise itself above the child editor.
                // Reorder only for native code, preserving designers and docked panes.
                if (type == 0 && GetWindow(workspace, 5) != editor.Handle)
                    SetWindowPos(editor.Handle, IntPtr.Zero, 0, 0, 0, 0, 0x0013);
            }
            catch (COMException) { ResizeBounds(); return; }
            ResizeBounds();
        }

        /// <summary>Updates workspace geometry without changing visibility, child ordering or focus.</summary>
        private void ResizeBounds()
        {
            if (GetClientRect(workspace, out Rect r) &&
                editor.Bounds != new System.Drawing.Rectangle(0, 0, r.Right, r.Bottom))
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

        /// <summary>Disposes  for editor workspace host.</summary>
        public void Dispose() { timer.Stop(); timer.Dispose(); }

        /// <summary>Describes the edges of the native editor workspace rectangle.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {

            /// <summary>Native client-coordinate edges returned by GetClientRect.</summary>
            public int Left, Top, Right, Bottom;
        }

        /// <summary>Callback signature used while enumerating child HWNDs.</summary>
        /// <param name="handle">Child window currently visited.</param>
        /// <param name="parameter">Caller context forwarded by EnumChildWindows.</param>
        /// <returns><see langword="true"/> to continue enumeration; <see langword="false"/> to stop.</returns>
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);

        /// <summary>Sets the calling thread's last-error value before a Win32 call whose failure is checked.</summary>
        /// <param name="error">Unsigned Win32 error value to store.</param>
        [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);

        /// <summary>Changes a child window's native parent and preserves its previous parent handle.</summary>
        /// <param name="child">Window to reparent.</param>
        /// <param name="parent">New workspace parent.</param>
        /// <returns>Previous parent handle, or zero when no previous parent exists or the call fails.</returns>
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);

        /// <summary>Enumerates descendant HWNDs of a native parent until the callback stops the walk.</summary>
        /// <param name="parent">Window whose descendants are visited.</param>
        /// <param name="callback">Managed callback invoked for each child.</param>
        /// <param name="parameter">Opaque callback context passed through to <paramref name="callback"/>.</param>
        /// <returns><see langword="true"/> when enumeration succeeds.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);

        /// <summary>Copies a window's Unicode class name into a caller-provided buffer.</summary>
        /// <param name="window">Window whose class is queried.</param>
        /// <param name="name">Buffer receiving the class name.</param>
        /// <param name="count">Buffer capacity in characters, including space for the terminator.</param>
        /// <returns>Number of copied characters, excluding the terminator.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);

        /// <summary>Reads the client-area rectangle in the target window's client coordinates.</summary>
        /// <param name="window">Window whose client area is queried.</param>
        /// <param name="rectangle">Receives the left, top, right, and bottom coordinates.</param>
        /// <returns><see langword="true"/> when the rectangle was retrieved.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);

        /// <summary>Tests whether an HWND still identifies an existing window.</summary>
        /// <param name="window">Handle to test.</param>
        /// <returns><see langword="true"/> while the handle identifies a window.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

        /// <summary>Positions a child surface and applies the requested resize, z-order, and activation flags.</summary>
        /// <param name="window">Window to position.</param>
        /// <param name="after">Sibling HWND controlling z-order, or zero when z-order is unchanged.</param>
        /// <param name="x">New horizontal client-coordinate position when not suppressed by flags.</param>
        /// <param name="y">New vertical client-coordinate position when not suppressed by flags.</param>
        /// <param name="width">New width in pixels when not suppressed by flags.</param>
        /// <param name="height">New height in pixels when not suppressed by flags.</param>
        /// <param name="flags">SetWindowPos flags controlling which values are applied.</param>
        /// <returns><see langword="true"/> when the native positioning request succeeds.</returns>
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
