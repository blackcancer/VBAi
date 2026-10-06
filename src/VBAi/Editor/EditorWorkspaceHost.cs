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

        /// <summary>Maintains the editor state for editor workspace host.</summary>
        private readonly ModernEditorWindow editor;

        /// <summary>Maintains the vbe state for editor workspace host.</summary>
        private readonly object vbe;

        /// <summary>Maintains the native document state for editor workspace host.</summary>
        private bool nativeDocument;

        /// <summary>Maintains the workspace state for editor workspace host.</summary>
        private readonly IntPtr workspace;

        /// <summary>Maintains the timer state for editor workspace host.</summary>
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
        /// <param name="vbe">object that supplies the vbe for this operation.</param>
        /// <param name="editor">modern editor window that supplies the editor for this operation.</param>
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
        [StructLayout(LayoutKind.Sequential)] private struct Rect {

/// <summary>Maintains the left and top and right and bottom state for rect.</summary>
public int Left, Top, Right, Bottom; }

        /// <summary>Defines the enum window callback.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on editor workspace host.</returns>
        private delegate bool EnumWindow(IntPtr handle, IntPtr parameter);

        /// <summary>Sets last error for editor workspace host.</summary>
        /// <param name="error">uint that supplies the error for this operation.</param>
        [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);

        /// <summary>Sets parent for editor workspace host.</summary>
        /// <param name="child">Native handle that supplies the child for this operation.</param>
        /// <param name="parent">Native handle that supplies the parent for this operation.</param>
        /// <returns>int ptr produced by the operation for set parent on editor workspace host.</returns>
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetParent(IntPtr child, IntPtr parent);

        /// <summary>Handles enum child windows for editor workspace host.</summary>
        /// <param name="parent">Native handle that supplies the parent for this operation.</param>
        /// <param name="callback">enum window that supplies the callback for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <returns>Boolean indicating the result of the check for enum child windows on editor workspace host.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);

        /// <summary>Returns class name for editor workspace host.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="name">string builder that supplies the name for this operation.</param>
        /// <param name="count">int that supplies the count for this operation.</param>
        /// <returns>int produced by the operation for get class name on editor workspace host.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);

        /// <summary>Returns client rect for editor workspace host.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="rectangle">rect that supplies the rectangle for this operation.</param>
        /// <returns>Boolean indicating the result of the check for get client rect on editor workspace host.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);

        /// <summary>Determines whether window for editor workspace host.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window on editor workspace host.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

        /// <summary>Sets window pos for editor workspace host.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="after">Native handle that supplies the after for this operation.</param>
        /// <param name="x">int that supplies the x for this operation.</param>
        /// <param name="y">int that supplies the y for this operation.</param>
        /// <param name="width">int that supplies the width for this operation.</param>
        /// <param name="height">int that supplies the height for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <returns>Boolean indicating the result of the check for set window pos on editor workspace host.</returns>
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
