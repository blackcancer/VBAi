using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CodexVBE
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
    }
}
