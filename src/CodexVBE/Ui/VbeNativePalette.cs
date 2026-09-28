using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Queues palette updates until the host has closed its settings dialog.</summary>
    internal sealed class VbeNativePalette : IDisposable
    {
        /// <summary>Checks the visibility state of a native window.</summary><param name="window">Window handle to inspect.</param><returns>Whether it is visible.</returns>
[DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        /// <summary>Checks whether a native window accepts input.</summary><param name="window">Window handle to inspect.</param><returns>Whether it is enabled.</returns>
[DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        /// <summary>VBE automation object used by palette transactions.</summary>
private readonly object vbe;
        /// <summary>Main editor window that gates when a palette transaction may run.</summary>
private readonly IntPtr editor;
        /// <summary>STA timer that defers updates until the host is ready.</summary>
private readonly System.Windows.Forms.Timer timer;
        /// <summary>Version-specific file that stores the original palette for recovery.</summary>
private readonly string path;
        /// <summary>Stores the change used by VbeNativePalette.</summary>
private readonly Action<object, bool, string> change;
        /// <summary>Stores the report failure used by VbeNativePalette.</summary>
private readonly Action<Exception> reportFailure;
        /// <summary>Latest requested apply or restore state.</summary>
private bool requested;
        /// <summary>Last state successfully applied, or <see langword="null"/> before a successful update.</summary>
private bool? applied;
        /// <summary>Whether this service has been disposed.</summary>
private bool disposed;
        /// <summary>Process-wide guard against overlapping modal palette transactions.</summary>
private static int updateInProgress;

        /// <summary>Initializes a VbeNativePalette instance with the supplied state.</summary>
/// <param name="vbe">The vbe used by this operation.</param>
/// <param name="editor">The editor used by this operation.</param>
/// <param name="recoveryPath">Text containing the recovery path.</param>
/// <param name="change">The change used by this operation.</param>
/// <param name="reportFailure">The report failure used by this operation.</param>
internal VbeNativePalette(object vbe, IntPtr editor, string recoveryPath = null,
            Action<object, bool, string> change = null, Action<Exception> reportFailure = null)
        {
            this.vbe = vbe;
            this.editor = editor;
            this.change = change ?? Change;
            this.reportFailure = reportFailure ?? ShowFailure;
            string version = Convert.ToString(((dynamic)vbe).Version);
            if (string.IsNullOrEmpty(version) || version.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                throw new InvalidOperationException("The VBE version cannot be used for palette recovery.");
            path = recoveryPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "native-theme", "palette-" + version + ".json");
            timer = new System.Windows.Forms.Timer { Interval = 250 };
            timer.Tick += ApplyPending;
        }

        /// <summary>Queues applying or restoring the native editor palette.</summary>
        /// <param name="enabled"><see langword="true"/> to apply the saved dark palette; otherwise restore the original.</param>
internal void Request(bool enabled)
        {
            if (disposed) return;
            requested = enabled;
            if (Volatile.Read(ref updateInProgress) == 0 &&
                (applied == enabled || (!enabled && !File.Exists(path)))) { timer.Stop(); return; }
            timer.Start();
        }

        /// <summary>Runs a queued palette transaction when the editor is visible and no other transaction is active.</summary>
        /// <param name="sender">Timer that raised the tick.</param><param name="args">Event arguments.</param>
private void ApplyPending(object sender, EventArgs args)
        {
            if (disposed || !IsWindowVisible(editor) || !IsWindowEnabled(editor)) return;
            // Options pumps STA messages. A setting change can replace this service
            // while its previous transaction is still on the stack.
            if (Interlocked.CompareExchange(ref updateInProgress, 1, 0) != 0) return;
            timer.Stop();
            bool target = requested;
            try
            {
                change(vbe, target, path);
                applied = target;
                LoadLog.Write("Native editor palette " + (target ? "applied" : "restored") + " and verified.");
            }
            catch (Exception error)
            {
                LoadLog.Write("Native editor palette failed: " + error.ToString());
                reportFailure(error);
            }
            finally
            {
                Interlocked.Exchange(ref updateInProgress, 0);
                if (!disposed && requested != target) timer.Start();
            }
        }

        /// <summary>Performs the show failure operation for VbeNativePalette.</summary>
/// <param name="error">The error used by this operation.</param>
private static void ShowFailure(Exception error)
        {
            MessageBox.Show(UiText.Get("Native editor colors could not be updated. See the log for details.") +
                Environment.NewLine + error.GetBaseException().Message, UiText.Get("VBAi settings"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>Applies or restores colors through the VBE Options dialog and verifies the committed result.</summary>
        /// <param name="vbe">VBE automation object whose Options dialog owns the color settings.</param>
        /// <param name="enabled">Whether to apply the dark palette or restore the captured original palette.</param>
        /// <param name="recoveryPath">Path of the durable recovery state used to preserve original colors.</param>
        /// <exception cref="InvalidOperationException">The dialog interaction or post-commit verification fails.</exception>
internal static void Change(object vbe, bool enabled, string recoveryPath)
        {
            string version = Convert.ToString(((dynamic)vbe).Version);
            Change(version, enabled, recoveryPath, update => VbeNativePaletteDialog.Visit(vbe, update));
        }

        /// <summary>Performs the change operation for VbeNativePalette.</summary>
/// <param name="version">Text containing the version.</param>
/// <param name="enabled">Indicates whether enabled is enabled.</param>
/// <param name="recoveryPath">Text containing the recovery path.</param>
/// <param name="visit">The visit used by this operation.</param>
internal static void Change(string version, bool enabled, string recoveryPath,
            Func<Func<VbeNativePaletteState.ColorRow[], VbeNativePaletteState.ColorRow[]>, VbeNativePaletteState.ColorRow[]> visit)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(recoveryPath));
            using (var transaction = new FileStream(recoveryPath + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                var state = VbeNativePaletteState.Load(recoveryPath, version);
                if (!enabled && state == null) return;
                VbeNativePaletteState.ColorRow[] expected = null;
                visit(current =>
                {
                    if (state == null)
                    {
                        state = new VbeNativePaletteState { VbeVersion = version, Original = current, Applied = VbeNativePaletteState.Dark(current) };
                        state.SaveNew(recoveryPath);
                    }
                    state.RequireUnchanged(current);
                    expected = enabled ? state.Applied : state.Original;
                    return VbeNativePaletteState.Equal(current, expected) ? null : expected;
                });
                // Reopen the dialog: checking the edited controls alone does not prove
                // that OK committed their values to the editor's native settings.
                var actual = visit(current => null);
                if (!VbeNativePaletteState.Equal(actual, expected))
                    throw new InvalidOperationException("The native palette differs after reopening Options; the recovery file has been retained.");
                if (!enabled) File.Delete(recoveryPath);
            }
        }

        /// <summary>Stops the timer without opening a modal Options dialog during shutdown.</summary>
public void Dispose()
        {
            disposed = true;
            // Never open a modal Options dialog while the host is shutting down.
            timer.Stop();
            timer.Dispose();
        }
    }
}
