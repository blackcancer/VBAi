using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace VBAi
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

        /// <summary>Palette transaction route; receives the VBE, requested apply/restore state, and recovery file path.</summary>
        private readonly Action<object, bool, string> change;

        /// <summary>Failure handler called after a palette transaction throws.</summary>
        private readonly Action<Exception> reportFailure;

        /// <summary>Latest requested apply or restore state.</summary>
        private bool requested;

        /// <summary>Last state successfully applied, or <see langword="null"/> before a successful update.</summary>
        private bool? applied;

        /// <summary>Whether this service has been disposed.</summary>
        private bool disposed;

        /// <summary>Process-wide guard against overlapping modal palette transactions.</summary>
        private static int updateInProgress;

        /// <summary>Creates a deferred palette updater for one visible VBE instance and its version-specific recovery state.</summary>
        /// <param name="vbe">VBE automation object whose native Options dialog owns the palette.</param>
        /// <param name="editor">Main editor HWND used to wait until the host can safely show a modal dialog.</param>
        /// <param name="recoveryPath">Optional durable palette-state file; defaults to a per-user, VBE-version-specific path.</param>
        /// <param name="change">Optional transaction implementation, receiving the VBE, desired state, and recovery path.</param>
        /// <param name="reportFailure">Optional handler for transaction failures.</param>
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
            path = recoveryPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "native-theme", "palette-" + version + ".json");
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

        /// <summary>Shows a localized warning that directs the user to the log for native palette failure details.</summary>
        /// <param name="error">Transaction failure whose base message is displayed.</param>
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

        /// <summary>Serializes a VBE palette transaction, rebases saved originals over manual changes, and verifies values after reopening Options.</summary>
        /// <param name="version">VBE version string used to validate the recovery record.</param>
        /// <param name="enabled"><see langword="true"/> to apply the dark palette; <see langword="false"/> to restore the captured original.</param>
        /// <param name="recoveryPath">Durable recovery file path; its sibling lock file excludes concurrent transactions.</param>
        /// <param name="visit">Reads color rows and optionally returns replacement rows through the owned Options dialog.</param>
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
                    if (!VbeNativePaletteState.Equal(current, state.Original) && !VbeNativePaletteState.Equal(current, state.Applied))
                    {
                        state = state.Rebase(current);
                        string archive = state.SaveReplacing(recoveryPath);
                        LoadLog.Write("Native editor palette reconciled with manual changes; previous recovery archived: " + archive);
                    }
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
