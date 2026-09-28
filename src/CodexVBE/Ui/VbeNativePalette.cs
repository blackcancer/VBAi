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
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        private readonly object vbe;
        private readonly IntPtr editor;
        private readonly System.Windows.Forms.Timer timer;
        private readonly string path;
        private readonly Action<object, bool, string> change;
        private readonly Action<Exception> reportFailure;
        private bool requested;
        private bool? applied;
        private bool disposed;
        private static int updateInProgress;

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

        internal void Request(bool enabled)
        {
            if (disposed) return;
            requested = enabled;
            if (Volatile.Read(ref updateInProgress) == 0 &&
                (applied == enabled || (!enabled && !File.Exists(path)))) { timer.Stop(); return; }
            timer.Start();
        }

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

        private static void ShowFailure(Exception error)
        {
            MessageBox.Show(UiText.Get("Native editor colors could not be updated. See the log for details.") +
                Environment.NewLine + error.GetBaseException().Message, UiText.Get("VBAi settings"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static void Change(object vbe, bool enabled, string recoveryPath)
        {
            string version = Convert.ToString(((dynamic)vbe).Version);
            Change(version, enabled, recoveryPath, update => VbeNativePaletteDialog.Visit(vbe, update));
        }

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

        public void Dispose()
        {
            disposed = true;
            // Never open a modal Options dialog while the host is shutting down.
            timer.Stop();
            timer.Dispose();
        }
    }
}
