using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Configuration et pilotage manuel des mises à jour ; aucune requête dans le constructeur Designer.</summary>
    internal sealed partial class UpdateWindow : Form
    {
        /// <summary>Release currently displayed in the window.</summary>
        private UpdateRelease release;
        /// <summary>Verified download path, or null before the selected installer is staged.</summary>
        private string downloaded;
        /// <summary>Whether an asynchronous check, download, or install action is active.</summary>
        private bool busy;
        /// <summary>Cancellation source for feed checks and downloads owned by this window.</summary>
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        /// <summary>Stores the runtime disposed used by UpdateWindow.</summary>
        private bool runtimeDisposed;
        /// <summary>Factory for the release feed.</summary>
        internal Func<UpdateFeed> CreateFeed = () => new UpdateFeed();
        /// <summary>Reads persisted update preferences.</summary>
        internal Func<UpdatePreferences> ReadPreferences = () => UpdateState.Load();
        /// <summary>Persists update preferences.</summary>
        internal Action<UpdatePreferences> StorePreferences = p => UpdateState.Save(p);
        /// <summary>Checks whether the current deployment is installer-managed.</summary>
        internal Func<bool> ManagedInstallation = () => UpdateInstallation.IsManaged(UpdateState.InstallationDirectory);
        /// <summary>Schedules a verified release installer with the separate updater process.</summary>
        internal Action<UpdateRelease, string, bool> Schedule = UpdateCoordinator.Schedule;
        /// <summary>Stores the check release used by UpdateWindow.</summary>
        internal Func<UpdateFeed, UpdateVersion, bool, CancellationToken, Task<UpdateRelease>> CheckRelease = (feed, version, previews, token) => feed.Check(version, previews, null, token);
        /// <summary>Stores the download installer used by UpdateWindow.</summary>
        internal Func<UpdateFeed, UpdateAsset, string, IProgress<int>, CancellationToken, Task<string>> DownloadInstaller = (feed, asset, root, progress, token) => feed.Download(asset, root, progress, token);
        /// <summary>Stores the run background used by UpdateWindow.</summary>
        internal Func<Action, Task> RunBackground = action => Task.Run(action);
        /// <summary>Creates the update settings window without starting a check in Designer mode.</summary>
        public UpdateWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            Icon = VbeWindowIcons.Icon("assistant"); UiText.Apply(this, components);
            version.Text = "VBAi " + UpdateState.ProductVersion;
            notes.RightToLeft = RightToLeft.No;
            version.RightToLeft = description.RightToLeft = RightToLeft.No;
            version.TextAlign = description.TextAlign = UiText.Culture.TextInfo.IsRightToLeft ? ContentAlignment.TopRight : ContentAlignment.TopLeft;
            ApplyAppearance(); UiTheme.Changed += ApplyAppearance;
        }
        /// <summary>Loads preferences and cached release data, then starts an automatic check when configured.</summary>
        /// <param name="e">Shown event data.</param>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            var preferences = ReadPreferences();
            automaticCheck.Checked = preferences.CheckAutomatically;
            automaticDownload.Checked = preferences.DownloadAutomatically;
            automaticInstall.Checked = preferences.InstallAutomatically;
            previews.Checked = preferences.IncludePrereleases;
            automaticInstall.Enabled = ManagedInstallation() && automaticDownload.Checked;
            release = UpdateState.CachedRelease();
            if (release != null)
            {
                version.Text += " → " + release.Version.Text;
                notes.Text = release.body ?? "";
                if (release.Installer?.Hash != null)
                {
                    var path = UpdatePaths.AssetPath(UpdatePaths.Root, release.Installer.Hash, release.Installer.name);
                    if (File.Exists(path)) downloaded = path; // Schedule and the worker both reverify before executing.
                }
            }
            SetBusy(false);
            try
            {
                var job = UpdateInstallJob.Load(UpdatePaths.Root);
                cancelPending.Enabled = job != null && !job.Completed;
                status.Text = UiText.Get(job?.Status ?? (release != null ? "An update is available." : ManagedInstallation() ? "Check for updates to get started." : "An installer-managed deployment is required."));
            }
            catch (Exception) { status.Text = UiText.Get("Unable to read the update state."); }
        }
        /// <summary>Applies the current theme to update actions and the release notes pane.</summary>
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            check.BackColor = UiTheme.HighContrast() ? SystemColors.Highlight : Color.FromArgb(37, 99, 235);
            check.ForeColor = UiTheme.HighContrast() ? SystemColors.HighlightText : Color.White;
        }
        /// <summary>Paints a theme-aware face for an action that is currently disabled.</summary>
        /// <param name="sender">Button being painted.</param><param name="e">Paint graphics and clip data.</param>
        private void PaintDisabled(object sender, PaintEventArgs e)
        {
            var control = (Control)sender;
            if (control.Enabled || !UiTheme.Dark || UiTheme.HighContrast()) return;
            var text = control.ClientRectangle;
            var color = Color.FromArgb(148, 163, 184);
            var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;
            if (control is CheckBox box)
            {
                bool rtl = box.RightToLeft == RightToLeft.Yes;
                var glyph = new Rectangle(rtl ? box.Width - 14 : 1, (box.Height - 12) / 2, 12, 12);
                using (var background = new SolidBrush(box.BackColor)) e.Graphics.FillRectangle(background, glyph);
                using (var pen = new Pen(color))
                {
                    e.Graphics.DrawRectangle(pen, glyph);
                    if (box.Checked) e.Graphics.DrawLines(pen, new[] { new Point(glyph.Left + 2, glyph.Top + 6), new Point(glyph.Left + 5, glyph.Top + 9), new Point(glyph.Left + 10, glyph.Top + 2) });
                }
                text = new Rectangle(rtl ? 0 : 18, 0, Math.Max(0, box.Width - 18), box.Height);
                flags |= rtl ? TextFormatFlags.Right | TextFormatFlags.RightToLeft : TextFormatFlags.Left;
            }
            else flags |= TextFormatFlags.HorizontalCenter;
            TextRenderer.DrawText(e.Graphics, control.Text, control.Font, text, color, control.BackColor, flags);
        }

        /// <summary>Copies checkbox states into the preferences model and persists it.</summary>
        private void SavePreferences()
        {
            var preferences = ReadPreferences();
            preferences.CheckAutomatically = automaticCheck.Checked;
            preferences.DownloadAutomatically = automaticDownload.Checked;
            preferences.InstallAutomatically = ManagedInstallation() && automaticInstall.Checked && automaticDownload.Checked;
            preferences.IncludePrereleases = previews.Checked;
            StorePreferences(preferences);
        }
        /// <summary>Saves automatic checks, downloads, installation, and prerelease preferences.</summary>
        /// <param name="sender">Save button.</param><param name="e">Click event data.</param>
        private void Save_Click(object sender, EventArgs e)
        {
            try { SavePreferences(); status.Text = UiText.Get("Update preferences saved."); }
            catch (Exception) { status.Text = UiText.Get("Unable to save update preferences."); }
        }
        /// <summary>Enables automatic installation only when automatic download is selected.</summary>
        /// <param name="sender">Download preference control.</param><param name="e">Change event data.</param>
        private void DownloadPreferenceChanged(object sender, EventArgs e)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            automaticInstall.Enabled = !busy && ManagedInstallation() && automaticDownload.Checked;
            if (!automaticDownload.Checked) automaticInstall.Checked = false;
        }
        /// <summary>Checks the release feed and displays the newest eligible release.</summary>
        /// <param name="sender">Check button.</param><param name="e">Click event data.</param>
        private async void Check_Click(object sender, EventArgs e)
        {
            if (busy) return;
            SetBusy(true); status.Text = UiText.Get("Checking for updates…");
            try
            {
                SavePreferences();
                using (var feed = CreateFeed()) release = await CheckRelease(feed, UpdateVersion.Parse(UpdateState.ProductVersion), previews.Checked, cancellation.Token);
                if (IsDisposed) return;
                UpdateState.CacheRelease(release);
                downloaded = null;
                version.Text = "VBAi " + UpdateState.ProductVersion + (release == null ? "" : " → " + release.Version.Text);
                notes.Text = (release?.body ?? "").Length > 64000 ? release.body.Substring(0, 64000) : release?.body ?? "";
                status.Text = UiText.Get(release == null ? "VBAi is up to date." : release.Installer?.Hash == null ? "This release has no compatible verified installer." : "An update is available.");
                var p = ReadPreferences(); p.LastCheckUtc = DateTime.UtcNow; StorePreferences(p);
            }
            catch (Exception) { if (!IsDisposed) status.Text = UiText.Get("Unable to check for updates. Check the GitHub account or retry later."); }
            finally { if (!IsDisposed) SetBusy(false); }
        }
        /// <summary>Downloads the selected release installer while reporting progress.</summary>
        /// <param name="sender">Download button.</param><param name="e">Click event data.</param>
        private async void Download_Click(object sender, EventArgs e)
        {
            if (busy || release?.Installer == null) return;
            SetBusy(true); progress.Style = ProgressBarStyle.Continuous; progress.Value = 0; status.Text = UiText.Get("Downloading update…");
            try
            {
                var reporter = new Progress<int>(value => { if (!IsDisposed) progress.Value = value; });
                using (var feed = CreateFeed()) downloaded = await DownloadInstaller(feed, release.Installer, UpdatePaths.Root, reporter, cancellation.Token);
                if (!IsDisposed) status.Text = UiText.Get("Update downloaded and verified.");
            }
            catch (Exception) { if (!IsDisposed) status.Text = UiText.Get("Unable to download or verify the installer."); }
            finally { if (!IsDisposed) SetBusy(false); }
        }
        /// <summary>Schedules the staged installer through the separate updater process.</summary>
        /// <param name="sender">Install button.</param><param name="e">Click event data.</param>
        private async void Install_Click(object sender, EventArgs e)
        {
            if (busy || downloaded == null || release == null) return;
            SetBusy(true);
            try { SavePreferences(); await RunBackground(() => Schedule(release, downloaded, false)); if (!IsDisposed) { status.Text = UiText.Get("Waiting for VBA hosts to close."); cancelPending.Enabled = true; } }
            catch (Exception) { if (!IsDisposed) status.Text = UiText.Get("Unable to schedule the update."); }
            finally { if (!IsDisposed) { SetBusy(false); install.Enabled = false; } }
        }
        /// <summary>Saves the displayed version as skipped and closes the window.</summary>
        /// <param name="sender">Skip button.</param><param name="e">Click event data.</param>
        private void Skip_Click(object sender, EventArgs e)
        {
            if (release == null) return;
            try { var p = ReadPreferences(); p.SkippedVersion = release.Version.Text; StorePreferences(p); status.Text = UiText.Get("This version will be skipped automatically."); }
            catch (Exception) { status.Text = UiText.Get("Unable to save update preferences."); }
        }
        /// <summary>Marks an unfinished installation job as cancelled.</summary>
        /// <param name="sender">Cancel-pending button.</param><param name="e">Click event data.</param>
        private void CancelPending_Click(object sender, EventArgs e)
        {
            try
            {
                using (var gate = new FileStream(Path.Combine(UpdatePaths.Root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    var job = UpdateInstallJob.Load(UpdatePaths.Root);
                    if (job == null || job.Completed) return;
                    job.Completed = true; job.Status = "Update cancelled."; job.Save(UpdatePaths.Root);
                    status.Text = UiText.Get(job.Status); cancelPending.Enabled = false;
                }
            }
            catch (IOException) { status.Text = UiText.Get("Installing update…"); }
            catch (Exception) { status.Text = UiText.Get("Unable to read the update state."); }
        }
        /// <summary>Updates operation state and action availability while work is active.</summary>
        /// <param name="value">Whether an operation is active.</param>
        private void SetBusy(bool value)
        {
            busy = value; check.Enabled = save.Enabled = automaticCheck.Enabled = automaticDownload.Enabled = previews.Enabled = !value;
            automaticInstall.Enabled = !value && ManagedInstallation() && automaticDownload.Checked;
            download.Enabled = !value && release?.Installer?.Hash != null;
            install.Enabled = !value && downloaded != null && ManagedInstallation();
            skip.Enabled = !value && release != null;
            progress.Visible = value; progress.Style = ProgressBarStyle.Marquee;
        }
        /// <summary>Cancels outstanding asynchronous work as the window closes.</summary><param name="e">Form-closing event data.</param>
        protected override void OnFormClosing(FormClosingEventArgs e) { cancellation.Cancel(); base.OnFormClosing(e); }
        /// <summary>Removes theme notifications and disposes the cancellation source.</summary>
        private void DisposeRuntime() { if (runtimeDisposed) return; runtimeDisposed = true; UiTheme.Changed -= ApplyAppearance; cancellation.Cancel(); cancellation.Dispose(); }
        /// <summary>Opens the update window as a modal child of the VBE host.</summary><param name="vbe">VBE automation object.</param>
        internal static void ShowForVbe(object vbe)
        {
            IWin32Window owner = null;
            try { owner = new VbeOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd))); } catch (Exception) { }
            using (var window = new UpdateWindow()) AddIn.ShowModal(window, owner);
        }
        /// <summary>WinForms owner wrapper around the native VBE main window.</summary>
        private sealed class VbeOwner : IWin32Window { /// <summary>Creates a modal owner wrapper.</summary><param name="handle">VBE main-window handle.</param>
internal VbeOwner(IntPtr handle) { Handle = handle; } /// <summary>Gets the native window handle used as the modal owner.</summary><value>VBE main-window handle.</value>
public IntPtr Handle { get; } }
    }
}
