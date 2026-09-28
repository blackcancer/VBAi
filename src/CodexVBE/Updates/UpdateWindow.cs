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
        private UpdateRelease release;
        private string downloaded;
        private bool busy;
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private bool runtimeDisposed;
        internal Func<UpdateFeed> CreateFeed = () => new UpdateFeed();
        internal Func<UpdatePreferences> ReadPreferences = () => UpdateState.Load();
        internal Action<UpdatePreferences> StorePreferences = p => UpdateState.Save(p);
        internal Func<bool> ManagedInstallation = () => UpdateInstallation.IsManaged(UpdateState.InstallationDirectory);
        internal Action<UpdateRelease, string, bool> Schedule = UpdateCoordinator.Schedule;
        internal Func<UpdateFeed, UpdateVersion, bool, CancellationToken, Task<UpdateRelease>> CheckRelease = (feed, version, previews, token) => feed.Check(version, previews, null, token);
        internal Func<UpdateFeed, UpdateAsset, string, IProgress<int>, CancellationToken, Task<string>> DownloadInstaller = (feed, asset, root, progress, token) => feed.Download(asset, root, progress, token);
        internal Func<Action, Task> RunBackground = action => Task.Run(action);
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
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
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
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            check.BackColor = UiTheme.HighContrast() ? SystemColors.Highlight : Color.FromArgb(37, 99, 235);
            check.ForeColor = UiTheme.HighContrast() ? SystemColors.HighlightText : Color.White;
        }
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

        private void SavePreferences()
        {
            var preferences = ReadPreferences();
            preferences.CheckAutomatically = automaticCheck.Checked;
            preferences.DownloadAutomatically = automaticDownload.Checked;
            preferences.InstallAutomatically = ManagedInstallation() && automaticInstall.Checked && automaticDownload.Checked;
            preferences.IncludePrereleases = previews.Checked;
            StorePreferences(preferences);
        }
        private void Save_Click(object sender, EventArgs e)
        {
            try { SavePreferences(); status.Text = UiText.Get("Update preferences saved."); }
            catch (Exception) { status.Text = UiText.Get("Unable to save update preferences."); }
        }
        private void DownloadPreferenceChanged(object sender, EventArgs e)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            automaticInstall.Enabled = !busy && ManagedInstallation() && automaticDownload.Checked;
            if (!automaticDownload.Checked) automaticInstall.Checked = false;
        }
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
        private async void Install_Click(object sender, EventArgs e)
        {
            if (busy || downloaded == null || release == null) return;
            SetBusy(true);
            try { SavePreferences(); await RunBackground(() => Schedule(release, downloaded, false)); if (!IsDisposed) { status.Text = UiText.Get("Waiting for VBA hosts to close."); cancelPending.Enabled = true; } }
            catch (Exception) { if (!IsDisposed) status.Text = UiText.Get("Unable to schedule the update."); }
            finally { if (!IsDisposed) { SetBusy(false); install.Enabled = false; } }
        }
        private void Skip_Click(object sender, EventArgs e)
        {
            if (release == null) return;
            try { var p = ReadPreferences(); p.SkippedVersion = release.Version.Text; StorePreferences(p); status.Text = UiText.Get("This version will be skipped automatically."); }
            catch (Exception) { status.Text = UiText.Get("Unable to save update preferences."); }
        }
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
        private void SetBusy(bool value)
        {
            busy = value; check.Enabled = save.Enabled = automaticCheck.Enabled = automaticDownload.Enabled = previews.Enabled = !value;
            automaticInstall.Enabled = !value && ManagedInstallation() && automaticDownload.Checked;
            download.Enabled = !value && release?.Installer?.Hash != null;
            install.Enabled = !value && downloaded != null && ManagedInstallation();
            skip.Enabled = !value && release != null;
            progress.Visible = value; progress.Style = ProgressBarStyle.Marquee;
        }
        protected override void OnFormClosing(FormClosingEventArgs e) { cancellation.Cancel(); base.OnFormClosing(e); }
        private void DisposeRuntime() { if (runtimeDisposed) return; runtimeDisposed = true; UiTheme.Changed -= ApplyAppearance; cancellation.Cancel(); cancellation.Dispose(); }
        internal static void ShowForVbe(object vbe)
        {
            IWin32Window owner = null;
            try { owner = new VbeOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd))); } catch (Exception) { }
            using (var window = new UpdateWindow()) AddIn.ShowModal(window, owner);
        }
        private sealed class VbeOwner : IWin32Window { internal VbeOwner(IntPtr handle) { Handle = handle; } public IntPtr Handle { get; } }
    }
}
