using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Progression de l’installateur externe, entièrement construite dans le Designer.</summary>
    internal sealed partial class UpdateProgressWindow : Form
    {

        /// <summary>Install job polled while the separate updater waits for hosts to close.</summary>
        private UpdateInstallJob job;

        /// <summary>Runner that verifies host leases and starts the installer.</summary>
        private UpdateInstallerRunner runner;

        /// <summary>Update root containing the pending job and host leases.</summary>
        private string root;

        /// <summary>Whether a poll is active and whether the updater was launched in background mode.</summary>
        private bool polling, background;

        /// <summary>Whether this window is handling runtime prerequisite installation instead of a product update.</summary>
        private bool prerequisite, closingInternally;

        /// <summary>Maintains the create prerequisite state for update progress window.</summary>
        internal Func<WebViewRuntimePrerequisite> CreatePrerequisite = () => new WebViewRuntimePrerequisite();

        /// <summary>Maintains the set exit code state for update progress window.</summary>
        internal Action<int> SetExitCode = code => Environment.ExitCode = code;

        /// <summary>Configures this window for the WebView2 runtime prerequisite flow.</summary>
        internal void ConfigureWebView()
        {
            prerequisite = true; Text = heading.Text = "VBAi · WebView2";
            version.Text = "Microsoft Edge WebView2 Runtime";
            status.Text = UpdateText.Get("Installing update…");
            cancel.Text = UpdateText.Get("Close"); cancel.Enabled = false;
        }

        /// <summary>Creates the progress dialog and attaches current theme handling.</summary>
        public UpdateProgressWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            UiTheme.Attach(this);
        }

        /// <summary>Sets the update job, selects its language, and begins progress polling.</summary>
        /// <param name="root">Update root containing job state.</param><param name="job">Job to display and monitor.</param><param name="background">Whether the worker should hide after showing.</param>
        internal void Configure(string root, UpdateInstallJob job, bool background)
        {
            this.root = root; this.job = job; this.background = background;
            runner = new UpdateInstallerRunner(root);
            UpdateText.Culture = job.Culture ?? System.Globalization.CultureInfo.CurrentUICulture.Name;
            var language = UiLanguages.For(System.Globalization.CultureInfo.GetCultureInfo(UpdateText.Culture));
            bool rtl = System.Globalization.CultureInfo.GetCultureInfo(language.CultureName).TextInfo.IsRightToLeft;
            RightToLeft = rtl ? RightToLeft.Yes : RightToLeft.No; RightToLeftLayout = rtl;
            version.RightToLeft = RightToLeft.No; version.TextAlign = rtl ? System.Drawing.ContentAlignment.TopRight : System.Drawing.ContentAlignment.TopLeft;
            Text = UpdateText.Get(Text); heading.Text = UpdateText.Get(heading.Text); cancel.Text = UpdateText.Get(cancel.Text);
            status.Text = UpdateText.Get(job.Status);
            version.Text = "VBAi " + job.TargetVersion;
            timer.Start();
        }

        /// <summary>Hides background windows and installs the WebView prerequisite when configured for that flow.</summary>
        /// <param name="e">Shown event data.</param>
        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            if (background) Hide();
            if (!prerequisite) return;
            polling = true;
            try
            {
                await CreatePrerequisite().Ensure(System.IO.Path.GetTempPath());
                status.Text = UpdateText.Get("Update installed. Restart the VBA host."); SetExitCode(0);
            }
            catch (Exception) { status.Text = UpdateText.Get("Installation failed. Check the installer log."); SetExitCode(1); }
            finally { polling = false; cancel.Enabled = true; progress.Visible = false; }
        }

        /// <summary>Refreshes worker status and advances one runner tick when no poll is already active.</summary>
        /// <param name="sender">Timer that raised the tick.</param><param name="e">Tick event arguments.</param>
        private async void Poll(object sender, EventArgs e)
        {
            if (job != null) status.Text = UpdateText.Get(job.Status);
            if (runner != null) cancel.Enabled = !runner.Installing;
            if (polling || job == null) return;
            polling = true;
            try
            {
                var latest = UpdateInstallJob.Load(root);
                if (latest?.Completed == true && latest.Sha256 == job.Sha256) job = latest;
                bool completed = await Task.Run(() => runner.Tick(job));
                if (IsDisposed) return;
                status.Text = UpdateText.Get(job.Status); cancel.Enabled = true;
                if (completed) { SetExitCode(job.Completed && job.Succeeded ? 0 : 1); timer.Stop(); progress.Visible = false; cancel.Text = UpdateText.Get("Close"); CloseBackground(); }
            }
            catch (Exception) { SetExitCode(1); timer.Stop(); status.Text = UpdateText.Get("Installation failed. Check the installer log."); CloseBackground(); }
            finally { polling = false; }
        }

        /// <summary>Closes background for update progress window.</summary>
        private void CloseBackground()
        {
            if (!background) return;
            closingInternally = true;
            try { Close(); }
            finally { closingInternally = false; }
        }

        /// <summary>Marks an unfinished job cancelled and closes when no installation is in progress.</summary>
        /// <param name="sender">Close or cancel button.</param><param name="e">Click event arguments.</param>
        private void Close_Click(object sender, EventArgs e)
        {
            if (polling || runner?.Installing == true) return;
            CancelPendingJob();
            Close();
        }

        /// <summary>Blocks user closure during polling and records cancellation for an unfinished job.</summary>
        /// <param name="sender">Progress dialog.</param><param name="e">Closing event data that may be cancelled.</param>
        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            if (closingInternally) return;
            if (e.CloseReason == CloseReason.UserClosing && (polling || runner?.Installing == true)) { e.Cancel = true; return; }
            if (e.CloseReason == CloseReason.UserClosing) CancelPendingJob();
        }

        /// <summary>Records a user cancellation as an unsuccessful final result.</summary>
        private void CancelPendingJob()
        {
            if (job == null || job.Completed) return;
            job.Completed = true; job.Succeeded = false; job.Status = "Update cancelled.";
            SetExitCode(1);
            job.Save(root);
        }
    }
}
