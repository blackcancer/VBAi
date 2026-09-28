using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Progression de l’installateur externe, entièrement construite dans le Designer.</summary>
    internal sealed partial class UpdateProgressWindow : Form
    {
        private UpdateInstallJob job;
        private UpdateInstallerRunner runner;
        private string root;
        private bool polling, background;
        private bool prerequisite, closingInternally;
        internal Func<WebViewRuntimePrerequisite> CreatePrerequisite = () => new WebViewRuntimePrerequisite();
        internal Action<int> SetExitCode = code => Environment.ExitCode = code;
        internal void ConfigureWebView()
        {
            prerequisite = true; Text = heading.Text = "VBAi · WebView2";
            version.Text = "Microsoft Edge WebView2 Runtime";
            status.Text = UpdateText.Get("Installing update…");
            cancel.Text = UpdateText.Get("Close"); cancel.Enabled = false;
        }
        public UpdateProgressWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            UiTheme.Attach(this);
        }
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
                if (completed) { timer.Stop(); progress.Visible = false; cancel.Text = UpdateText.Get("Close"); CloseBackground(); }
            }
            catch (Exception) { timer.Stop(); status.Text = UpdateText.Get("Installation failed. Check the installer log."); CloseBackground(); }
            finally { polling = false; }
        }
        private void CloseBackground()
        {
            if (!background) return;
            closingInternally = true;
            try { Close(); }
            finally { closingInternally = false; }
        }
        private void Close_Click(object sender, EventArgs e)
        {
            if (polling || runner?.Installing == true) return;
            if (job != null && !job.Completed) { job.Completed = true; job.Status = "Update cancelled."; job.Save(root); }
            Close();
        }
        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            if (closingInternally) return;
            if (e.CloseReason == CloseReason.UserClosing && polling) { e.Cancel = true; return; }
            if (e.CloseReason == CloseReason.UserClosing && job != null && !job.Completed)
            { job.Completed = true; job.Status = "Update cancelled."; job.Save(root); }
        }
    }
}
