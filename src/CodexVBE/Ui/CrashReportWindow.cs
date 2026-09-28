using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Écran de support éditable dans le Designer : aperçu avant transmission et secours local.</summary>
    internal sealed partial class CrashReportWindow : Form
    {
        private CrashReport report;
        private bool busy, submitted, uncertain, runtimeInitialized;
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        internal CrashReportDelivery Delivery = new CrashReportDelivery();
        internal Action<string> CopyText = Clipboard.SetText;
        internal Action<string> OpenLink = SafeLinks.Open;
        internal Func<CrashReport, string, string> Store = (report, body) => report.Save(body);
        private string savedPath;

        public CrashReportWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            runtimeInitialized = true;
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
            preview.RightToLeft = RightToLeft.No;
            reportIdentity.RightToLeft = RightToLeft.No;
            destination.TextAlign = UiText.Culture.TextInfo.IsRightToLeft ? ContentAlignment.TopRight : ContentAlignment.TopLeft;
            Configure(new CrashReport());
            ApplyAppearance();
            UiTheme.Changed += ApplyAppearance;
        }

        internal void Configure(CrashReport value)
        {
            report = value ?? throw new ArgumentNullException(nameof(value));
            titleInput.Text = report.DefaultTitle;
            reportIdentity.Text = "VBAi · " + report.Id;
            RefreshPreview();
        }

        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            sendButton.BackColor = UiTheme.HighContrast() ? SystemColors.Highlight : Color.FromArgb(37, 99, 235);
            sendButton.ForeColor = UiTheme.HighContrast() ? SystemColors.HighlightText : Color.White;
            issueLink.LinkColor = issueLink.ActiveLinkColor = issueLink.VisitedLinkColor = UiTheme.HighContrast() ? SystemColors.HotTrack :
                UiTheme.Dark ? Color.FromArgb(147, 197, 253) : Color.FromArgb(29, 78, 216);
            preview.BackColor = UiTheme.Surface;
            preview.ForeColor = UiTheme.Foreground;
        }

        private void ContentChanged(object sender, EventArgs e) { RefreshPreview(); }
        private void RefreshPreview()
        {
            if (report == null) return;
            try { preview.Text = report.Body(titleInput.Text, descriptionInput.Text); }
            catch (ArgumentException) { preview.Text = report.TechnicalDetails; }
            sendButton.Enabled = emailButton.Enabled = !busy && !submitted && !uncertain && !string.IsNullOrWhiteSpace(titleInput.Text);
        }

        private async void Send_Click(object sender, EventArgs e)
        {
            if (busy || submitted || uncertain) return;
            SetBusy(true);
            try
            {
                string body = report.Body(titleInput.Text, descriptionInput.Text);
                try { savedPath = Store(report, body); }
                catch (Exception) { status.Text = UiText.Get("Unable to save the report."); return; }
                status.Text = UiText.Get("Sending report…");
                var result = await Delivery.Send(titleInput.Text, body, savedPath, cancellation.Token);
                if (!IsDisposed) ShowResult(result);
            }
            catch (Exception) { if (!IsDisposed) status.Text = UiText.Get("Unable to send the report. The local copy is available."); }
            finally { if (!IsDisposed) SetBusy(false); }
        }

        private void Email_Click(object sender, EventArgs e)
        {
            if (busy || submitted || uncertain) return;
            SetBusy(true);
            try
            {
                string body = report.Body(titleInput.Text, descriptionInput.Text);
                try { savedPath = Store(report, body); }
                catch (Exception) { status.Text = UiText.Get("Unable to save the report."); return; }
                ShowResult(Delivery.Email(titleInput.Text, body, savedPath));
            }
            catch (CrashMailUncertain) { uncertain = true; status.Text = UiText.Get("Delivery is uncertain. Check GitHub or Outlook before trying again."); }
            catch (Exception) { status.Text = UiText.Get("Unable to send the report. The local copy is available."); }
            finally { SetBusy(false); }
        }

        private void ShowResult(CrashDeliveryResult result)
        {
            uncertain = result == CrashDeliveryResult.Uncertain;
            submitted = !uncertain;
            status.Text = UiText.Get(result == CrashDeliveryResult.GitHub ? "Report published on GitHub." :
                result == CrashDeliveryResult.Outlook ? "Report handed to Outlook for sending." :
                result == CrashDeliveryResult.Draft ? "Email draft opened. Attach the saved report, then send it." :
                "Delivery is uncertain. Check GitHub or Outlook before trying again.");
            issueLink.Visible = result == CrashDeliveryResult.GitHub;
        }

        private void SetBusy(bool value)
        {
            busy = value;
            progress.Visible = value;
            titleInput.ReadOnly = descriptionInput.ReadOnly = value || submitted || uncertain;
            // Once sending starts, keep the modal owner alive until delivery has returned.
            closeButton.Enabled = !value;
            RefreshPreview();
        }

        private void Copy_Click(object sender, EventArgs e)
        {
            try { CopyText(report.Body(titleInput.Text, descriptionInput.Text)); status.Text = UiText.Get("Technical details copied."); }
            catch (Exception) { status.Text = UiText.Get("Unable to copy technical details."); }
        }
        private void Save_Click(object sender, EventArgs e)
        {
            try
            {
                savedPath = Store(report, report.Body(titleInput.Text, descriptionInput.Text));
                CopyText(savedPath);
                status.Text = UiText.Get("Local report saved. Its path has been copied.");
            }
            catch (Exception) { status.Text = UiText.Get("Unable to save the report."); }
        }
        private void Issue_Click(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try { OpenLink(Delivery.IssueUrl); }
            catch (Exception) { status.Text = UiText.Get("Unable to open the link."); }
        }
        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            if (busy && e.CloseReason == CloseReason.UserClosing) e.Cancel = true;
        }
        private void DisposeRuntime()
        {
            if (runtimeInitialized) UiTheme.Changed -= ApplyAppearance;
            cancellation.Cancel();
            cancellation.Dispose();
        }

        internal static void ShowForVbe(object vbe, Exception error = null) { ShowReportForVbe(vbe, new CrashReport(error)); }

        internal static void ShowReportForVbe(object vbe, CrashReport report)
        {
            IWin32Window owner = null;
            try { owner = new NativeOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd))); }
            catch (Exception) { }
            using (var window = new CrashReportWindow())
            {
                window.Configure(report);
                AddIn.ShowModal(window, owner);
            }
        }
        private sealed class NativeOwner : IWin32Window
        {
            internal NativeOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
    }
}
