using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Écran de support éditable dans le Designer : aperçu avant transmission et secours local.</summary>
    internal sealed partial class CrashReportWindow : Form
    {

        /// <summary>Report currently displayed and edited by the window.</summary>
        private CrashReport report;

        /// <summary>Delivery in progress, completed submission, uncertain outcome, and theme-subscription flags.</summary>
        private bool busy;
        private bool submitted;
        private bool uncertain;
        private readonly bool runtimeInitialized;

        /// <summary>Cancellation source used when the dialog closes during an asynchronous send.</summary>
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();

        /// <summary>Delivery service used by the send and email actions.</summary>
        internal CrashReportDelivery Delivery = new CrashReportDelivery();

        /// <summary>Action used to copy report text or its saved path.</summary>
        internal Action<string> CopyText = Clipboard.SetText;

        /// <summary>Opens a validated support or issue URL using the shared safe-link policy.</summary>
        internal Action<string> OpenLink = SafeLinks.Open;

        /// <summary>Delegate that saves a report body and returns its local path.</summary>
        internal Func<CrashReport, string, string> Store = (report, body) => report.Save(body);

        /// <summary>Path of the local backup saved before delivery.</summary>
        private string savedPath;

        /// <summary>Creates the report dialog and initializes its default editable report.</summary>
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

        /// <summary>Sets the report shown by this window and refreshes its identity and preview.</summary>
        /// <param name="value">Report to display.</param>
        /// <exception cref="ArgumentNullException">The report is null.</exception>
        internal void Configure(CrashReport value)
        {
            report = value ?? throw new ArgumentNullException(nameof(value));
            titleInput.Text = report.DefaultTitle;
            reportIdentity.Text = "VBAi · " + report.Id;
            RefreshPreview();
        }

        /// <summary>Applies the current UI theme to the send button, issue link, and preview.</summary>
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

        /// <summary>Refreshes the report preview after an editable field changes.</summary>
        /// <param name="sender">Input control that changed.</param><param name="e">Change event arguments.</param>
        private void ContentChanged(object sender, EventArgs e) { RefreshPreview(); }

        /// <summary>Formats the current input as Markdown and updates whether delivery actions are available.</summary>
        private void RefreshPreview()
        {
            if (report == null) return;
            try { preview.Text = report.Body(titleInput.Text, descriptionInput.Text); }
            catch (ArgumentException) { preview.Text = report.TechnicalDetails; }
            sendButton.Enabled = emailButton.Enabled = !busy && !submitted && !uncertain && !string.IsNullOrWhiteSpace(titleInput.Text);
        }

        /// <summary>Saves and asynchronously sends the report through the configured delivery service.</summary>
        /// <param name="sender">Send button.</param><param name="e">Click event arguments.</param>
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

        /// <summary>Saves the report and hands it to Outlook or a local mail draft.</summary>
        /// <param name="sender">Email button.</param><param name="e">Click event arguments.</param>
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

        /// <summary>Updates delivery flags, status text, and issue-link visibility for a delivery result.</summary>
        /// <param name="result">Result reported by the delivery service.</param>
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

        /// <summary>Updates the in-progress state and disables editing while delivery runs.</summary>
        /// <param name="value">Whether delivery is currently running.</param>
        private void SetBusy(bool value)
        {
            busy = value;
            progress.Visible = value;
            titleInput.ReadOnly = descriptionInput.ReadOnly = value || submitted || uncertain;
            // Once sending starts, keep the modal owner alive until delivery has returned.
            closeButton.Enabled = !value;
            RefreshPreview();
        }

        /// <summary>Copies the formatted report body to the clipboard.</summary>
        /// <param name="sender">Copy button.</param><param name="e">Click event arguments.</param>
        private void Copy_Click(object sender, EventArgs e)
        {
            try { CopyText(report.Body(titleInput.Text, descriptionInput.Text)); status.Text = UiText.Get("Technical details copied."); }
            catch (Exception) { status.Text = UiText.Get("Unable to copy technical details."); }
        }

        /// <summary>Saves a local Markdown copy and copies its path to the clipboard.</summary>
        /// <param name="sender">Save button.</param><param name="e">Click event arguments.</param>
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

        /// <summary>Opens the successfully published GitHub issue.</summary>
        /// <param name="sender">Issue link.</param><param name="e">Link-click event arguments.</param>
        private void Issue_Click(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try { OpenLink(Delivery.IssueUrl); }
            catch (Exception) { status.Text = UiText.Get("Unable to open the link."); }
        }

        /// <summary>Prevents user closure while a delivery operation is still running.</summary>
        /// <param name="sender">Report window.</param><param name="e">Closing event arguments that may be cancelled.</param>
        private void WindowClosing(object sender, FormClosingEventArgs e)
        {
            if (busy && e.CloseReason == CloseReason.UserClosing) e.Cancel = true;
        }

        /// <summary>Unsubscribes from theme changes and cancels any outstanding send.</summary>
        private void DisposeRuntime()
        {
            if (runtimeInitialized) UiTheme.Changed -= ApplyAppearance;
            cancellation.Cancel();
            cancellation.Dispose();
        }

        /// <summary>Creates and shows a report dialog owned by the VBE main window.</summary>
        /// <param name="vbe">VBE automation object used to obtain the native owner handle.</param>
        /// <param name="error">Optional exception used to seed the report metadata.</param>
        internal static void ShowForVbe(object vbe, Exception error = null) { ShowReportForVbe(vbe, new CrashReport(error)); }

        /// <summary>Shows the supplied report in a modal window owned by the VBE when its handle is available.</summary>
        /// <param name="vbe">VBE automation object used to obtain the native owner handle.</param>
        /// <param name="report">Report to show.</param>
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

        /// <summary>WinForms owner wrapper for a native VBE window handle.</summary>
        private sealed class NativeOwner : IWin32Window
        {

            /// <summary>Creates the wrapper for a native owner window.</summary><param name="handle">Native window handle.</param>
            internal NativeOwner(IntPtr handle) { Handle = handle; }

            /// <summary>Gets the native owner-window handle.</summary><value>Handle supplied to the constructor.</value>
            public IntPtr Handle { get; }
        }
    }
}
