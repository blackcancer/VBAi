using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class CrashReportWindowTests
    {
        [STATestMethod]
        public void WindowSendRequiresPreviewSavesBeforePublishingAndPreventsDuplicateSubmission()
        {
            using (var window = new CrashReportWindow())
            {
                string stored = null, copied = null; int sent = 0;
                window.Configure(new CrashReport());
                window.CopyText = value => copied = value;
                window.Store = (report, body) => { stored = body; return "report.md"; };
                UiInvoke.Field<TextBox>(window, "descriptionInput").Text = "My reproduction steps";
                StringAssert.Contains(UiInvoke.Field<TextBox>(window, "preview").Text, "\r\n\r\nMy reproduction steps");
                window.Delivery.Publish = (title, body, ct) => { Assert.AreEqual(stored, body); sent++; return Task.FromResult(CrashReport.Repository + "/issues/4"); };
                LlmBoundaryScope.Call(window, "Copy_Click", null, EventArgs.Empty);
                StringAssert.Contains(copied, "My reproduction steps");
                LlmBoundaryScope.Call(window, "Send_Click", null, EventArgs.Empty);
                Assert.AreEqual(1, sent);
                Assert.AreEqual(UiText.Get("Report published on GitHub."), UiInvoke.Field<Label>(window, "status").Text);
                Assert.IsFalse(UiInvoke.Field<Button>(window, "sendButton").Enabled);
                LlmBoundaryScope.Call(window, "Send_Click", null, EventArgs.Empty); Assert.AreEqual(1, sent);
            }
        }
        [STATestMethod]
        public void WindowLocksAnUncertainSubmissionAndBusyCloseButRetainsCopyAndSave()
        {
            using (var window = new CrashReportWindow())
            {
                var completion = new TaskCompletionSource<string>();
                window.Store = (report, body) => "report.md";
                window.Delivery.Publish = (title, body, ct) => completion.Task;
                window.Delivery.SendOutlook = (title, body) => throw new AssertFailedException("Email must not be sent");
                LlmBoundaryScope.Call(window, "Send_Click", null, EventArgs.Empty);
                Assert.IsFalse(UiInvoke.Field<Button>(window, "sendButton").Enabled);
                var args = new FormClosingEventArgs(CloseReason.UserClosing, false); LlmBoundaryScope.Call(window, "WindowClosing", null, args); Assert.IsTrue(args.Cancel);
                completion.SetException(new HttpRequestException("Disconnected"));
                for (int i = 0; i < 100 && !UiInvoke.Field<Button>(window, "closeButton").Enabled; i++) { Application.DoEvents(); Thread.Sleep(10); }
                Assert.IsTrue(UiInvoke.Field<Button>(window, "closeButton").Enabled);
                Assert.IsFalse(UiInvoke.Field<Button>(window, "sendButton").Enabled);
                Assert.IsTrue(UiInvoke.Field<Button>(window, "saveButton").Enabled);
                Assert.AreEqual(UiText.Get("Delivery is uncertain. Check GitHub or Outlook before trying again."), UiInvoke.Field<Label>(window, "status").Text);
            }
        }
        [STATestMethod]
        public void SaveFailureStopsAllTransportsAndReportsTheStorageFailure()
        {
            using (var window = new CrashReportWindow())
            {
                window.Store = (report, body) => throw new IOException("Disk unavailable");
                window.Delivery.Publish = (title, body, ct) => throw new AssertFailedException("No network on save failure");
                window.Delivery.SendOutlook = (title, body) => throw new AssertFailedException("No mail on save failure");
                LlmBoundaryScope.Call(window, "Send_Click", null, EventArgs.Empty);
                Assert.AreEqual(UiText.Get("Unable to save the report."), UiInvoke.Field<Label>(window, "status").Text);
                Assert.IsTrue(UiInvoke.Field<Button>(window, "sendButton").Enabled);
                LlmBoundaryScope.Call(window, "Email_Click", null, EventArgs.Empty);
                Assert.AreEqual(UiText.Get("Unable to save the report."), UiInvoke.Field<Label>(window, "status").Text);
            }
        }

        [STATestMethod]
        public void WindowIsLocalizedAndOpensFromTheVbeWithoutChat()
        {
            foreach (var language in UiLanguages.All)
            using (var scope = new LocalizationScope(language.CultureName))
            using (var window = new CrashReportWindow())
            {
                Assert.AreEqual(UiText.Get("VBAi · Report an issue"), window.Text);
                Assert.AreEqual(RightToLeft.No, UiInvoke.Field<TextBox>(window, "preview").RightToLeft);
                Assert.IsTrue(UiInvoke.Field<TextBox>(window, "preview").ReadOnly);
                Assert.AreEqual(language.CultureName == "ar-SA" ? System.Drawing.ContentAlignment.TopRight : System.Drawing.ContentAlignment.TopLeft,
                    UiInvoke.Field<Label>(window, "destination").TextAlign);
                StringAssert.Contains(UiInvoke.Field<Label>(window, "destination").Text, CrashReport.Recipient);
            }
            var original = AddIn.ShowModal; int shown = 0;
            try
            {
                AddIn.ShowModal = (form, owner) => { Assert.IsInstanceOfType(form, typeof(CrashReportWindow)); Assert.AreEqual(new IntPtr(123), owner.Handle); shown++; return DialogResult.Cancel; };
                CrashReportWindow.ShowForVbe(new Host()); Assert.AreEqual(1, shown);
            }
            finally { AddIn.ShowModal = original; }
        }
        public sealed class Host { public Main MainWindow { get; } = new Main(); }
        public sealed class Main { public int HWnd { get; } = 123; }
        private static string Status(CrashReportWindow window) => UiInvoke.Field<Label>(window, "status").Text;
        private static void Invoke(CrashReportWindow window, string method) => LlmBoundaryScope.Call(window, method, null, EventArgs.Empty);
        private static void Pump(Func<bool> complete)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            while (!complete() && watch.ElapsedMilliseconds < 3000) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(complete(), "The owned UI continuation must complete.");
        }

        [STATestMethod]
        public void ConfigurationPreviewAndDesignerStatesKeepInvalidReportsUnsendable()
        {
            using (var theme = new ThemeScope())
            using (var window = new CrashReportWindow())
            {
                Assert.ThrowsException<ArgumentNullException>(() => window.Configure(null));
                UiInvoke.Field<TextBox>(window, "descriptionInput").Text = new string('x', 8001);
                Assert.AreEqual(UiInvoke.Field<CrashReport>(window, "report").TechnicalDetails, UiInvoke.Field<TextBox>(window, "preview").Text);
                UiInvoke.Field<TextBox>(window, "descriptionInput").Text = "";
                UiInvoke.Field<TextBox>(window, "titleInput").Text = " ";
                Assert.IsFalse(UiInvoke.Field<Button>(window, "sendButton").Enabled);
                UiInvoke.Field<TextBox>(window, "titleInput").Text = new string('x', 181);
                Invoke(window, "Send_Click");
                Assert.AreEqual(UiText.Get("Unable to send the report. The local copy is available."), Status(window));
            }
            var context = System.ComponentModel.LicenseManager.CurrentContext;
            try
            {
                System.ComponentModel.LicenseManager.CurrentContext = new DesignContext();
                using (var window = new CrashReportWindow())
                {
                    Assert.IsNull(UiInvoke.Field<CrashReport>(window, "report"));
                    LlmBoundaryScope.Call(window, "RefreshPreview");
                }
            }
            finally { System.ComponentModel.LicenseManager.CurrentContext = context; }
        }

        [STATestMethod]
        public void AppearanceCoversOwnedThreadDispatchDarkLightAndHighContrast()
        {
            using (var theme = new ThemeScope())
            {
                var window = new CrashReportWindow();
                try
                {
                    IntPtr owned = window.Handle;
                    foreach (bool contrast in new[] { false, true })
                    foreach (bool dark in new[] { false, true })
                    {
                        UiTheme.HighContrast = () => contrast; ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                        LlmBoundaryScope.Call(window, "ApplyAppearance");
                        Assert.AreEqual(contrast ? System.Drawing.SystemColors.Highlight : System.Drawing.Color.FromArgb(37, 99, 235), UiInvoke.Field<Button>(window, "sendButton").BackColor);
                        Assert.AreEqual(contrast ? System.Drawing.SystemColors.HotTrack : dark ? System.Drawing.Color.FromArgb(147, 197, 253) : System.Drawing.Color.FromArgb(29, 78, 216), UiInvoke.Field<LinkLabel>(window, "issueLink").LinkColor);
                    }
                    UiTheme.HighContrast = () => false; ThemeScope.SetChoice(ThemeChoice.Light);
                    var worker = new Thread(() => LlmBoundaryScope.Call(window, "ApplyAppearance")); worker.Start(); Assert.IsTrue(worker.Join(3000));
                    Pump(() => UiInvoke.Field<LinkLabel>(window, "issueLink").LinkColor == System.Drawing.Color.FromArgb(29, 78, 216));
                }
                finally { window.Dispose(); }
                LlmBoundaryScope.Call(window, "ApplyAppearance");
            }
        }

        [STATestMethod]
        public void CopySaveAndIssueActionsReportTheirExactInjectedOutcomes()
        {
            using (var theme = new ThemeScope())
            using (var window = new CrashReportWindow())
            {
                string copied = null, opened = null;
                window.CopyText = text => copied = text;
                window.Store = (report, body) => { StringAssert.Contains(body, report.Id); return "owned-report.md"; };
                Invoke(window, "Save_Click"); Assert.AreEqual("owned-report.md", copied);
                Assert.AreEqual(UiText.Get("Local report saved. Its path has been copied."), Status(window));
                window.CopyText = text => throw new IOException("owned clipboard error");
                Invoke(window, "Copy_Click"); Assert.AreEqual(UiText.Get("Unable to copy technical details."), Status(window));
                Invoke(window, "Save_Click"); Assert.AreEqual(UiText.Get("Unable to save the report."), Status(window));
                window.Delivery.Publish = (title, body, token) => Task.FromResult(CrashReport.Repository + "/issues/9");
                window.Delivery.Send("title", "body", "owned.md", CancellationToken.None).GetAwaiter().GetResult();
                window.OpenLink = text => opened = text;
                LlmBoundaryScope.Call(window, "Issue_Click", null, null); Assert.AreEqual(window.Delivery.IssueUrl, opened);
                window.OpenLink = text => throw new IOException("owned browser refusal");
                LlmBoundaryScope.Call(window, "Issue_Click", null, null); Assert.AreEqual(UiText.Get("Unable to open the link."), Status(window));
            }
        }

        [STATestMethod]
        public void EmailSuccessUncertaintyAndFailurePreventDuplicateTransport()
        {
            foreach (int outcome in new[] { 0, 1, 2, 3 })
            using (var theme = new ThemeScope())
            using (var window = new CrashReportWindow())
            {
                int sends = 0, drafts = 0;
                window.Store = (report, body) => "owned.md";
                window.Delivery.OpenDraft = url => { StringAssert.StartsWith(url, "mailto:"); drafts++; };
                window.Delivery.SendOutlook = (title, body) =>
                {
                    sends++;
                    if (outcome == 2) throw new CrashMailUncertain();
                    if (outcome == 3) throw new IOException("owned error");
                    return outcome == 0;
                };
                Invoke(window, "Email_Click");
                Assert.AreEqual(1, sends); Assert.AreEqual(outcome == 1 ? 1 : 0, drafts);
                Assert.AreEqual(UiText.Get(outcome == 0 ? "Report handed to Outlook for sending." : outcome == 1 ? "Email draft opened. Attach the saved report, then send it." : outcome == 2 ? "Delivery is uncertain. Check GitHub or Outlook before trying again." : "Unable to send the report. The local copy is available."), Status(window));
                if (outcome < 3) { Invoke(window, "Email_Click"); Invoke(window, "Send_Click"); Assert.AreEqual(1, sends); }
            }
        }

        [STATestMethod]
        public void BusyAndDisposedContinuationsCannotCloseOrMutateTheOwnedWindow()
        {
            foreach (bool refused in new[] { false, true })
            using (var theme = new ThemeScope())
            {
                var window = new CrashReportWindow(); var completion = new TaskCompletionSource<string>(); int calls = 0;
                window.Store = (report, body) => "owned.md";
                window.Delivery.Publish = (title, body, token) => { calls++; return completion.Task; };
                Invoke(window, "Send_Click"); Invoke(window, "Send_Click"); Invoke(window, "Email_Click"); Assert.AreEqual(1, calls);
                var closing = new FormClosingEventArgs(CloseReason.ApplicationExitCall, false);
                LlmBoundaryScope.Call(window, "WindowClosing", null, closing); Assert.IsFalse(closing.Cancel);
                window.Dispose();
                if (refused) completion.SetException(new CrashCredentialUnavailable()); else completion.SetResult(CrashReport.Repository + "/issues/9");
                Application.DoEvents(); Application.DoEvents(); Assert.IsTrue(window.IsDisposed);
            }
            using (var window = new CrashReportWindow())
            {
                var closing = new FormClosingEventArgs(CloseReason.UserClosing, false);
                LlmBoundaryScope.Call(window, "WindowClosing", null, closing); Assert.IsFalse(closing.Cancel);
            }
        }

        [STATestMethod]
        public void ModalReportFallsBackToNoOwnerForAnInvalidHost()
        {
            var original = AddIn.ShowModal; int shown = 0;
            try
            {
                AddIn.ShowModal = (form, owner) => { Assert.IsNull(owner); shown++; return DialogResult.Cancel; };
                CrashReportWindow.ShowForVbe(null); CrashReportWindow.ShowReportForVbe(new object(), new CrashReport()); Assert.AreEqual(2, shown);
            }
            finally { AddIn.ShowModal = original; }
        }
    }
}
