using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class CrashReportWindowTests
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
    }
}
