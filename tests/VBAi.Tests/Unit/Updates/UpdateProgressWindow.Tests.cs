using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class UpdateProgressWindowTests
    {
        [STATestMethod]
        public void ExternalProgressWindowCancellationPersistsWithoutLaunchingAnything()
        {
            using (var scope = new UpdateScope())
            using (var window = new UpdateProgressWindow())
            {
                var job = UpdateInstallerRunnerTests.Job(scope); job.Culture = "fr-FR";
                window.Configure(scope.Root, job, false);
                Assert.AreEqual("Mises à jour VBAi", window.Text);
                LlmBoundaryScope.Call(window, "Close_Click", null, EventArgs.Empty);
                Assert.AreEqual("Update cancelled.", UpdateInstallJob.Load(scope.Root).Status);
            }
        }
        [STATestMethod]
        public void ProgressConfigurationSupportsCurrentCultureRtlAndForegroundOrBackgroundVisibility()
        {
            foreach (string language in new[] { null, "ar-SA", "en-US" })
            foreach (bool background in new[] { false, true })
            using (var fixture = new UpdateProgressFixture(background, language))
            {
                Assert.AreEqual(language == "ar-SA", fixture.Window.RightToLeft == RightToLeft.Yes);
                Assert.AreEqual(language == "ar-SA" ? System.Drawing.ContentAlignment.TopRight : System.Drawing.ContentAlignment.TopLeft, UiInvoke.Field<Label>(fixture.Window, "version").TextAlign);
                fixture.Window.Show(); UpdateUiPump.Until(() => fixture.Window.Visible == !background);
                Assert.AreEqual("VBAi " + fixture.Job.TargetVersion, UiInvoke.Field<Label>(fixture.Window, "version").Text);
            }
        }

        [STATestMethod]
        public void PrerequisiteProgressReportsOwnedSuccessAndFailureWithoutGlobalExitCodeOrInstallation()
        {
            foreach (bool failed in new[] { false, true })
            using (var fixture = new UpdateProgressFixture(background: true))
            {
                int? exit = null; fixture.Window.ConfigureWebView();
                Assert.AreEqual("VBAi · WebView2", fixture.Window.Text); Assert.IsFalse(UiInvoke.Field<Button>(fixture.Window, "cancel").Enabled);
                fixture.Window.SetExitCode = code => exit = code;
                fixture.Window.CreatePrerequisite = () => new WebViewRuntimePrerequisite
                { IsInstalled = () => failed ? throw new IOException("owned prerequisite failure") : true };
                fixture.Window.Show(); UpdateUiPump.Until(() => exit.HasValue);
                Assert.AreEqual(failed ? 1 : 0, exit.Value);
                Assert.AreEqual(UpdateText.Get(failed ? "Installation failed. Check the installer log." : "Update installed. Restart the VBA host."), UiInvoke.Field<Label>(fixture.Window, "status").Text);
                Assert.IsTrue(UiInvoke.Field<Button>(fixture.Window, "cancel").Enabled); Assert.IsFalse(UiInvoke.Field<ProgressBar>(fixture.Window, "progress").Visible);
                Assert.IsFalse(UiInvoke.Field<bool>(fixture.Window, "polling"));
            }
        }

        [STATestMethod]
        public void ProgressPollingHonorsLoadedCompletionIdentityAndSettlesForegroundAndBackgroundWindows()
        {
            foreach (string outcome in new[] { "no-file", "matching", "different", "waiting", "completed" })
            foreach (bool background in new[] { false, true })
            using (var fixture = new UpdateProgressFixture(background))
            {
                if (outcome == "no-file") File.Delete(Path.Combine(fixture.Scope.Root, "pending.json"));
                if (outcome == "matching") { var latest = UpdateInstallJob.Load(fixture.Scope.Root); latest.Completed = true; latest.Status = "Update cancelled."; latest.Save(fixture.Scope.Root); }
                if (outcome == "different") { var latest = UpdateInstallJob.Load(fixture.Scope.Root); latest.Completed = true; latest.Sha256 = new string('a', 64); latest.InstallerPath = UpdatePaths.AssetPath(fixture.Scope.Root, latest.Sha256, "VBAi-Setup-win-x64.exe"); latest.Save(fixture.Scope.Root); }
                if (outcome == "waiting") { UpdateState.RecordHost(); fixture.Runner.IsAlive = lease => true; }
                if (outcome == "completed") { fixture.Job.Completed = true; fixture.Job.Save(fixture.Scope.Root); }
                fixture.Window.Show(); fixture.Poll();
                Assert.IsFalse(UiInvoke.Field<bool>(fixture.Window, "polling"));
                Assert.AreEqual(background && outcome != "waiting", fixture.Window.IsDisposed, outcome);
                if (outcome == "different") Assert.AreEqual(new string('a', 64), UpdateInstallJob.Load(fixture.Scope.Root).Sha256, "Internal completion must preserve the new pending job.");
                if (outcome == "waiting") Assert.IsFalse(fixture.Job.Completed);
                else if (!background) Assert.AreEqual(UpdateText.Get("Close"), UiInvoke.Field<Button>(fixture.Window, "cancel").Text);
            }
            using (var window = new UpdateProgressWindow())
            { UpdateUiPump.Call(window, "Poll"); Assert.IsFalse(UiInvoke.Field<bool>(window, "polling")); }
        }

        [STATestMethod]
        public void ProgressPollingErrorsStopAndCloseOnlyItsBackgroundWindow()
        {
            foreach (bool background in new[] { false, true })
            using (var fixture = new UpdateProgressFixture(background))
            {
                File.WriteAllText(Path.Combine(fixture.Scope.Root, "pending.json"), "malformed owned state");
                string closingStatus = null; fixture.Window.FormClosing += (sender, args) => closingStatus = UiInvoke.Field<Label>(fixture.Window, "status").Text;
                fixture.Window.Show(); fixture.Poll();
                Assert.AreEqual(UpdateText.Get("Installation failed. Check the installer log."), background ? closingStatus : UiInvoke.Field<Label>(fixture.Window, "status").Text);
                Assert.IsFalse(UiInvoke.Field<System.Windows.Forms.Timer>(fixture.Window, "timer").Enabled);
                Assert.AreEqual(background, fixture.Window.IsDisposed);
            }
        }

        [STATestMethod]
        public void PendingPollAndUserCloseAreBlockedUntilTheOwnedRunnerCompletes()
        {
            using (var fixture = new UpdateProgressFixture())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                fixture.Runner.Install = path => { entered.Set(); Assert.IsTrue(release.Wait(3000)); return 0; };
                UpdateUiPump.Call(fixture.Window, "Poll"); Assert.IsTrue(entered.Wait(3000));
                UpdateUiPump.Call(fixture.Window, "Poll"); UpdateUiPump.Call(fixture.Window, "Close_Click"); Assert.IsFalse(fixture.Window.IsDisposed);
                var closing = new FormClosingEventArgs(CloseReason.UserClosing, false);
                LlmBoundaryScope.Call(fixture.Window, "WindowClosing", null, closing); Assert.IsTrue(closing.Cancel);
                release.Set(); UpdateUiPump.Until(() => !UiInvoke.Field<bool>(fixture.Window, "polling"));
                Assert.IsTrue(fixture.Job.Completed); Assert.IsFalse(fixture.Runner.Installing);
            }
        }

        [STATestMethod]
        public void InstallingRunnerAndDisposedPollContinuationPreserveTheOwnedState()
        {
            foreach (bool disposed in new[] { false, true })
            using (var fixture = new UpdateProgressFixture())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                fixture.Runner.Install = path => { entered.Set(); Assert.IsTrue(release.Wait(3000)); return 0; };
                if (disposed)
                {
                    UpdateUiPump.Call(fixture.Window, "Poll"); Assert.IsTrue(entered.Wait(3000)); fixture.Window.Dispose(); release.Set();
                    UpdateUiPump.Until(() => !UiInvoke.Field<bool>(fixture.Window, "polling"));
                    Assert.AreEqual(UpdateText.Get("Waiting for VBA hosts to close."), UiInvoke.Field<Label>(fixture.Window, "status").Text);
                }
                else
                {
                    var tick = Task.Run(() => fixture.Runner.Tick(fixture.Job)); Assert.IsTrue(entered.Wait(3000));
                    UpdateUiPump.Call(fixture.Window, "Close_Click"); Assert.IsFalse(fixture.Window.IsDisposed);
                    release.Set(); Assert.IsTrue(tick.GetAwaiter().GetResult());
                    UpdateUiPump.Call(fixture.Window, "Close_Click"); Assert.IsTrue(fixture.Window.IsDisposed);
                }
            }
        }

        [STATestMethod]
        public void ClosingMessagesOnlyCancelPendingJobsForUserInitiatedClose()
        {
            foreach (bool user in new[] { false, true })
            foreach (bool completed in new[] { false, true })
            using (var fixture = new UpdateProgressFixture())
            {
                fixture.Job.Completed = completed;
                var closing = new FormClosingEventArgs(user ? CloseReason.UserClosing : CloseReason.ApplicationExitCall, false);
                LlmBoundaryScope.Call(fixture.Window, "WindowClosing", null, closing);
                Assert.IsFalse(closing.Cancel); Assert.AreEqual(completed || user, fixture.Job.Completed);
                if (user && !completed) Assert.AreEqual("Update cancelled.", UpdateInstallJob.Load(fixture.Scope.Root).Status);
            }
            using (var window = new UpdateProgressWindow())
            {
                var closing = new FormClosingEventArgs(CloseReason.UserClosing, false);
                LlmBoundaryScope.Call(window, "WindowClosing", null, closing); Assert.IsFalse(closing.Cancel);
                UpdateUiPump.Call(window, "Close_Click"); Assert.IsTrue(window.IsDisposed);
            }
        }
        [STATestMethod]
        public void ProgressDesignerDisposeHandlesUnmanagedAndAlreadyReleasedOwnedComponentContainers()
        {
            using (var window = new UpdateProgressWindow())
            {
                UiInvoke.Call(typeof(UpdateProgressWindow), "Dispose", window, false);
                Assert.IsFalse(window.IsDisposed); Assert.IsNotNull(UiInvoke.Field<System.ComponentModel.IContainer>(window, "components"));
            }
            using (var window = new UpdateProgressWindow())
            {
                UiInvoke.Field<System.ComponentModel.IContainer>(window, "components").Dispose();
                LlmBoundaryScope.Set(window, "components", null); window.Dispose(); Assert.IsTrue(window.IsDisposed);
            }
            var context = System.ComponentModel.LicenseManager.CurrentContext;
            try
            {
                System.ComponentModel.LicenseManager.CurrentContext = new DesignContext();
                using (var window = new UpdateProgressWindow()) Assert.IsFalse(window.IsHandleCreated);
            }
            finally { System.ComponentModel.LicenseManager.CurrentContext = context; }
        }
    }
}
