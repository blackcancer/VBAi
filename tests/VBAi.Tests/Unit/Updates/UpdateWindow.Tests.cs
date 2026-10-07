using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class UpdateWindowTests
    {
        [STATestMethod]
        public void WindowKeepsSettingsAndReleaseNotesInDesignerControlsWithoutAnyNativeInstaller()
        {
            using (var scope = new UpdateScope())
            using (var window = new UpdateWindow())
            {
                var prefs = new UpdatePreferences(); window.ReadPreferences = () => prefs; window.StorePreferences = p => prefs = p;
                window.ManagedInstallation = () => false;
                window.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler("[{\"tag_name\":\"v1.2.3\",\"body\":\"Fixture release notes\"}]"), ct => throw new AssertFailedException());
                LlmBoundaryScope.Call(window, "Check_Click", null, EventArgs.Empty);
                Assert.AreEqual("Fixture release notes", UiInvoke.Field<TextBox>(window, "notes").Text);
                Assert.IsFalse(UiInvoke.Field<Button>(window, "install").Enabled); Assert.IsFalse(UiInvoke.Field<Button>(window, "download").Enabled);
                UiInvoke.Field<CheckBox>(window, "automaticDownload").Checked = false;
                LlmBoundaryScope.Call(window, "Save_Click", null, EventArgs.Empty); Assert.IsFalse(prefs.DownloadAutomatically); Assert.IsFalse(prefs.InstallAutomatically);
            }
        }
        [STATestMethod]
        public void WindowLocalizesAllCulturesAndCanCancelAWaitingJob()
        {
            using (var scope = new UpdateScope())
            {
                var job = UpdateInstallerRunnerTests.Job(scope);
                using (var window = new UpdateWindow())
                {
                    LlmBoundaryScope.Call(window, "CancelPending_Click", null, EventArgs.Empty);
                    Assert.IsTrue(UpdateInstallJob.Load(scope.Root).Completed);
                }
                foreach (var language in UiLanguages.All)
                    using (var culture = new LocalizationScope(language.CultureName))
                    using (var window = new UpdateWindow())
                    { Assert.AreEqual(UiText.Get("VBAi updates"), window.Text); Assert.AreEqual(RightToLeft.No, UiInvoke.Field<TextBox>(window, "notes").RightToLeft); }
            }
        }
        [STATestMethod]
        public void ShownWindowLoadsOwnedReleaseJobsAndVerifiedStagedBytesAndReportsMalformedState()
        {
            foreach (string outcome in new[] { "none", "unmanaged", "notes", "installer", "staged", "pending", "completed", "malformed" })
                using (var fixture = new UpdateWindowFixture())
                {
                    fixture.Managed = outcome != "unmanaged";
                    var release = fixture.Release(outcome != "notes", outcome == "notes" ? null : "Owned notes");
                    if (outcome == "notes" || outcome == "installer" || outcome == "staged") UpdateState.CacheRelease(release);
                    if (outcome == "staged")
                    {
                        string path = UpdatePaths.AssetPath(fixture.Scope.Root, release.Installer.Hash, release.Installer.name);
                        Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "fixture installer", new UTF8Encoding(false));
                    }
                    if (outcome == "pending" || outcome == "completed") { var job = UpdateInstallerRunnerTests.Job(fixture.Scope); job.Completed = outcome == "completed"; job.Save(fixture.Scope.Root); }
                    if (outcome == "malformed") File.WriteAllText(Path.Combine(fixture.Scope.Root, "pending.json"), "owned malformed state");
                    fixture.Window.Show(); Application.DoEvents();
                    Assert.AreEqual(outcome == "staged", UiInvoke.Field<Button>(fixture.Window, "install").Enabled);
                    Assert.AreEqual(outcome == "pending", UiInvoke.Field<Button>(fixture.Window, "cancelPending").Enabled);
                    Assert.AreEqual(UiText.Get(outcome == "malformed" ? "Unable to read the update state." : outcome == "pending" || outcome == "completed" ? "Waiting for VBA hosts to close." : outcome == "notes" || outcome == "installer" || outcome == "staged" ? "An update is available." : outcome == "unmanaged" ? "An installer-managed deployment is required." : "Check for updates to get started."), fixture.Status);
                    if (outcome == "notes") Assert.AreEqual("", UiInvoke.Field<TextBox>(fixture.Window, "notes").Text);
                }
        }

        [STATestMethod]
        public void AppearanceAndDisabledPaintingRespectContrastDirectionAndOwnedGraphics()
        {
            using (var fixture = new UpdateWindowFixture())
            {
                foreach (bool contrast in new[] { false, true })
                {
                    UiTheme.HighContrast = () => contrast; LlmBoundaryScope.Call(fixture.Window, "ApplyAppearance");
                    Assert.AreEqual(contrast ? System.Drawing.SystemColors.Highlight : System.Drawing.Color.FromArgb(37, 99, 235), UiInvoke.Field<Button>(fixture.Window, "check").BackColor);
                }
                UiTheme.HighContrast = () => false; var worker = new Thread(() => LlmBoundaryScope.Call(fixture.Window, "ApplyAppearance")); worker.Start(); Assert.IsTrue(worker.Join(3000)); Application.DoEvents();
                using (var bitmap = new System.Drawing.Bitmap(180, 40))
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                using (var args = new PaintEventArgs(graphics, new System.Drawing.Rectangle(0, 0, 180, 40)))
                using (var box = new CheckBox { Width = 180, Height = 40, Text = "Owned check", BackColor = System.Drawing.Color.Black })
                using (var button = new Button { Width = 180, Height = 40, Text = "Owned button", BackColor = System.Drawing.Color.Black })
                {
                    foreach (bool dark in new[] { false, true })
                        foreach (bool contrast in new[] { false, true })
                            foreach (bool enabled in new[] { false, true })
                            {
                                ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light); UiTheme.HighContrast = () => contrast; box.Enabled = enabled;
                                graphics.Clear(System.Drawing.Color.Black);
                                foreach (RightToLeft rtl in new[] { RightToLeft.No, RightToLeft.Yes })
                                    foreach (bool selected in new[] { false, true })
                                    { box.RightToLeft = rtl; box.Checked = selected; LlmBoundaryScope.Call(fixture.Window, "PaintDisabled", box, args); }
                                button.Enabled = enabled; LlmBoundaryScope.Call(fixture.Window, "PaintDisabled", button, args);
                                Assert.AreEqual(dark && !contrast && !enabled, BitmapContainsPaint(bitmap));
                            }
                }
                fixture.Window.Dispose(); LlmBoundaryScope.Call(fixture.Window, "ApplyAppearance");
            }
        }

        private static bool BitmapContainsPaint(System.Drawing.Bitmap bitmap)
        {
            for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++) if (bitmap.GetPixel(x, y).ToArgb() != System.Drawing.Color.Black.ToArgb()) return true;
            return false;
        }

        [STATestMethod]
        public void PreferencesSkipAndCancellationActionsKeepOwnedFilesAndFailureStatuses()
        {
            using (var fixture = new UpdateWindowFixture())
            {
                UpdateUiPump.Call(fixture.Window, "Skip_Click"); Assert.AreEqual(0, fixture.Stored);
                fixture.Managed = true; fixture.Check(fixture.Release()); UpdateUiPump.Call(fixture.Window, "Skip_Click");
                Assert.AreEqual("999.0.0", fixture.Preferences.SkippedVersion);
                Assert.AreEqual(UiText.Get("This version will be skipped automatically."), fixture.Status);
                fixture.Window.StorePreferences = value => throw new IOException("owned preferences failure");
                UpdateUiPump.Call(fixture.Window, "Skip_Click"); Assert.AreEqual(UiText.Get("Unable to save update preferences."), fixture.Status);
                UpdateUiPump.Call(fixture.Window, "Save_Click"); Assert.AreEqual(UiText.Get("Unable to save update preferences."), fixture.Status);
                UpdateUiPump.Call(fixture.Window, "CancelPending_Click");
                var job = UpdateInstallerRunnerTests.Job(fixture.Scope); job.Completed = true; job.Save(fixture.Scope.Root); UpdateUiPump.Call(fixture.Window, "CancelPending_Click"); Assert.IsTrue(UpdateInstallJob.Load(fixture.Scope.Root).Completed);
                using (File.Open(Path.Combine(fixture.Scope.Root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                { UpdateUiPump.Call(fixture.Window, "CancelPending_Click"); Assert.AreEqual(UiText.Get("Installing update…"), fixture.Status); }
                File.WriteAllText(Path.Combine(fixture.Scope.Root, "pending.json"), "invalid owned job");
                UpdateUiPump.Call(fixture.Window, "CancelPending_Click"); Assert.AreEqual(UiText.Get("Unable to read the update state."), fixture.Status);
                UiInvoke.Field<CheckBox>(fixture.Window, "automaticInstall").Checked = true;
                UiInvoke.Field<CheckBox>(fixture.Window, "automaticDownload").Checked = false;
                Assert.IsFalse(UiInvoke.Field<CheckBox>(fixture.Window, "automaticInstall").Checked);
                LlmBoundaryScope.Call(fixture.Window, "SetBusy", true); UpdateUiPump.Call(fixture.Window, "DownloadPreferenceChanged"); Assert.IsFalse(UiInvoke.Field<CheckBox>(fixture.Window, "automaticInstall").Enabled);
            }
        }

        [STATestMethod]
        public void CheckReportsNoReleaseMissingInstallerAndBoundedNotesWithoutMutatingItsFeedResult()
        {
            using (var fixture = new UpdateWindowFixture())
            {
                foreach (var release in new[] { null, fixture.Release(false), fixture.Release(body: null), fixture.Release(body: new string('x', 64001)) })
                {
                    fixture.Check(release);
                    Assert.AreEqual(UiText.Get(release == null ? "VBAi is up to date." : release.Installer == null ? "This release has no compatible verified installer." : "An update is available."), fixture.Status);
                    Assert.IsTrue(UiInvoke.Field<TextBox>(fixture.Window, "notes").Text.Length <= 64000);
                    if (release?.body?.Length > 64000)
                    {
                        Assert.AreEqual(64001, UiInvoke.Field<UpdateRelease>(fixture.Window, "release").body.Length);
                        Assert.AreEqual(64000, new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<UpdateRelease>(File.ReadAllText(Path.Combine(fixture.Scope.Root, "latest-release.json"))).body.Length);
                    }
                }
                fixture.Window.CreateFeed = () => throw new IOException("owned feed unavailable");
                UpdateUiPump.Call(fixture.Window, "Check_Click"); Assert.AreEqual(UiText.Get("Unable to check for updates. Check the GitHub account or retry later."), fixture.Status);
            }
        }

        [STATestMethod]
        public void DownloadAndInstallUseOwnedPayloadAndExactBackgroundScheduleThenReportFailures()
        {
            using (var fixture = new UpdateWindowFixture())
            {
                UpdateUiPump.Call(fixture.Window, "Download_Click"); UpdateUiPump.Call(fixture.Window, "Install_Click");
                fixture.Check(fixture.Release(false)); UpdateUiPump.Call(fixture.Window, "Download_Click"); Assert.IsFalse(UiInvoke.Field<bool>(fixture.Window, "busy"));
                fixture.Managed = true; fixture.Check(fixture.Release()); fixture.Download();
                Assert.AreEqual(UiText.Get("Update downloaded and verified."), fixture.Status); Assert.IsNotNull(UiInvoke.Field<string>(fixture.Window, "downloaded"));
                int scheduled = 0;
                fixture.Window.Schedule = (release, path, background) => { Assert.AreSame(UiInvoke.Field<UpdateRelease>(fixture.Window, "release"), release); Assert.IsFalse(background); Assert.AreEqual(release.Installer.Hash, UpdatePaths.Hash(path)); scheduled++; };
                UpdateUiPump.Call(fixture.Window, "Install_Click"); UpdateUiPump.Until(() => !UiInvoke.Field<bool>(fixture.Window, "busy"));
                Assert.AreEqual(1, scheduled); Assert.AreEqual(UiText.Get("Waiting for VBA hosts to close."), fixture.Status); Assert.IsTrue(UiInvoke.Field<Button>(fixture.Window, "cancelPending").Enabled); Assert.IsFalse(UiInvoke.Field<Button>(fixture.Window, "install").Enabled);
                fixture.Window.Schedule = (release, path, background) => throw new IOException("owned scheduling refusal");
                UpdateUiPump.Call(fixture.Window, "Install_Click"); UpdateUiPump.Until(() => !UiInvoke.Field<bool>(fixture.Window, "busy")); Assert.AreEqual(UiText.Get("Unable to schedule the update."), fixture.Status);
                LlmBoundaryScope.Set(fixture.Window, "release", null); UpdateUiPump.Call(fixture.Window, "Install_Click"); Assert.IsFalse(UiInvoke.Field<bool>(fixture.Window, "busy"));
            }
            using (var fixture = new UpdateWindowFixture())
            {
                fixture.Check(fixture.Release()); fixture.Window.CreateFeed = () => throw new IOException("owned download refusal");
                UpdateUiPump.Call(fixture.Window, "Download_Click"); Assert.AreEqual(UiText.Get("Unable to download or verify the installer."), fixture.Status);
            }
        }

        [STATestMethod]
        public void AsyncCheckDownloadAndInstallIgnoreTheirResultsAfterOwnedWindowDisposal()
        {
            foreach (string operation in new[] { "check", "download", "install" })
                foreach (bool failed in new[] { false, true })
                    using (var fixture = new UpdateWindowFixture())
                    {
                        Action complete;
                        if (operation == "check")
                        {
                            var pending = new TaskCompletionSource<UpdateRelease>(); fixture.Window.CheckRelease = (feed, version, previews, token) => pending.Task;
                            complete = () => { if (failed) pending.SetException(new IOException("owned late error")); else pending.SetResult(fixture.Release()); };
                        }
                        else
                        {
                            fixture.Managed = true; fixture.Check(fixture.Release());
                            if (operation == "download")
                            {
                                var pending = new TaskCompletionSource<string>(); fixture.Window.DownloadInstaller = (feed, asset, root, progress, token) => pending.Task;
                                complete = () => { if (failed) pending.SetException(new IOException("owned late error")); else pending.SetResult("owned-staged.exe"); };
                            }
                            else
                            {
                                fixture.Download(); var pending = new TaskCompletionSource<bool>();
                                fixture.Window.RunBackground = action => pending.Task;
                                complete = () => { if (failed) pending.SetException(new IOException("owned late error")); else pending.SetResult(true); };
                            }
                        }
                        string method = operation == "check" ? "Check_Click" : operation == "download" ? "Download_Click" : "Install_Click";
                        UpdateUiPump.Call(fixture.Window, method); UpdateUiPump.Call(fixture.Window, method); Assert.IsTrue(UiInvoke.Field<bool>(fixture.Window, "busy"));
                        fixture.Window.Dispose(); complete(); Application.DoEvents(); Application.DoEvents();
                        Assert.IsTrue(fixture.Window.IsDisposed); Assert.IsTrue(UiInvoke.Field<bool>(fixture.Window, "busy"), "A disposed window must not be re-enabled by a late continuation.");
                    }
        }

        [STATestMethod]
        public void DownloadProgressOnlyUpdatesALiveOwnedWindow()
        {
            foreach (bool disposed in new[] { false, true })
                using (var fixture = new UpdateWindowFixture())
                {
                    fixture.Check(fixture.Release()); var pending = new TaskCompletionSource<string>(); IProgress<int> reporter = null;
                    fixture.Window.DownloadInstaller = (feed, asset, root, progress, token) => { reporter = progress; return pending.Task; };
                    UpdateUiPump.Call(fixture.Window, "Download_Click"); Assert.IsNotNull(reporter);
                    if (disposed) fixture.Window.Dispose(); reporter.Report(50); Application.DoEvents(); Application.DoEvents();
                    Assert.AreEqual(disposed ? 0 : 50, UiInvoke.Field<ProgressBar>(fixture.Window, "progress").Value);
                    pending.SetResult("owned-staged.exe"); Application.DoEvents(); Application.DoEvents();
                }
        }

        [STATestMethod]
        public void SavingPreferencesRejectsStaleInstallChecksWhenDownloadsAreDisabled()
        {
            foreach (bool managed in new[] { false, true })
                foreach (bool download in new[] { false, true })
                    foreach (bool install in new[] { false, true })
                        using (var fixture = new UpdateWindowFixture())
                        {
                            fixture.Managed = managed;
                            UiInvoke.Field<CheckBox>(fixture.Window, "automaticDownload").Checked = download;
                            UiInvoke.Field<CheckBox>(fixture.Window, "automaticInstall").Checked = install;
                            UpdateUiPump.Call(fixture.Window, "Save_Click");
                            Assert.AreEqual(managed && download && install, fixture.Preferences.InstallAutomatically);
                            Assert.AreEqual(download, fixture.Preferences.DownloadAutomatically); Assert.AreEqual(1, fixture.Stored);
                        }
        }

        [STATestMethod]
        public void UpdateDesignerAndOwnerFallbackPreserveCloseCancellationAndOwnedResourceDisposal()
        {
            var context = System.ComponentModel.LicenseManager.CurrentContext;
            try
            {
                System.ComponentModel.LicenseManager.CurrentContext = new DesignContext();
                using (var window = new UpdateWindow()) UpdateUiPump.Call(window, "DownloadPreferenceChanged");
            }
            finally { System.ComponentModel.LicenseManager.CurrentContext = context; }
            using (var fixture = new UpdateWindowFixture())
            {
                var cancellation = UiInvoke.Field<CancellationTokenSource>(fixture.Window, "cancellation");
                LlmBoundaryScope.Call(fixture.Window, "OnFormClosing", new FormClosingEventArgs(CloseReason.UserClosing, false)); Assert.IsTrue(cancellation.IsCancellationRequested);
            }
            using (var window = new UpdateWindow()) { UiInvoke.Call(typeof(UpdateWindow), "Dispose", window, false); Assert.IsFalse(window.IsDisposed); Assert.IsNotNull(UiInvoke.Field<System.ComponentModel.IContainer>(window, "components")); }
            using (var window = new UpdateWindow())
            { UiInvoke.Field<System.ComponentModel.IContainer>(window, "components").Dispose(); LlmBoundaryScope.Set(window, "components", null); window.Dispose(); Assert.IsTrue(window.IsDisposed); }
            var modal = AddIn.ShowModal; int shown = 0;
            try
            {
                AddIn.ShowModal = (form, owner) => { shown++; Assert.AreEqual(shown == 1 ? new IntPtr(123) : IntPtr.Zero, owner?.Handle ?? IntPtr.Zero); return DialogResult.Cancel; };
                UpdateWindow.ShowForVbe(new CrashReportWindowTests.Host()); UpdateWindow.ShowForVbe(null); UpdateWindow.ShowForVbe(new object()); Assert.AreEqual(3, shown);
            }
            finally { AddIn.ShowModal = modal; }
        }
    }
}
