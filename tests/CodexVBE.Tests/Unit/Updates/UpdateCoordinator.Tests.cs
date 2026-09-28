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
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class UpdateCoordinatorTests
    {
        [TestMethod]
        public async Task AutomaticChecksRespectPreferencesAndScheduleOnlyManagedDeployments()
        {
            using (var scope = new UpdateScope())
            {
                scope.Managed(); var asset = UpdateFeedTests.Asset("fixture installer");
                string release = new JavaScriptSerializer().Serialize(new[] { new { tag_name = "v1.2.3", assets = new[] { asset } } });
                var handler = new UpdateFeedTests.Handler(release, "fixture installer"); int launched = 0;
                UpdateCoordinator.CreateFeed = () => new UpdateFeed(handler, ct => throw new AssertFailedException());
                UpdateCoordinator.LaunchWorker = background => { Assert.IsTrue(background); launched++; };
                var p = UpdateState.Load(); p.InstallAutomatically = true; UpdateState.Save(p);
                var result = await UpdateCoordinator.Check(true, null, CancellationToken.None);
                Assert.AreEqual("1.2.3", result.Version.Text); Assert.AreEqual(1, launched);
                Assert.IsNotNull(UpdateInstallJob.Load(scope.Root)); Assert.IsNull(await UpdateCoordinator.Check(true, null, CancellationToken.None)); Assert.AreEqual(1, launched);
                p = UpdateState.Load(); p.LastCheckUtc = DateTime.MinValue; p.CheckAutomatically = false; UpdateState.Save(p);
                Assert.IsNull(await UpdateCoordinator.Check(true, null, CancellationToken.None));
            }
        }
        [TestMethod]
        public void SchedulingRejectsDeveloperCheckoutsAndDuplicatePendingInstalls()
        {
            using (var scope = new UpdateScope())
            {
                var asset = UpdateFeedTests.Asset("fixture installer"); var release = new UpdateRelease { tag_name = "1.2.3", assets = new[] { asset } };
                string path = UpdatePaths.AssetPath(scope.Root, asset.Hash, asset.name);
                Assert.ThrowsException<InvalidOperationException>(() => UpdateCoordinator.Schedule(release, path, false));
                scope.Managed(); Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "fixture installer", new UTF8Encoding(false)); UpdateCoordinator.LaunchWorker = background => { };
                UpdateCoordinator.Schedule(release, path, false);
                Assert.ThrowsException<InvalidOperationException>(() => UpdateCoordinator.Schedule(release, path, false));
            }
        }
        [TestMethod]
        public void NativeWorkerCopiesVerifiedOwnedBytesAndPreservesBackgroundProcessArguments()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                Assert.ThrowsException<FileNotFoundException>(() => UpdatesNativeFixture.Invoke<object>(typeof(UpdateCoordinator), "LaunchWorkerNative", true));
                string source = Path.Combine(UpdateState.InstallationDirectory, "VBAi.Updater.exe");
                File.WriteAllText(source, "owned synthetic worker bytes"); File.WriteAllText(source + ".config", "owned config");
                foreach (bool background in new[] { false, true })
                    UpdatesNativeFixture.Invoke<object>(typeof(UpdateCoordinator), "LaunchWorkerNative", background);
                Assert.AreEqual(2, fixture.Starts.Count);
                Assert.AreEqual("", fixture.Starts[0].Arguments); Assert.AreEqual(ProcessWindowStyle.Normal, fixture.Starts[0].WindowStyle);
                Assert.AreEqual("--background", fixture.Starts[1].Arguments); Assert.AreEqual(ProcessWindowStyle.Hidden, fixture.Starts[1].WindowStyle);
                Assert.IsTrue(fixture.Starts.All(start => start.UseShellExecute));
                string worker = fixture.Starts[0].FileName; Assert.AreEqual(UpdatePaths.Hash(source), UpdatePaths.Hash(worker));
                Assert.AreEqual("owned config", File.ReadAllText(worker + ".config"));
                File.WriteAllText(worker, "changed cached worker");
                Assert.ThrowsException<InvalidDataException>(() => UpdatesNativeFixture.Invoke<object>(typeof(UpdateCoordinator), "LaunchWorkerNative", true));
                Assert.AreEqual(2, fixture.Starts.Count);
            }
            using (var fixture = new UpdatesNativeFixture())
            {
                File.WriteAllText(Path.Combine(UpdateState.InstallationDirectory, "VBAi.Updater.exe"), "owned worker without config");
                UpdatesNativeFixture.Invoke<object>(typeof(UpdateCoordinator), "LaunchWorkerNative", true);
                Assert.IsFalse(File.Exists(fixture.Starts.Single().FileName + ".config"));
            }
        }

        [TestMethod]
        public void StartupCreatesOneOwnedTimerReplaysPendingWorkAndStopsIdempotently()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                UpdateCoordinator.Start(); Assert.AreEqual(0, fixture.Callbacks.Count);
                fixture.Scope.Managed(); var preferences = UpdateState.Load(); preferences.CheckAutomatically = false; UpdateState.Save(preferences);
                var job = UpdateInstallerRunnerTests.Job(fixture.Scope); int launches = 0;
                UpdateCoordinator.LaunchWorker = background => { Assert.IsTrue(background); launches++; };
                UpdateCoordinator.Start(); UpdateCoordinator.Start(); Assert.AreEqual(2, launches); Assert.AreEqual(1, fixture.Callbacks.Count);
                fixture.Callbacks[0](null); Assert.AreEqual(0, fixture.Logs.Count);
                job.Completed = true; job.Save(fixture.Scope.Root); UpdateCoordinator.Start(); Assert.AreEqual(2, launches);
                UpdateCoordinator.Stop(); UpdateCoordinator.Stop(); Assert.IsNull(UpdatesNativeFixture.Read("timer")); Assert.IsNull(UpdatesNativeFixture.Read("lifetime"));
            }
            using (var fixture = new UpdatesNativeFixture())
            {
                fixture.Scope.Managed(); UpdateCoordinator.CreateTimer = (callback, state, due, period) => throw new InvalidOperationException("owned timer refusal");
                UpdateCoordinator.Start(); Assert.AreEqual(1, fixture.Logs.Count); StringAssert.Contains(fixture.Logs[0], "Update startup is unavailable.");
            }
        }

        [TestMethod]
        public async Task AutomaticRunnerDistinguishesOwnedCancellationFromOtherFailures()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                UpdateCoordinator.CreateFeed = () => throw new IOException("owned feed unavailable");
                await UpdatesNativeFixture.Invoke<Task>(typeof(UpdateCoordinator), "RunAutomatic", CancellationToken.None);
                Assert.AreEqual(1, fixture.Logs.Count);
                var preferences = UpdateState.Load(); preferences.LastAttemptUtc = DateTime.MinValue; UpdateState.Save(preferences);
                UpdateCoordinator.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler("[]"), token => throw new AssertFailedException("Unexpected credentials"));
                using (var cancellation = new CancellationTokenSource())
                {
                    cancellation.Cancel(); await UpdatesNativeFixture.Invoke<Task>(typeof(UpdateCoordinator), "RunAutomatic", cancellation.Token);
                    Assert.AreEqual(1, fixture.Logs.Count);
                }
                preferences = UpdateState.Load(); preferences.LastAttemptUtc = DateTime.MinValue; UpdateState.Save(preferences);
                await UpdatesNativeFixture.Invoke<Task>(typeof(UpdateCoordinator), "RunAutomatic", CancellationToken.None); Assert.AreEqual(1, fixture.Logs.Count);
            }
        }

        [TestMethod]
        public async Task CoordinatorChecksRespectEachPreferenceAndPriorCancellationIdentity()
        {
            foreach (string outcome in new[] { "manual", "no-release", "no-installer", "no-download", "cancel", "uncertain", "other-status", "other-version", "pending", "no-check", "no-install", "unmanaged" })
            using (var fixture = new UpdatesNativeFixture())
            {
                if (outcome != "unmanaged") fixture.Scope.Managed();
                const string bytes = "fixture installer"; var asset = UpdateFeedTests.Asset(bytes);
                var release = new UpdateRelease { tag_name = "1.2.3", assets = outcome == "no-installer" ? null : new[] { asset } };
                var handler = new UpdateFeedTests.Handler(outcome == "no-release" ? "[]" : new JavaScriptSerializer().Serialize(new[] { release }), bytes);
                UpdateCoordinator.CreateFeed = () => new UpdateFeed(handler, token => throw new AssertFailedException("Unexpected credentials"));
                int launches = 0; UpdateCoordinator.LaunchWorker = background => launches++;
                var preferences = new UpdatePreferences { InstallAutomatically = outcome != "no-install", DownloadAutomatically = outcome != "no-download" };
                UpdateState.Save(preferences);
                if (new[] { "cancel", "uncertain", "other-status", "other-version", "pending" }.Contains(outcome))
                {
                    var previous = UpdateInstallerRunnerTests.Job(fixture.Scope);
                    previous.Completed = outcome != "pending"; previous.Status = outcome == "uncertain" ? "Installation status is uncertain. Check the installed version." : outcome == "other-status" ? "Completed fixture" : "Update cancelled.";
                    if (outcome == "other-version") previous.TargetVersion = "1.2.2";
                    previous.Save(fixture.Scope.Root);
                }
                if (outcome == "no-check") UpdateCoordinator.CreateFeed = () =>
                {
                    var current = UpdateState.Load(); current.CheckAutomatically = false; UpdateState.Save(current);
                    return new UpdateFeed(handler, token => throw new AssertFailedException("Unexpected credentials"));
                };
                var result = await UpdateCoordinator.Check(outcome != "manual", null, CancellationToken.None);
                Assert.AreEqual(outcome == "other-status" || outcome == "other-version" ? 1 : 0, launches, outcome);
                Assert.AreEqual(outcome != "no-release", result != null);
                Assert.IsTrue(UpdateState.Load().LastCheckUtc > DateTime.MinValue);
            }
            using (var fixture = new UpdatesNativeFixture())
            using (File.Open(Path.Combine(fixture.Scope.Root, "operation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => UpdateCoordinator.Check(false, null, CancellationToken.None));
        }

        [TestMethod]
        public void ScheduleRejectsEveryUnverifiedPathAndMissingOrChangedOwnedInstaller()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                fixture.Scope.Managed();
                var asset = UpdateFeedTests.Asset("fixture installer"); var release = new UpdateRelease { tag_name = "1.2.3", assets = new[] { asset } };
                string expected = UpdatePaths.AssetPath(fixture.Scope.Root, asset.Hash, asset.name);
                foreach (var invalid in new UpdateRelease[] { null, new UpdateRelease { tag_name = "1.2.3" }, new UpdateRelease { tag_name = "1.2.3", assets = new[] { new UpdateAsset { name = asset.name } } } })
                    Assert.ThrowsException<InvalidDataException>(() => UpdateCoordinator.Schedule(invalid, expected, false));
                Assert.ThrowsException<InvalidDataException>(() => UpdateCoordinator.Schedule(release, Path.Combine(fixture.Scope.Root, "foreign.exe"), false));
                Assert.ThrowsException<InvalidDataException>(() => UpdateCoordinator.Schedule(release, expected, false));
                Directory.CreateDirectory(Path.GetDirectoryName(expected)); File.WriteAllText(expected, "changed");
                Assert.ThrowsException<InvalidDataException>(() => UpdateCoordinator.Schedule(release, expected, false));
            }
        }
    }
}
