using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class UpdateInstallerRunnerTests
    {
        internal static UpdateInstallJob Job(UpdateScope scope)
        {
            var asset = UpdateFeedTests.Asset("fixture installer"); string path = UpdatePaths.AssetPath(scope.Root, asset.Hash, asset.name);
            Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "fixture installer", new UTF8Encoding(false));
            var job = new UpdateInstallJob { InstallerPath = path, Sha256 = asset.Hash, TargetVersion = "1.2.3", InstallationDirectory = UpdateState.InstallationDirectory };
            job.Save(scope.Root); return job;
        }
        [TestMethod]
        public void UnapprovedPublisherIsFinalBeforeWaitingForLiveHosts()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); UpdateState.RecordHost();
                int hostChecks = 0, installs = 0;
                var runner = new UpdateInstallerRunner(scope.Root)
                {
                    VerifyPublisher = path => false,
                    IsAlive = lease => { hostChecks++; return true; },
                    Install = path => { installs++; return 0; }
                };
                Assert.IsTrue(runner.Tick(job));
                Assert.IsTrue(job.Completed); Assert.IsFalse(job.Succeeded);
                Assert.AreEqual(0, hostChecks); Assert.AreEqual(0, installs);
                Assert.IsFalse(File.Exists(Path.Combine(scope.Root, "installation.started")));
                Assert.IsFalse(UpdateInstallJob.Load(scope.Root).Succeeded);
            }
        }

        [TestMethod]
        public void RunnerWaitsForRegisteredHostsAndVerifiesAgainBeforeExactlyOneInstallation()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); UpdateState.RecordHost(); int installed = 0;
                var runner = new UpdateInstallerRunner(scope.Root) { InstalledVersion = directory => "1.2.3", IsAlive = lease => true, VerifySignature = path => true, VerifyPublisher = path => true, Install = path => { installed++; return 0; } };
                Assert.IsFalse(runner.Tick(job)); Assert.AreEqual(0, installed);
                runner.IsAlive = lease => false; Assert.IsTrue(runner.Tick(job)); Assert.AreEqual(1, installed); Assert.IsTrue(job.Completed);
                Assert.IsFalse(File.Exists(Path.Combine(scope.Root, "installation.started")));
                Assert.IsTrue(runner.Tick(job)); Assert.AreEqual(1, installed); Assert.AreEqual("Update installed. Restart the VBA host.", UpdateInstallJob.Load(scope.Root).Status);
            }
        }
        [TestMethod]
        public void RunnerHonorsCancellationChecksumSignatureAndInstallerFailuresWithoutStartingRealProcesses()
        {
            foreach (string scenario in new[] { "cancel", "hash", "signature", "failed", "reboot" })
                using (var scope = new UpdateScope())
                {
                    var job = Job(scope); int installed = 0;
                    var runner = new UpdateInstallerRunner(scope.Root) { VerifySignature = path => scenario != "signature", VerifyPublisher = path => true, Install = path => { installed++; return scenario == "reboot" ? 3010 : 1603; } };
                    if (scenario == "cancel") { var cancelled = UpdateInstallJob.Load(scope.Root); cancelled.Completed = true; cancelled.Status = "Update cancelled."; cancelled.Save(scope.Root); }
                    if (scenario == "hash") File.WriteAllText(job.InstallerPath, "changed");
                    Assert.IsTrue(runner.Tick(job)); Assert.IsTrue(job.Completed);
                    Assert.AreEqual(scenario == "failed" || scenario == "reboot" ? 1 : 0, installed);
                    Assert.AreEqual(scenario == "reboot", job.RestartRequired);
                }
        }
        [TestMethod]
        public void NativeProcessIdentityAndWindowsSignatureChecksAreReadOnly()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); UpdateState.RecordHost(); var runner = new UpdateInstallerRunner(scope.Root);
                Assert.IsTrue(runner.HostsOpen(job)); Assert.IsFalse(runner.VerifySignature(job.InstallerPath));
                Assert.IsTrue(runner.VerifySignature(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "dotnet.exe")));
            }
        }
        [TestMethod]
        public void NativeInstallerArgumentsAreObservedWhileOnlyControlledExitHelpersRun()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var runner = new UpdateInstallerRunner(fixture.Scope.Root);
                UpdateInstallerRunner.StartProcess = start => null;
                Assert.ThrowsException<InvalidOperationException>(() => runner.Install("owned-installer.exe"));
                foreach (string path in new[] { "owned-installer.exe", "owned-installer.MSI" })
                {
                    UpdateInstallerRunner.StartProcess = start =>
                    {
                        bool msi = path.EndsWith(".MSI", StringComparison.Ordinal);
                        Assert.AreEqual(msi ? "msiexec.exe" : path, Path.GetFileName(start.FileName));
                        Assert.AreEqual(msi ? "/i \"" + path + "\" /quiet /norestart /L*v \"" + path + ".install.log\"" : "/update /quiet /norestart /VERYSILENT /SUPPRESSMSGBOXES /SP-", start.Arguments);
                        Assert.IsTrue(start.UseShellExecute); Assert.AreEqual(ProcessWindowStyle.Hidden, start.WindowStyle);
                        return UpdatesNativeFixture.ExitHelper(7);
                    };
                    Assert.AreEqual(7, runner.Install(path));
                }
            }
        }

        [TestMethod]
        public void NativeHostIdentityRejectsMissingReusedExitedAndUnassociatedProcessObjects()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var runner = new UpdateInstallerRunner(fixture.Scope.Root);
                using (var current = Process.GetCurrentProcess())
                {
                    var lease = new UpdateHostLease { Pid = current.Id, StartTimeUtcTicks = current.StartTime.ToUniversalTime().Ticks };
                    Assert.IsTrue(runner.IsAlive(lease)); lease.StartTimeUtcTicks++; Assert.IsFalse(runner.IsAlive(lease));
                }
                Assert.IsFalse(runner.IsAlive(new UpdateHostLease { Pid = int.MaxValue }));
                UpdateInstallerRunner.ReadProcess = id => new Process();
                Assert.IsFalse(runner.IsAlive(new UpdateHostLease { Pid = 1 }));
                var exited = UpdatesNativeFixture.ExitHelper(0); exited.WaitForExit();
                UpdateInstallerRunner.ReadProcess = id => exited;
                Assert.IsFalse(runner.IsAlive(new UpdateHostLease { Pid = 1 }));
            }
        }

        [TestMethod]
        public void HostLeaseGuardsWaitOnUncertainRecordsAndDiscardOnlyConfirmedDeadForeignHosts()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var job = Job(fixture.Scope); var runner = new UpdateInstallerRunner(fixture.Scope.Root) { IsAlive = lease => true };
                string hosts = Path.Combine(fixture.Scope.Root, "hosts"); Directory.CreateDirectory(hosts);
                string path = Path.Combine(hosts, "owned.json");
                foreach (string value in new[] { "null", "{}", "{\"Pid\":1}", "{\"Pid\":1,\"StartTimeUtcTicks\":1}" })
                { File.WriteAllText(path, value); Assert.IsTrue(runner.HostsOpen(job)); }
                var foreign = new UpdateHostLease { Pid = 1, StartTimeUtcTicks = 1, InstallationDirectory = Path.Combine(fixture.Scope.Root, "foreign") };
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(foreign));
                Assert.IsFalse(runner.HostsOpen(job)); Assert.IsTrue(File.Exists(path));
                runner.IsAlive = lease => false; Assert.IsFalse(runner.HostsOpen(job)); Assert.IsFalse(File.Exists(path));
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(foreign)); runner.IsAlive = lease => throw new UnauthorizedAccessException("owned inaccessible process");
                Assert.IsTrue(runner.HostsOpen(job)); Assert.IsTrue(File.Exists(path));
            }
        }

        [TestMethod]
        public void InstallerTickHonorsLockChangedPendingIdentityAndBothSuccessfulRestartCodes()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var job = Job(fixture.Scope);
                var runner = new UpdateInstallerRunner(fixture.Scope.Root) { VerifySignature = path => true, VerifyPublisher = path => true, Install = path => 0, InstalledVersion = path => "1.2.3" };
                using (File.Open(Path.Combine(fixture.Scope.Root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)) Assert.IsFalse(runner.Tick(job));
                var next = Job(fixture.Scope); next.TargetVersion = "1.2.4"; next.Save(fixture.Scope.Root);
                Assert.IsTrue(runner.Tick(job)); Assert.IsFalse(job.Completed);
                next.TargetVersion = "1.2.3"; next.Sha256 = new string('a', 64); next.InstallerPath = UpdatePaths.AssetPath(fixture.Scope.Root, next.Sha256, "VBAi-Setup-win-x64.exe"); next.Save(fixture.Scope.Root);
                Assert.IsTrue(runner.Tick(job)); Assert.IsFalse(job.Completed);
            }
            foreach (string outcome in new[] { "1641", "0", "invalid", "mismatch", "missing", "no-pending", "host-reopened" })
                using (var fixture = new UpdatesNativeFixture())
                {
                    var job = Job(fixture.Scope);
                    var runner = new UpdateInstallerRunner(fixture.Scope.Root) { VerifySignature = path => true, VerifyPublisher = path => true, Install = path => outcome == "1641" ? 1641 : 0, InstalledVersion = path => outcome == "invalid" ? "invalid" : outcome == "mismatch" ? "1.2.2" : "1.2.3" };
                    if (outcome == "missing") File.Delete(job.InstallerPath);
                    if (outcome == "no-pending") File.Delete(Path.Combine(fixture.Scope.Root, "pending.json"));
                    if (outcome == "host-reopened") { runner.IsAlive = lease => true; runner.VerifySignature = path => { UpdateState.RecordHost(); return true; }; }
                    Assert.AreEqual(outcome != "host-reopened", runner.Tick(job));
                    Assert.AreEqual(outcome != "host-reopened", job.Completed);
                    if (outcome != "host-reopened") Assert.AreEqual(outcome == "1641" ? "Update installed. Restart Windows to finish." : outcome == "0" || outcome == "no-pending" ? "Update installed. Restart the VBA host." : "Installation failed. Check the installer log.", job.Status);
                }
        }

        [TestMethod]
        public void PublisherRefusalPersistsEvenWhenWindowsTrustWouldAcceptThePackage()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); int signatureChecks = 0; int installs = 0;
                var runner = new UpdateInstallerRunner(scope.Root)
                {
                    VerifySignature = path => { signatureChecks++; return true; },
                    Install = path => { installs++; return 0; }
                };
                Assert.IsTrue(runner.Tick(job));
                Assert.AreEqual(0, signatureChecks);
                Assert.AreEqual(0, installs);
                Assert.AreEqual("Installation refused: no trusted VBAi publisher is configured or the publisher does not match.", job.Status);
                Assert.AreEqual(job.Status, UpdateInstallJob.Load(scope.Root).Status);
                Assert.IsTrue(UpdateInstallJob.Load(scope.Root).Completed);
                Assert.IsFalse(File.Exists(Path.Combine(scope.Root, "installation.started")));
            }
        }

        [TestMethod]
        public void UnknownInstallerExceptionLeavesAttemptMarkerAndBlocksASecondJob()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); int installs = 0;
                var runner = new UpdateInstallerRunner(scope.Root)
                {
                    VerifyPublisher = path => true,
                    VerifySignature = path => true,
                    Install = path => { installs++; throw new IOException("owned installer outcome unknown"); }
                };
                Assert.IsTrue(runner.Tick(job));
                Assert.AreEqual(1, installs);
                Assert.IsTrue(File.Exists(Path.Combine(scope.Root, "installation.started")));
                Assert.AreEqual(UpdateInstallerRunner.UncertainStatus, job.Status);
                Assert.AreEqual(job.Status, UpdateInstallJob.Load(scope.Root).Status);

                var second = Job(scope);
                Assert.IsTrue(runner.Tick(second));
                Assert.AreEqual(1, installs);
                Assert.AreEqual(UpdateInstallerRunner.UncertainStatus, second.Status);
                Assert.AreEqual(second.Status, UpdateInstallJob.Load(scope.Root).Status);
                Assert.IsTrue(File.Exists(Path.Combine(scope.Root, "installation.started")));
            }
        }

        [TestMethod]
        public void NativeInstallerTimeoutIsBoundedLeavesOwnedProcessRunningAndBlocksReplay()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var job = Job(fixture.Scope); int processId = 0; int starts = 0; int waits = 0;
                var originalStartProcess = UpdateInstallerRunner.StartProcess;
                var originalWaitForInstaller = UpdateInstallerRunner.WaitForInstaller;
                UpdateInstallerRunner.StartProcess = start =>
                {
                    starts++;
                    var helper = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"), "/d /c ping -n 8 127.0.0.1 >nul")
                    { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
                    processId = helper.Id;
                    return helper;
                };
                UpdateInstallerRunner.WaitForInstaller = (process, milliseconds) =>
                {
                    waits++;
                    Assert.AreEqual(UpdateInstallerRunner.InstallerTimeoutMilliseconds, milliseconds);
                    Assert.AreEqual(processId, process.Id);
                    return false;
                };
                var runner = new UpdateInstallerRunner(fixture.Scope.Root)
                { VerifyPublisher = path => true, VerifySignature = path => true };
                try
                {
                    Assert.IsTrue(runner.Tick(job));
                    Assert.AreEqual(1, starts);
                    Assert.AreEqual(1, waits);
                    Assert.AreEqual(UpdateInstallerRunner.UncertainStatus, job.Status);
                    Assert.AreEqual(job.Status, UpdateInstallJob.Load(fixture.Scope.Root).Status);
                    Assert.IsTrue(File.Exists(Path.Combine(fixture.Scope.Root, "installation.started")));
                    using (var helper = Process.GetProcessById(processId))
                        Assert.IsFalse(helper.HasExited, "The runner must not kill a process with an uncertain outcome.");

                    var second = Job(fixture.Scope);
                    Assert.IsTrue(runner.Tick(second));
                    Assert.AreEqual(1, starts);
                    Assert.AreEqual(1, waits);
                    Assert.AreEqual(UpdateInstallerRunner.UncertainStatus, second.Status);
                    Assert.AreEqual(second.Status, UpdateInstallJob.Load(fixture.Scope.Root).Status);
                }
                finally
                {
                    UpdateInstallerRunner.StartProcess = originalStartProcess;
                    UpdateInstallerRunner.WaitForInstaller = originalWaitForInstaller;
                    if (processId != 0)
                        try
                        {
                            using (var helper = Process.GetProcessById(processId))
                                Assert.IsTrue(helper.WaitForExit(15000), "The owned helper did not exit naturally.");
                        }
                        catch (ArgumentException) { } // It already exited on its own.
                }
            }
        }
    }
}
