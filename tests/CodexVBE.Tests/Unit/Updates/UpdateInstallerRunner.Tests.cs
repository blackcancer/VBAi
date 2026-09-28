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
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateInstallerRunnerTests
    {
        internal static UpdateInstallJob Job(UpdateScope scope)
        {
            var asset = UpdateFeedTests.Asset("fixture installer"); string path = UpdatePaths.AssetPath(scope.Root, asset.Hash, asset.name);
            Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "fixture installer", new UTF8Encoding(false));
            var job = new UpdateInstallJob { InstallerPath = path, Sha256 = asset.Hash, TargetVersion = "1.2.3", InstallationDirectory = UpdateState.InstallationDirectory };
            job.Save(scope.Root); return job;
        }
        [TestMethod]
        public void RunnerWaitsForRegisteredHostsAndVerifiesAgainBeforeExactlyOneInstallation()
        {
            using (var scope = new UpdateScope())
            {
                var job = Job(scope); UpdateState.RecordHost(); int installed = 0;
                var runner = new UpdateInstallerRunner(scope.Root) { InstalledVersion = directory => "1.2.3", IsAlive = lease => true, VerifySignature = path => true, Install = path => { installed++; return 0; } };
                Assert.IsFalse(runner.Tick(job)); Assert.AreEqual(0, installed);
                runner.IsAlive = lease => false; Assert.IsTrue(runner.Tick(job)); Assert.AreEqual(1, installed); Assert.IsTrue(job.Completed);
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
                var runner = new UpdateInstallerRunner(scope.Root) { VerifySignature = path => scenario != "signature", Install = path => { installed++; return scenario == "reboot" ? 3010 : 1603; } };
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
    }
}
