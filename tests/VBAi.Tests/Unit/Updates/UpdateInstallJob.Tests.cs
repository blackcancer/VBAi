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
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateInstallJobTests
    {
        [TestMethod]
        public void PendingJobsRoundTripButCannotExecuteAnArbitraryLocalPath()
        {
            using (var scope = new UpdateScope())
            {
                var job = new UpdateInstallJob { Sha256 = new string('a', 64), TargetVersion = "1.2.3", InstallationDirectory = UpdateState.InstallationDirectory };
                job.InstallerPath = UpdatePaths.AssetPath(scope.Root, job.Sha256, "VBAi-Setup-win-x64.exe"); job.Save(scope.Root);
                Assert.AreEqual(job.InstallerPath, UpdateInstallJob.Load(scope.Root).InstallerPath);
                job.InstallerPath = Path.Combine(scope.Root, "other.exe"); Assert.ThrowsException<InvalidDataException>(() => job.Validate(scope.Root));
                File.WriteAllText(Path.Combine(scope.Root, "pending.json"), new string('x', 16385)); Assert.ThrowsException<InvalidDataException>(() => UpdateInstallJob.Load(scope.Root));
            }
        }
        [TestMethod]
        public void PendingValidationRejectsEveryUnsafeOwnedContractBeforeExecution()
        {
            using (var scope = new UpdateScope())
            {
                Assert.IsNull(UpdateInstallJob.Load(scope.Root));
                foreach (string kind in new[] { "hash", "version", "installation-null", "installation-empty", "installation-relative", "path-null", "path-empty", "path-other-root" })
                {
                    var job = UpdateInstallerRunnerTests.Job(scope);
                    switch (kind)
                    {
                        case "hash": job.Sha256 = null; break;
                        case "version": job.TargetVersion = "invalid"; break;
                        case "installation-null": job.InstallationDirectory = null; break;
                        case "installation-empty": job.InstallationDirectory = " "; break;
                        case "installation-relative": job.InstallationDirectory = "relative"; break;
                        case "path-null": job.InstallerPath = null; break;
                        case "path-empty": job.InstallerPath = " "; break;
                        case "path-other-root": job.InstallerPath = Path.Combine(scope.Root, "elsewhere", "VBAi-Setup-win-x64.exe"); break;
                    }
                    Assert.ThrowsException<InvalidDataException>(() => job.Validate(scope.Root), kind);
                }
                File.WriteAllText(Path.Combine(scope.Root, "pending.json"), "null"); Assert.ThrowsException<InvalidDataException>(() => UpdateInstallJob.Load(scope.Root));
                var valid = UpdateInstallerRunnerTests.Job(scope); valid.InstallerPath = UpdatePaths.AssetPath(scope.Root, valid.Sha256, "VBAi-Setup-win-x64.msi"); valid.Save(scope.Root);
                Assert.AreEqual(valid.InstallerPath, UpdateInstallJob.Load(scope.Root).InstallerPath);
            }
        }
    }
}
