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
    }
}
