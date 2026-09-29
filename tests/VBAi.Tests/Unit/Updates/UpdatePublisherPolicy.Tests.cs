using System;
using System.IO;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdatePublisherPolicyTests
    {
        [TestMethod]
        public void NoReleasePublisherIsPinnedSoNeitherUnsignedNorWindowsTrustedFilesAreAccepted()
        {
            using (var scope = new UpdateScope())
            {
                var job = UpdateInstallerRunnerTests.Job(scope);
                Assert.IsFalse(UpdateInstallerRunner.VerifySignatureNative(job.InstallerPath));
                Assert.IsFalse(UpdatePublisherPolicy.Accepts(job.InstallerPath));
            }

            string signedWindowsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "dotnet.exe");
            Assert.IsTrue(File.Exists(signedWindowsFile));
            Assert.IsTrue(UpdateInstallerRunner.VerifySignatureNative(signedWindowsFile));
            Assert.IsFalse(UpdatePublisherPolicy.Accepts(signedWindowsFile));
        }
    }
}
