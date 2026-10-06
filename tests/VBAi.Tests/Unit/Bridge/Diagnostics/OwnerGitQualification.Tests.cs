using System;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OwnerGitQualificationTests
    {
        [TestMethod]
        public void ExpectedCorruptionRefusalNeedsMeasuredUnchangedProjectAndRecoveryRefs()
        {
            Func<bool, bool, bool, string, string, bool, bool> prove = (started, pending, sameProject, before, after, sameAfter) =>
                OwnerGitQualification.IsProvenPrewriteRefusal("Form resources are missing", "Form resources are missing: Form1",
                    started, false, pending, before, after, "after-0", sameAfter ? "after-0" : "after-1", sameProject);
            Assert.IsTrue(prove(false, false, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(true, false, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, true, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, false, false, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, false, true, "backup-0", "backup-1", true));
            Assert.IsFalse(prove(false, false, true, "backup-0", "backup-0", false));
            Assert.IsFalse(OwnerGitQualification.IsProvenPrewriteRefusal("missing", "other error", false,
                false, false, null, null, null, null, true));
        }

        [TestMethod]
        public async Task DisabledAtConnectionCannotBeEnabledByChangingEnvironmentLater()
        {
            string previous = Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, null);
                var disabled = new OwnerGitQualification(null, 42);
                Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName,
                    @"C:\Evidence\0123456789abcdef0123456789abcdef.owner-git.json");
                var request = new Request { Command = OwnerGitQualificationManifest.CommandName,
                    Action = "0123456789abcdef0123456789abcdef", ExpectedSha256 = new string('a', 64) };
                string json = "{\"Command\":\"diagnostic_userform_git\",\"Action\":\"" + request.Action +
                    "\",\"ExpectedSha256\":\"" + request.ExpectedSha256 + "\"}";
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => disabled.ExecuteAsync(request, json));
                StringAssert.Contains(error.Message, "disabled");
            }
            finally { Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, previous); }
        }
    }
}
