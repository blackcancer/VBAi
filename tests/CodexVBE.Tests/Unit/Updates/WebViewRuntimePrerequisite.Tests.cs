using System;
using System.IO;
using System.Threading.Tasks;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Updates
{
    [TestClass, TestCategory("Unit")]
    public sealed class WebViewRuntimePrerequisiteTests
    {
        [TestMethod]
        public async Task InstalledRuntimeSkipsDownloadAndInstallation()
        {
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => true, Download = p => throw new AssertFailedException("Unexpected download") };
            await runtime.Ensure(Path.GetTempPath());
            Assert.IsTrue(WebViewRuntimePrerequisite.ValidVersion("120.0.1.2"));
            Assert.IsFalse(WebViewRuntimePrerequisite.ValidVersion("0.0.0.0"));
            Assert.IsFalse(WebViewRuntimePrerequisite.ValidVersion(null));
        }
        [TestMethod]
        public async Task UnsignedInstallerIsNeverExecuted()
        {
            string downloaded = null;
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => false, Download = p => { downloaded = p; File.WriteAllText(p, "fixture"); return Task.CompletedTask; }, Verify = p => false, Install = p => throw new AssertFailedException("Untrusted execution") };
            await Assert.ThrowsExceptionAsync<InvalidDataException>(() => runtime.Ensure(Path.GetTempPath()));
            Assert.IsFalse(File.Exists(downloaded));
        }
        [TestMethod]
        public async Task ExitSuccessRequiresInstalledRuntimeReadback()
        {
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => false, Download = p => { File.WriteAllText(p, "fixture"); return Task.CompletedTask; }, Verify = p => true, Install = p => 0 };
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => runtime.Ensure(Path.GetTempPath()));
        }
    }
}
