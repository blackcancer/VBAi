using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class GitCoverageTests
    {
        [TestMethod]
        [STATestMethod]
        public void GitWorkflowUsesTheProductionAssembly()
        {
            GitTests.RunSuite();
        }
    }
}
