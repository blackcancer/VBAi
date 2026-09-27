using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class ProviderCoverageTests
    {
        [TestMethod]
        [STATestMethod]
        public void ProviderProtocolsUseTheProductionAssembly()
        {
            var fixture = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProviderTests.exe");
            Assert.IsTrue(File.Exists(fixture), "The Copilot protocol fixture was not copied to the VSTest output.");
            ProviderTests.CopilotFixtureExecutable = fixture;
            try { ProviderTests.RunCoverageSuite(); }
            finally { ProviderTests.CopilotFixtureExecutable = null; }
        }
    }
}
