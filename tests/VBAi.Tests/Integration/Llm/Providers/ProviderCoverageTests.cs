using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Intègre le fixture protocolaire fournisseur à la suite de tests locale.</summary>
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class ProviderCoverageTests
    {
        /// <summary>Exécute les scénarios de protocole du binaire fixture et réinitialise son chemin après usage.</summary>
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
