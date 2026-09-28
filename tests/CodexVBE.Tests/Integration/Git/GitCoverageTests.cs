using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Intègre la suite smoke Git au runner de tests avec l’assembly de production.</summary>
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class GitCoverageTests
    {
        /// <summary>Exécute le workflow Git et les contrôles Designer dans un projet de test.</summary>
        [TestMethod]
        [STATestMethod]
        public void GitWorkflowUsesTheProductionAssembly()
        {
            GitTests.RunSuite();
        }
    }
}
