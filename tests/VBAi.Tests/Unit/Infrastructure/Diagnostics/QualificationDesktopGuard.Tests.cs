using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class QualificationDesktopGuardTests
    {
        [TestMethod]
        public void OrdinaryExplicitlyInteractiveInvocationDoesNotRequireCampaignEvidence()
        {
            WithEnvironment(null, "invalid-relative-evidence", () => QualificationDesktopGuard.Initialize(null));
        }

        [TestMethod]
        public void CampaignCannotUseDefaultDesktopOrAnUnrelatedGeneratedDesktop()
        {
            WithEnvironment("Default", "invalid-relative-evidence", () =>
                Assert.ThrowsException<ArgumentException>(() => QualificationDesktopGuard.Initialize(null)));
            // A generated name belonging to no desktop must be refused before any result file or test executes.
            WithEnvironment("VBAiTests_" + Guid.NewGuid().ToString("N"), "invalid-relative-evidence", () =>
                Assert.ThrowsException<InvalidOperationException>(() => QualificationDesktopGuard.Initialize(null)));
        }

        [TestMethod]
        public void CampaignEvidenceMustBeAnExistingAbsoluteDirectory()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-DesktopGuard-" + Guid.NewGuid().ToString("N"));
            foreach (string invalid in new[] { null, "", "relative-evidence", directory })
                Assert.ThrowsException<InvalidOperationException>(() => QualificationDesktopGuard.RequireEvidence(invalid));
            Directory.CreateDirectory(directory);
            try { QualificationDesktopGuard.RequireEvidence(directory); }
            finally { Directory.Delete(directory); }
        }

        private static void WithEnvironment(string desktop, string evidence, Action action)
        {
            string originalDesktop = Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP");
            string originalEvidence = Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP_EVIDENCE");
            try
            {
                Environment.SetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP", desktop);
                Environment.SetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP_EVIDENCE", evidence);
                action();
            }
            finally
            {
                Environment.SetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP", originalDesktop);
                Environment.SetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP_EVIDENCE", originalEvidence);
            }
        }
    }
}
