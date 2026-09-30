using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Office"), DoNotParallelize]
    public sealed class WordGitProjectIsolationTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void TwoSameNameWordDocumentsUseCanonicalGitScopesAndRefuseStaleSaveAs()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_GIT_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_WORD_GIT_TESTS=1 and VBAi_RUN_OFFICE_TESTS=1 for owned two-document Word qualification.");
            string root = null;
            try
            {
                using (var fixture = OfficeVbeFixture.Start("Word"))
                {
                    root = fixture.Root;
                    fixture.QualifyWordGitProjectIsolation();
                }
            }
            finally
            {
                string report = root == null ? null : Path.Combine(root, "qualification.json");
                if (report != null && File.Exists(report)) TestContext.AddResultFile(report);
            }
        }
    }
}
