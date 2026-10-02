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
            Run(fixture => fixture.QualifyWordCanonicalPathSelection());
        }

        [STATestMethod]
        public void DiagnosticOnlyExternalStaGitCaptureAndExportsRemainSeparate()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_GIT_CAPTURE_DIAGNOSTIC") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_WORD_GIT_TESTS") != "1")
                Assert.Inconclusive("The historical external-STA Word Git Capture/export diagnostic requires both explicit Word Git opt-ins.");
            Run(fixture => fixture.DiagnoseWordGitCaptureExports());
        }

        private void Run(Action<OfficeVbeFixture> scenario)
        {
            string root = null;
            try
            {
                using (var fixture = OfficeVbeFixture.Start("Word"))
                {
                    root = fixture.Root;
                    scenario(fixture);
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
