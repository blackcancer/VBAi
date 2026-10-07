using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormLocalGitTests
    {
        /// <summary>Fresh native import/export qualification; no attached-font diagnostic getter intervenes.</summary>
        [STATestMethod]
        public void OwnedNestedFontCloneImportRecoveryPreservesExactNativeState()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_NESTED_FONT_CLONE_QUALIFICATION") != "1")
                Assert.Inconclusive("Set VBAi_RUN_NESTED_FONT_CLONE_QUALIFICATION=1 for the fresh guarded native clone qualification.");
            Assert.IsTrue(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(FormFontObservation.ManifestVariable)),
                "Root font transfer diagnostics must be disabled for ordinary clone qualification.");
            Assert.IsTrue(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(FormFontObservation.AttachedManifestVariable)),
                "Attached metric observations must be disabled before the first authoritative native export.");
            // RunLayout creates a fresh saved/reopened synthetic workbook, declares
            // each step once, retains strict full-resource equality and stops after
            // any refused or uncertain mutation. It exports a refused state before
            // optional font getters. Root declaration loss remains a full failure.
            RunLayout("FrameMultiPage", ExcelVbeFixture.Run, "NestedFontCloneNativeOwnerQualification");
        }
    }
}
