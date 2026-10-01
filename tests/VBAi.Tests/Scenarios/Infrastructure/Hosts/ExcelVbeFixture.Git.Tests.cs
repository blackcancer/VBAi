using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Scenarios
{
    /// <summary>Checks real nested fixture COM lifetimes without changing project content or executing macros.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("ExcelGitFixtureLifetime"), DoNotParallelize]
    public sealed class ExcelGitFixtureLifetimeTests
    {
        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void OuterGitProjectSurvivesNativeReadbackAndNestedScopeCleanup(bool nestedFailure)
        {
            string root = Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS");
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Owned native Excel qualification requires VBAi_RUN_EXCEL_TESTS=1.");
            Assert.IsTrue(!string.IsNullOrEmpty(root) && Path.IsPathRooted(root));
            var host = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(root, "fixture-lifetime-" + Guid.NewGuid().ToString("N") + ".jsonl"));
            try
            {
                const string form = "LifetimeForm";
                string path = host.File("Lifetime.xlsm");
                host.PrepareGitForm(form, "Synthetic lifetime form", "NeverExecuted", path);
                host.WithGitProject(path, outer => {
                    VbaGitSnapshot before = outer.Capture();
                    Action nested = () => host.WithGitProject(path, inner => {
                        var state = host.ReadGitForm(form);
                        Assert.AreEqual("Synthetic lifetime form", state["Caption"]);
                        Assert.IsTrue(before.SameAs(inner.Capture()));
                        if (nestedFailure) throw new InvalidOperationException("Synthetic nested callback failure");
                    });
                    if (nestedFailure)
                    {
                        var error = Assert.ThrowsException<InvalidOperationException>(nested);
                        Assert.AreEqual("Synthetic nested callback failure", error.Message);
                    }
                    else nested();
                    Assert.IsTrue(before.SameAs(outer.Capture()), "Nested readback and cleanup must preserve the outer project and exact logical snapshot.");
                });
            }
            finally { host.Dispose(); }
        }
    }
}
