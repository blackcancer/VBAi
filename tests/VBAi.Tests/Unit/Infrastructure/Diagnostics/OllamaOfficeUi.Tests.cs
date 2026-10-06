using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Native guards reject foreign, hidden and replaced assistant ancestry without re-reading stale UIA providers.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeUiTests
    {
        [DataTestMethod]
        [DataRow(0, 11, false)]
        [DataRow(11, 12, false)]
        [DataRow(11, 11, true)]
        public void PendingOrForeignObserverRetainsItsReferences(int original, int current, bool pending)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.RequireReferenceRelease(original, current, pending));

        [TestMethod]
        public void OriginalSettledObserverCanDropItsManagedReferences()
            => OllamaOfficeUi.RequireReferenceRelease(11, 11, false);

        [DataTestMethod]
        [DataRow("pid")]
        [DataRow("thread")]
        [DataRow("sitePid")]
        [DataRow("siteThread")]
        [DataRow("containerClass")]
        [DataRow("siteClass")]
        [DataRow("caption")]
        [DataRow("containerHidden")]
        [DataRow("siteHidden")]
        [DataRow("parent")]
        [DataRow("owner")]
        public void OriginalNativeHostAncestryRejectsEveryChangedIdentity(string changed)
        {
            Assert.IsFalse(OllamaOfficeUi.HostedNodesMatch(42, 7,
                changed == "pid" ? 43 : 42, changed == "thread" ? 8u : 7u,
                changed == "sitePid" ? 43 : 42, changed == "siteThread" ? 8u : 7u,
                changed == "containerClass" ? "Other" : "WindowsForms10.Window",
                changed == "siteClass" ? "GenericPaneOther" : "GenericPane",
                changed == "caption" ? "Other" : "VBAi",
                changed != "containerHidden", changed != "siteHidden", changed != "parent", changed != "owner"));
        }

        [TestMethod]
        public void ExactOriginalNativeAssistantAncestryIsAdmitted()
            => Assert.IsTrue(OllamaOfficeUi.HostedNodesMatch(42, 7, 42, 7, 42, 7,
                "WindowsForms10.Window", "GenericPane", "VBAi", true, true, true, true));
    }
}
