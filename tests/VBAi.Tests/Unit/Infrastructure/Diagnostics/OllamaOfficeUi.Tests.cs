using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Native guards reject foreign, hidden and replaced assistant ancestry without re-reading stale UIA providers.</summary>
    [TestClass]
    public sealed class OllamaOfficeUiTests
    {
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
