namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmVbeContextTests
    {
        /// <summary>Preserves the selected host-document or unsaved-project selector without changing the user's request.</summary>
        [DataTestMethod]
        [DataRow("Macro document", @"C:\VBAiSynthetic\Disposable.docm")]
        [DataRow("Macro publication", @"C:\VBAiSynthetic\Disposable.pub")]
        [DataRow("Classeur français", @"E:\Développement\Essai.xlsm")]
        [DataRow("Unsaved", "VBAProject")]
        [DataRow(null, @"C:\VBAiSynthetic\Disposable.accdb")]
        [DataRow("Ignored when unbound", null)]
        public void ProjectContextPreservesExactSelectorAndUserRequest(string label, string project)
        {
            const string request = "Read ObservedMarker from Q028Marker.\nDo not edit or execute code.";
            string actual = LlmVbeContext.WithProjectScope(request, label, project);
            if (project == null)
                Assert.AreEqual(request, actual, "An unbound request must not inherit a stale project label.");
            else
            {
                StringAssert.StartsWith(actual, UiText.Get("VBA project for this conversation: ") + label);
                StringAssert.Contains(actual, UiText.Get("\nProject identifier to use in tools: ") + project + "\n\n");
                StringAssert.EndsWith(actual, request);
            }
        }

        /// <summary>Refuses a missing request rather than manufacturing a context-only conversation turn.</summary>
        [TestMethod]
        public void ProjectContextRejectsMissingRequest()
        {
            Assert.ThrowsException<ArgumentNullException>(() => LlmVbeContext.WithProjectScope(null, "Project", "VBAProject"));
        }

        [TestMethod]
        public void LiveEnvironmentSnapshotDistinguishesAccessibleProjectsAndDiscoveryFailure()
        {
            StringAssert.Contains(LlmVbeContext.DeveloperInstructions, "Visual Basic Editor");
            StringAssert.Contains(LlmVbeContext.EncodingInstructions, "inspect_code_file");
            var available = LlmVbeContext.LiveSnapshot(new VbeSession(new VbeSessionCoverageTests.SessionHost()));
            var unavailable = LlmVbeContext.LiveSnapshot(new VbeSession(new object()));
            Func<object, string, object> field = (value, name) => value.GetType().GetProperty(name).GetValue(value);
            Assert.IsNotNull(field(available, "Projects")); Assert.IsNull(field(available, "ProjectsError"));
            Assert.IsNull(field(unavailable, "Projects")); Assert.IsFalse(string.IsNullOrWhiteSpace((string)field(unavailable, "ProjectsError")));
            Assert.IsTrue((int)field(available, "HostProcessId") > 0); Assert.AreEqual(true, field(available, "VbeConnected"));
        }
    }
}
