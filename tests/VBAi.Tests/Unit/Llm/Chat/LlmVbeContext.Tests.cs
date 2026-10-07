namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmVbeContextTests
    {
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
