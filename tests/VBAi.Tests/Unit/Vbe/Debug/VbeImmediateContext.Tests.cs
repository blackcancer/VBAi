namespace VBAi.Tests.Unit
{
    using System;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeImmediateContextTests
    {
        [TestMethod]
        public void CanonicalWordSelectionRefusesBackingPathOtherDocumentAndMissingIdentityPath()
        {
            var state = new { Project = "Project", SelectedProject = "Project", ActiveModule = "Owned",
                SelectedProjectPath = @"C:\Temp\~WRL0001.tmp", SelectedHostPath = @"C:\Owned\First.docm" };
            VbeImmediateContext.RequireProject(@"C:\Owned\First.docm", state);
            Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject(@"C:\Owned\Second.docm", state));
            Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject(@"C:\Temp\~WRL0001.tmp", state));
            var unavailable = new { Project = "Project", SelectedProject = "Project", ActiveModule = "Owned",
                SelectedProjectPath = @"C:\Temp\~WRL0001.tmp", SelectedHostPath = (string)null };
            Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject(@"C:\Temp\~WRL0001.tmp", unavailable));
        }

        [TestMethod]
        public void NativeSelectionRequiresIdentityAndFailsClosedOnMissingContext()
        {
            var state = new { Project = "P", SelectedProject = "P", SelectedProjectPath = @"C:\Temp\P.xlsm", ActiveModule = "M" };
            VbeImmediateContext.RequireProject("p", state);
            VbeImmediateContext.RequireProject(@"c:\temp\p.xlsm", state);
            Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject("Other", state));
            Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject(@"C:\Other\P.xlsm", state));
            foreach (object invalid in new object[] { null, new { }, new { Project = "P", SelectedProject = "P" },
                new { Project = "P", SelectedProject = "Other", ActiveModule = "M" },
                new { Project = "P", SelectedProject = "P", ActiveModule = "" },
                new { Project = "P", SelectedProject = (string)null, ActiveModule = "M" } })
                Assert.ThrowsException<InvalidOperationException>(() => VbeImmediateContext.RequireProject("P", invalid));
        }
    }
}
