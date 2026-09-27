using System;
using System.Linq;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeChatReferenceTests
    {
        private static VbeSession Session()
        {
            var vbe = new VbeSessionTests.FakeVbe();
            var project = new VbeSessionTests.FakeProject {
                Name = "VBAProject", FileName = @"C:\Temp\ReferenceTest.xlsm", Mode = 2 };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent {
                Name = "Module1", Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("Sub Example()\r\nEnd Sub") });
            vbe.VBProjects.Add(project);
            return new VbeSession(vbe);
        }

        [TestMethod]
        public void LiveContextNamesHostAndRealAnsiCodePageWithoutAssumingExcel()
        {
            dynamic snapshot = LlmVbeContext.LiveSnapshot(Session());
            Assert.IsTrue((bool)snapshot.VbeConnected);
            Assert.AreEqual(Encoding.Default.CodePage, (int)snapshot.HostAnsiCodePage);
            Assert.IsNotNull(snapshot.Projects);
            StringAssert.Contains(LlmVbeContext.DeveloperInstructions, "never assume Excel");
            StringAssert.Contains(LlmVbeContext.EncodingInstructions,
                Encoding.Default.CodePage.ToString());
        }

        [TestMethod]
        public void ReferenceDiscoveryLoadsProjectsThenModulesAndReportsProcedureProbeError()
        {
            var references = new VbeChatReferences(Session());
            int changed = 0;
            references.Changed += () => changed++;
            references.Refresh();
            Assert.AreEqual(1, references.Entries.Count);
            Assert.AreEqual("#VBAProject", references.Entries[0].Token);
            Assert.IsTrue(references.IsLoading);
            references.Step();
            Assert.AreEqual(2, references.Entries.Count);
            Assert.AreEqual("#VBAProject.Module1", references.Entries[1].Token);
            Assert.IsTrue(references.IsLoading);
            references.Step();
            Assert.IsFalse(references.IsLoading);
            Assert.IsFalse(string.IsNullOrWhiteSpace(references.Error));
            Assert.AreEqual(3, changed);
        }

        [TestMethod]
        public void TokensFilterProjectModuleAndPropertyProcedurePrefixes()
        {
            var references = new VbeChatReferences(Session());
            references.Refresh();
            references.Step();
            references.Entries.Add(new VbeChatReference { Project = "VBAProject", Module = "Module1",
                Name = "Title", Kind = "Property", ProcKind = 3 });
            Assert.AreEqual("@VBAProject.Module1.Title:Get", references.Entries.Last().Token);
            Assert.AreEqual(2, references.MatchPrefix("#VBAProject", '#').Count());
            Assert.AreEqual(1, references.MatchPrefix("@Title", '@').Count());
            Assert.AreEqual(0, references.MatchPrefix("@Title", '#').Count());
            Assert.AreEqual(1, references.Match("Title").Count());
        }

        [TestMethod]
        public void ResolveUsesFreshModuleSourceAndRefusesStaleProcedureSelection()
        {
            var references = new VbeChatReferences(Session());
            var project = new VbeChatReference { Project = "VBAProject", Kind = "Projet" };
            StringAssert.Contains(references.Resolve(project), "Module1");
            var module = new VbeChatReference { Project = "VBAProject", Module = "Module1", Kind = "Module" };
            string resolved = references.Resolve(module);
            StringAssert.Contains(resolved, "Sub Example()");
            Assert.IsFalse(string.IsNullOrWhiteSpace(module.Sha256));
            var procedure = new VbeChatReference { Project = "VBAProject", Module = "Module1",
                Name = "Example", Kind = "Sub", StartLine = 1, EndLine = 2, Sha256 = "stale" };
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(procedure));
            procedure.Sha256 = module.Sha256;
            StringAssert.Contains(references.Resolve(procedure), "End Sub");
            procedure.StartLine = 99;
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(procedure));
        }

        [TestMethod]
        public void ProjectReferenceCannotNavigateDirectlyToCode()
        {
            var references = new VbeChatReferences(Session());
            var result = references.Navigate(new VbeChatReference { Project = "VBAProject", Kind = "Projet" });
            Assert.IsFalse(result.Ok);
            StringAssert.Contains(result.Error, "module");
        }
    }
}
