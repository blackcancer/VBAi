using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors project navigation reads, preserving bridge refusals and rejecting incomplete inventories.</summary>
    public sealed partial class VbeChatReferenceTests
    {
        [TestMethod]
        public void ProjectNavigationAlwaysReadsCurrentModulesInsteadOfCachedReferences()
        {
            int reads = 0;
            var references = new VbeChatReferences(null);
            references.Entries.Add(new VbeChatReference { Project = "P", Module = "Removed" });
            references.Execute = request =>
            {
                Assert.AreEqual("list_modules", request.Command); Assert.AreEqual("P", request.Project);
                Assert.IsNull(request.Module); reads++;
                return Response.Success(reads == 1 ? new object[] { new { Name = "ZForm" }, new { Name = "aClass" } } : new object[0]);
            };
            var project = new VbeChatReference { Project = "P" };
            var modules = references.ModulesForProject(project);
            CollectionAssert.AreEqual(new[] { "aClass", "ZForm" }, modules.Select(item => item.Module).ToArray());
            Assert.IsTrue(modules.All(item => item.Project == "P" && item.Name == null && item.Kind == "Module"));
            Assert.AreEqual(0, references.ModulesForProject(project).Length);
            Assert.AreEqual(2, reads); Assert.AreEqual("Removed", references.Entries.Single().Module);
        }

        [DataTestMethod, DataRow("missing"), DataRow("ambiguous"), DataRow("protected")]
        public void ProjectNavigationSurfacesLiveRefusalsAndNeverFallsBackToCachedEntries(string error)
        {
            var references = new VbeChatReferences(null);
            references.Entries.Add(new VbeChatReference { Project = "P", Module = "Old" });
            int calls = 0;
            references.Execute = request => { calls++; return Response.Failure(error); };
            Assert.AreEqual(error, Assert.ThrowsException<InvalidOperationException>(() => references.ModulesForProject(new VbeChatReference { Project = "P" })).Message);
            Assert.AreEqual(1, calls);
        }

        [DataTestMethod, DataRow("null"), DataRow("empty"), DataRow("module"), DataRow("procedure")]
        public void OnlyProjectReferencesAreAdmittedBeforeAnyBridgeRead(string kind)
        {
            var references = new VbeChatReferences(null) { Execute = request => throw new AssertFailedException("Malformed selector reached the bridge.") };
            var project = new VbeChatReference { Project = "P" };
            if (kind == "null") project = null;
            if (kind == "empty") project.Project = " ";
            if (kind == "module") project.Module = "M";
            if (kind == "procedure") project.Name = "Run";
            Assert.ThrowsException<ArgumentException>(() => references.ModulesForProject(project));
        }

        [DataTestMethod]
        [DataRow("null"), DataRow("object"), DataRow("null row"), DataRow("foreign row"), DataRow("missing name"), DataRow("blank name"), DataRow("duplicate")]
        public void IncompleteModuleInventoriesAreRejectedInsteadOfDisplayedAsEmptyOrPartial(string kind)
        {
            object data = null;
            switch (kind)
            {
                case "object": data = new { Name = "M" }; break;
                case "null row": data = new object[] { new { Name = "M" }, null }; break;
                case "foreign row": data = new object[] { new { Name = "M" }, "foreign" }; break;
                case "missing name": data = new object[] { new { Type = 1 } }; break;
                case "blank name": data = new object[] { new { Name = " " } }; break;
                case "duplicate": data = new object[] { new { Name = "M" }, new { Name = "m" } }; break;
            }
            var references = new VbeChatReferences(null) { Execute = request => Response.Success(data) };
            Assert.ThrowsException<ArgumentException>(() => references.ModulesForProject(new VbeChatReference { Project = "P" }));
        }
    }
}
