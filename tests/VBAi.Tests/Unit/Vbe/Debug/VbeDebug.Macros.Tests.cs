using System;
using System.Linq;
using System.Web.Script.Serialization;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class MacroCatalogueTests
    {
        /// <summary>Le catalogue distingue macros, fonctions, paramètres, privés et compilation conditionnelle.</summary>
        [TestMethod]
        public void MacroCatalogueDoesNotMistakePublicFunctionsOrConditionalAndPrivateCodeForChooserMacros()
        {
            var vbe = new IdeSurfaceFixture.Vbe(); var project = new IdeSurfaceFixture.Project(); vbe.VBProjects.Add(project);
            var module = new IdeSurfaceFixture.Component(); project.VBComponents.Add(module);
            module.CodeModule.Source = "Option Explicit\r\nPublic Sub Entry()\r\nEnd Sub\r\nPrivate Sub Hidden()\r\nEnd Sub\r\nPublic Function Value$(ByVal n As Long)\r\nEnd Function\r\n#If VBA7 Then\r\nPublic Sub ConditionalEntry()\r\nEnd Sub\r\n#End If\r\nSub ImplicitPublic()\r\nEnd Sub";
            project.VBComponents.Add(new IdeSurfaceFixture.Component { Name = "Class1", Type = 2 });
            var service = new VbeDebug(vbe); dynamic response = service.ListMacros(new Request { Project = "P" });
            var json = new JavaScriptSerializer(); var rows = ((object[])response.Macros).Select(row => json.Deserialize<dynamic>(json.Serialize(row))).ToArray();
            Assert.AreEqual(4, (int)response.Total); Assert.AreEqual(2, rows.Count(r => (bool)r["NativeMacroCandidate"]));
            Assert.IsFalse(rows.Any(r => r["Procedure"] == "Hidden"));
            var function = rows.Single(r => r["Procedure"] == "Value"); Assert.IsFalse((bool)function["NativeMacroCandidate"]);
            Assert.AreEqual(1, ((object[])function["Parameters"]).Length);
            var conditional = rows.Single(r => r["Procedure"] == "ConditionalEntry"); Assert.IsTrue((bool)conditional["Conditional"]); Assert.IsNull(conditional["ExecutionTool"]);
            dynamic filtered = service.ListMacros(new Request { Project = "P", Query = "entry", Limit = 1 });
            Assert.AreEqual(2, (int)filtered.Total); Assert.IsTrue((bool)filtered.HasMore);
            Assert.AreEqual((string)response.CatalogVersion, (string)filtered.CatalogVersion);
            module.CodeModule.Source = "Option Private Module\r\nSub Entry()\r\nEnd Sub";
            dynamic changed = service.ListMacros(new Request { Project = "P" }); Assert.AreNotEqual((string)response.CatalogVersion, (string)changed.CatalogVersion);
            Assert.IsFalse((bool)json.Deserialize<dynamic>(json.Serialize(((object[])changed.Macros)[0]))["NativeMacroCandidate"]);
        }

        /// <summary>Les limites de requête sont contrôlées sans interroger la source.</summary>
        [TestMethod]
        public void MacroCatalogueRejectsInvalidPaginationAndIdentifiers()
        {
            var service = new VbeDebug(new IdeSurfaceFixture.Vbe());
            foreach (int i in Enumerable.Range(0, 8))
            {
                var r = new Request { Project = "P" };
                if (i == 0) r = null; if (i == 1) r.Project = " "; if (i == 2) r.Offset = -1; if (i == 3) r.Offset = 100001;
                if (i == 4) r.Limit = -1; if (i == 5) r.Limit = 501; if (i == 6) r.Query = new string('x', 257); if (i == 7) r.Project = null;
                Assert.ThrowsException<ArgumentException>(() => service.ListMacros(r));
            }
        }

        [TestMethod]
        public void MacroCatalogueBoundsSourceAndEntriesAndPreservesIncompleteDirectiveSemantics()
        {
            var vbe = new IdeSurfaceFixture.Vbe(); var project = new IdeSurfaceFixture.Project(); vbe.VBProjects.Add(project);
            var module = new IdeSurfaceFixture.Component(); project.VBComponents.Add(module); var service = new VbeDebug(vbe);
            dynamic empty = service.ListMacros(new Request { Project = "P" }); Assert.AreEqual(0, (int)empty.Total);
            module.CodeModule.Source = new string('\n', 200000);
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.ListMacros(new Request { Project = "P" })).Message, "source limit");
            module.CodeModule.Source = "#\r\n#End\r\nPublic\r\nSub Visible()\r\nEnd Sub";
            dynamic incomplete = service.ListMacros(new Request { Project = "P" }); Assert.AreEqual(1, (int)incomplete.Total);
            dynamic row = new JavaScriptSerializer().Deserialize<dynamic>(new JavaScriptSerializer().Serialize(((object[])incomplete.Macros)[0]));
            Assert.AreEqual("Visible", (string)row["Procedure"]); Assert.IsTrue((bool)row["NativeMacroCandidate"]);
            module.CodeModule.Source = string.Join("\r\n", Enumerable.Range(0, 10001).Select(i => "Sub Entry" + i + "()\r\nEnd Sub"));
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.ListMacros(new Request { Project = "P" })).Message, "10000 procedures");
        }
    }
}
