using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsContractTests
    {
        [TestMethod]
        public void FormListingIncludesOnlyUserFormsAndTheirDesignerState()
        {
            var project = new FakeProject { Name = "Projet" };
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1 });
            project.VBComponents.Add(new FakeComponent { Name = "Accueil", Type = 3, HasOpenDesigner = true });
            project.VBComponents.Add(new FakeComponent { Name = "Détails", Type = 3, HasOpenDesigner = false });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var forms = new VbeForms(host);
            dynamic list = forms.List("projet");
            Assert.AreEqual(2, ((IEnumerable<object>)list).Count());
            var rows = ((IEnumerable<object>)list).Cast<dynamic>().ToArray();
            Assert.AreEqual("Accueil", (string)rows[0].Name);
            Assert.IsTrue((bool)rows[0].DesignerOpen);
            Assert.AreEqual("Détails", (string)rows[1].Name);
            Assert.IsFalse((bool)rows[1].DesignerOpen);
            Assert.ThrowsException<ArgumentException>(() => forms.SetProperty(new Request { Property = "", Value = "value" }));
        }

        [TestMethod]
        public void NativeControlsAndMissingComMetadataAreReportedHonestly()
        {
            var forms = new VbeForms(new FakeVbe());
            var types = ((IEnumerable<object>)forms.ControlTypes()).Cast<dynamic>().ToArray();
            Assert.IsTrue(types.Any(item => (string)item.ProgId == "Forms.Label.1" &&
                (string)item.Source == "MSForms native"));
            Assert.IsTrue(types.Any(item => (string)item.ProgId == "Forms.TextBox.1"));
            Assert.IsTrue(VbeControlCatalog.IsCandidate("Forms.Label.1",
                new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Forms.Label.1" }));
            dynamic events = VbeComEvents.Read(new object());
            Assert.IsFalse((bool)events.SourceInterfacesComplete);
            Assert.IsFalse((bool)events.VbeEventCatalogComplete);
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(null, "Caption"));
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(new object(), " "));
            dynamic metadata = VbeComPropertyAccessors.Inspect(new object(), "Caption");
            Assert.IsFalse((bool)metadata.MetadataComplete);
            Assert.IsNull((object)metadata.SetterDeclared);
            Assert.IsFalse((bool)metadata.ActualSetterVerified);
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }
        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public bool HasOpenDesigner { get; set; }
        }
    }
}
