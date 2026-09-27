using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeProjectComponentsTests
    {
        private static VbeProjectComponents Service(FakeProject project)
        {
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new VbeProjectComponents(vbe, null);
        }

        [TestMethod]
        public void ProjectVersionChangesWhenComponentInventoryChanges()
        {
            var project = new FakeProject();
            var service = Service(project);
            dynamic before = service.ProjectProperties(project.Name);
            Assert.AreEqual(0, ((IEnumerable)before.Components).Cast<object>().Count());
            project.VBComponents.Add(new FakeComponent("Module1", 1));
            dynamic after = service.ProjectProperties(project.Name);
            Assert.AreNotEqual((string)before.Version, (string)after.Version);
            Assert.AreEqual(1, ((IEnumerable)after.Components).Cast<object>().Count());
        }

        [TestMethod]
        public void ProjectRenameIsRefusedBeforeAnyProjectLookup()
        {
            var service = new VbeProjectComponents(new FakeVbe(), null);
            var request = new Request { Project = "missing", Property = "Name", Value = "Renamed" };
            var error = Assert.ThrowsException<InvalidOperationException>(() => service.SetProjectProperty(request));
            StringAssert.Contains(error.Message, "rename is disabled");
        }

        [TestMethod]
        public void ProjectPropertyRequiresFreshVersionAndPersistsReadback()
        {
            var project = new FakeProject();
            var service = Service(project);
            var request = new Request { Project = project.Name, Property = "Description", Value = "New description",
                ExpectedProjectVersion = "stale" };
            Assert.ThrowsException<InvalidOperationException>(() => service.SetProjectProperty(request));
            Assert.AreEqual("Original", project.Description);
            dynamic before = service.ProjectProperties(project.Name);
            request.ExpectedProjectVersion = before.Version;
            dynamic after = service.SetProjectProperty(request);
            Assert.AreEqual("New description", project.Description);
            Assert.AreNotEqual((string)before.Version, (string)after.Version);
        }

        [TestMethod]
        public void ComponentPropertyRejectsStaleVersionBeforeMutation()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1);
            project.VBComponents.Add(module);
            var service = Service(project);
            var request = new Request { Project = project.Name, Module = module.Name,
                Property = "Description", Value = "Changed", ExpectedComponentVersion = "stale" };
            Assert.ThrowsException<InvalidOperationException>(() => service.SetComponentProperty(request));
            Assert.AreEqual("Original", module.Description);
            dynamic before = service.ComponentProperties(project.Name, module.Name);
            request.ExpectedComponentVersion = before.Version;
            dynamic after = service.SetComponentProperty(request);
            Assert.AreEqual("Changed", module.Description);
            Assert.AreNotEqual((string)before.Version, (string)after.Version);
        }

        [TestMethod]
        public void RenameRejectsInvalidAndDuplicateNamesBeforeChangingComponent()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1);
            project.VBComponents.Add(module);
            project.VBComponents.Add(new FakeComponent("Existing", 1));
            var service = Service(project);
            var request = new Request { Project = project.Name, Module = module.Name, NewName = "../escape" };
            Assert.ThrowsException<ArgumentException>(() => service.RenameComponent(request));
            request.NewName = "existing";
            Assert.ThrowsException<InvalidOperationException>(() => service.RenameComponent(request));
            Assert.AreEqual("Module1", module.Name);
            request.NewName = "Renamed";
            dynamic state = service.ComponentProperties(project.Name, module.Name);
            request.ExpectedComponentVersion = state.Version;
            dynamic renamed = service.RenameComponent(request);
            Assert.AreEqual("Renamed", module.Name);
            Assert.AreEqual("Renamed", (string)renamed.Component);
        }

        [TestMethod]
        public void HostDocumentModuleCannotBeRemoved()
        {
            var project = new FakeProject();
            var document = new FakeComponent("ThisWorkbook", 100);
            project.VBComponents.Add(document);
            var service = Service(project);
            dynamic projectState = service.ProjectProperties(project.Name);
            dynamic componentState = service.ComponentProperties(project.Name, document.Name);
            var request = new Request { Project = project.Name, Module = document.Name,
                ExpectedProjectVersion = projectState.Version,
                ExpectedComponentVersion = componentState.Version };
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveComponent(request));
            Assert.AreEqual(0, project.VBComponents.RemoveAttempts);
            Assert.AreSame(document, project.VBComponents.Single());
        }

        [TestMethod]
        public void ImportErrorAfterAppliedChangeNeverTriggersAutomaticRetry()
        {
            var project = new FakeProject();
            project.VBComponents.ImportThenThrow = true;
            var service = Service(project);
            dynamic state = service.ProjectProperties(project.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-" + Guid.NewGuid().ToString("N") + ".bas");
            File.WriteAllText(path, "Attribute VB_Name = \"ImportedModule\"");
            try
            {
                dynamic result = service.ImportComponent(new Request {
                    Project = project.Name, Path = path, ExpectedProjectVersion = state.Version
                });
                Assert.IsTrue((bool)result.Applied);
                Assert.IsTrue((bool)result.Verified);
                Assert.AreEqual("ImportedModule", (string)result.ImportedName);
                StringAssert.Contains((string)result.ImportError, "after add");
                Assert.AreEqual(1, project.VBComponents.ImportAttempts);
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void ExportRequiresFileReadbackAfterVbeReportsSuccess()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1);
            project.VBComponents.Add(module);
            var service = Service(project);
            dynamic state = service.ComponentProperties(project.Name, module.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-export-" + Guid.NewGuid().ToString("N") + ".bas");
            try
            {
                var error = Assert.ThrowsException<IOException>(() => service.ExportComponent(new Request {
                    Project = project.Name, Module = module.Name, Path = path,
                    ExpectedComponentVersion = state.Version
                }));
                StringAssert.Contains(error.Message, "did not create");
                Assert.AreEqual(1, module.ExportAttempts);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        [TestMethod]
        public void ClassInstancingRejectsFractionAndNonClassThenVerifiesValue()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Class1", 2);
            project.VBComponents.Add(module);
            var service = Service(project);
            var request = new Request { Project = project.Name, Module = module.Name, Value = 1.5 };
            Assert.ThrowsException<ArgumentException>(() => service.SetClassInstancing(request));
            Assert.AreEqual(1, module.Properties.Item("Instancing").Value);
            module.Type = 1;
            request.Value = 2;
            Assert.ThrowsException<InvalidOperationException>(() => service.SetClassInstancing(request));
            module.Type = 2;
            dynamic state = service.ComponentProperties(project.Name, module.Name);
            request.ExpectedComponentVersion = state.Version;
            dynamic result = service.SetClassInstancing(request);
            Assert.AreEqual(2, module.Properties.Item("Instancing").Value);
            Assert.AreEqual("PublicNotCreatable", (string)result.Meaning);
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; } = "VBAProject";
            public string Description { get; set; } = "Original";
            public string FileName { get; set; } = @"C:\fixture\Book.xlsm";
            public int Mode { get; set; } = 2;
            public bool Saved { get; set; } = true;
            public FakeComponentCollection VBComponents { get; } = new FakeComponentCollection();
            public List<object> References { get; } = new List<object>();
        }

        public sealed class FakeComponentCollection : IEnumerable<FakeComponent>
        {
            private readonly List<FakeComponent> items = new List<FakeComponent>();
            public int RemoveAttempts { get; private set; }
            public int ImportAttempts { get; private set; }
            public bool ImportThenThrow { get; set; }
            public void Add(FakeComponent component) { items.Add(component); }
            public void Remove(FakeComponent component) { RemoveAttempts++; items.Remove(component); }
            public void Import(string path)
            {
                ImportAttempts++;
                items.Add(new FakeComponent("ImportedModule", 1));
                if (ImportThenThrow) throw new InvalidOperationException("COM error after add");
            }
            public IEnumerator<FakeComponent> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeComponent
        {
            public FakeComponent(string name, int type) { Name = name; Type = type; }
            public string Name { get; set; }
            public string Description { get; set; } = "Original";
            public int Type { get; set; }
            public FakePropertyCollection Properties { get; } = new FakePropertyCollection();
            public FakeCodeModule CodeModule { get; } = new FakeCodeModule();
            public int ExportAttempts { get; private set; }
            public void Export(string path) { ExportAttempts++; }
        }

        public sealed class FakePropertyCollection : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> items = new List<FakeProperty> {
                new FakeProperty { Name = "Instancing", Value = 1 }
            };
            public FakeProperty Item(string name) { return items.Single(p => p.Name == name); }
            public IEnumerator<FakeProperty> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeProperty
        {
            public string Name { get; set; }
            public object Value { get; set; }
        }

        public sealed class FakeCodeModule
        {
            public int CountOfLines { get { return 1; } }
            public string Lines(int start, int count) { return "Option Explicit"; }
        }
    }
}
