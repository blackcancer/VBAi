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
            return new VbeProjectComponents(vbe, new VbeForms(vbe));
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
        public void RemoveComponentChecksBothVersionsAndReturnsUpdatedInventory()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1);
            project.VBComponents.Add(module);
            var service = Service(project);
            dynamic projectState = service.ProjectProperties(project.Name);
            dynamic componentState = service.ComponentProperties(project.Name, module.Name);
            var request = new Request { Project = project.Name, Module = module.Name,
                ExpectedProjectVersion = "stale", ExpectedComponentVersion = componentState.Version };
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveComponent(request));
            Assert.AreEqual(0, project.VBComponents.RemoveAttempts);
            request.ExpectedProjectVersion = projectState.Version;
            request.ExpectedComponentVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveComponent(request));
            Assert.AreEqual(0, project.VBComponents.RemoveAttempts);
            request.ExpectedComponentVersion = componentState.Version;
            dynamic result = service.RemoveComponent(request);
            Assert.AreEqual(1, project.VBComponents.RemoveAttempts);
            Assert.AreEqual(0, project.VBComponents.Count());
            Assert.AreNotEqual((string)projectState.Version, (string)result.Version);
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
        public void ImportWithoutExactlyOneNewComponentReportsUncertainOutcome()
        {
            foreach (int addedCount in new[] { 0, 2 })
            {
                var project = new FakeProject();
                project.VBComponents.ImportAddedCount = addedCount;
                var service = Service(project);
                dynamic state = service.ProjectProperties(project.Name);
                string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-" + Guid.NewGuid().ToString("N") + ".bas");
                File.WriteAllText(path, "Attribute VB_Name = \"ImportedModule\"");
                try
                {
                    var error = Assert.ThrowsException<InvalidOperationException>(() => service.ImportComponent(
                        new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version }));
                    StringAssert.Contains(error.Message, "new components: " + addedCount);
                    Assert.AreEqual(1, project.VBComponents.ImportAttempts);
                }
                finally { File.Delete(path); }
            }
        }

        [TestMethod]
        public void ImportSuccessReturnsVerifiedComponentAndProject()
        {
            var project = new FakeProject();
            var service = Service(project);
            dynamic state = service.ProjectProperties(project.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-" + Guid.NewGuid().ToString("N") + ".bas");
            File.WriteAllText(path, "Attribute VB_Name = \"ImportedModule\"");
            try
            {
                dynamic result = service.ImportComponent(new Request {
                    Project = project.Name, Path = path, ExpectedProjectVersion = state.Version });
                Assert.IsTrue((bool)result.Applied);
                Assert.IsTrue((bool)result.Verified);
                Assert.IsFalse((bool)result.VerificationPending);
                Assert.IsNull((object)result.ImportError);
                Assert.IsNull((object)result.NextRead);
                Assert.AreEqual("ImportedModule", (string)result.ImportedName);
                Assert.AreEqual(1, project.VBComponents.ImportAttempts);
                Assert.AreEqual(1, project.VBComponents.Count());
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
        public void ExportCreatesFileAndNeverOverwritesIt()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1) { WriteExportFile = true };
            project.VBComponents.Add(module);
            var service = Service(project);
            dynamic state = service.ComponentProperties(project.Name, module.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-export-" + Guid.NewGuid().ToString("N") + ".bas");
            try
            {
                var request = new Request { Project = project.Name, Module = module.Name, Path = path,
                    ExpectedComponentVersion = state.Version };
                dynamic result = service.ExportComponent(request);
                Assert.AreEqual(path, (string)result.Path);
                Assert.IsTrue((long)result.Bytes > 0);
                Assert.AreEqual(1, module.ExportAttempts);
                Assert.ThrowsException<IOException>(() => service.ExportComponent(request));
                Assert.AreEqual(1, module.ExportAttempts);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        [TestMethod]
        public void FormExportPreservesExistingFrxCompanion()
        {
            var project = new FakeProject();
            var form = new FakeComponent("Form1", 3) { WriteExportFile = true };
            project.VBComponents.Add(form);
            var service = Service(project);
            dynamic state = service.ComponentProperties(project.Name, form.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-export-" + Guid.NewGuid().ToString("N") + ".frm");
            string companion = Path.ChangeExtension(path, ".frx");
            File.WriteAllText(companion, "existing binary companion");
            try
            {
                var error = Assert.ThrowsException<IOException>(() => service.ExportComponent(new Request {
                    Project = project.Name, Module = form.Name, Path = path,
                    ExpectedComponentVersion = state.Version }));
                StringAssert.Contains(error.Message, "FRX companion");
                Assert.AreEqual(0, form.ExportAttempts);
                Assert.IsFalse(File.Exists(path));
                Assert.AreEqual("existing binary companion", File.ReadAllText(companion));
            }
            finally { File.Delete(companion); if (File.Exists(path)) File.Delete(path); }
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

        [TestMethod]
        public void ComponentProbeCoversIdentityDescriptorsCodeAndNamedProperties()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1);
            project.VBComponents.Add(module);
            var service = Service(project);
            dynamic identity = service.ComponentProbe(project.Name, module.Name, "identity", null);
            Assert.AreEqual("Module1", (string)identity.Name);
            Assert.AreEqual(1, (int)identity.Type);
            var descriptors = ((IEnumerable)service.ComponentProbe(project.Name, module.Name,
                "descriptor_names", null)).Cast<object>().ToArray();
            Assert.IsTrue(descriptors.Any(item => (string)((dynamic)item).Name == "Description"));
            var propertyNames = ((IEnumerable)service.ComponentProbe(project.Name, module.Name,
                "designer_property_names", null)).Cast<string>().ToArray();
            CollectionAssert.Contains(propertyNames, "Instancing");
            Assert.AreEqual(1, (int)((dynamic)service.ComponentProbe(project.Name, module.Name,
                "code_count", null)).Lines);
            Assert.AreEqual(64, ((string)((dynamic)service.ComponentProbe(project.Name, module.Name,
                "code_sha", null)).Sha256).Length);
            dynamic descriptor = service.ComponentProbe(project.Name, module.Name,
                "descriptor_value", "Description");
            Assert.AreEqual("Original", (string)descriptor.Value);
            dynamic property = service.ComponentPropertyValue(project.Name, module.Name, "Instancing");
            Assert.AreEqual(1, (int)property.Value);
            Assert.ThrowsException<ArgumentException>(() => service.ComponentPropertyValue(
                project.Name, module.Name, " "));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(
                project.Name, module.Name, "descriptor_value", "Absent"));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(
                project.Name, module.Name, "unknown", null));
        }

        [TestMethod]
        public void DocumentComponentDoesNotReadHostPropertiesInBulkOrProbeMailEnvelope()
        {
            var project = new FakeProject();
            var document = new FakeComponent("ThisWorkbook", 100);
            document.Properties.Add(new FakeProperty { Name = "MailEnvelope", Value = "unsafe" });
            project.VBComponents.Add(document);
            var service = Service(project);
            dynamic snapshot = service.ComponentProperties(project.Name, document.Name);
            var host = ((IEnumerable)snapshot.HostProperties).Cast<VbePropertyInfo>().ToArray();
            Assert.AreEqual(2, host.Length);
            Assert.IsTrue(host.All(property => property.Value == null));
            StringAssert.Contains(host.Single(property => property.Name == "MailEnvelope").Error, "blocks");
            Assert.ThrowsException<InvalidOperationException>(() => service.ComponentPropertyValue(
                project.Name, document.Name, "MailEnvelope"));
        }

        [TestMethod]
        public void NonExcelHostReportsSignatureAndPersistenceLimitsAndRejectsSaving()
        {
            var project = new FakeProject { Saved = false };
            var service = Service(project);
            dynamic signature = service.SignatureStatus(project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsNull((object)signature.Signed);
            dynamic persistence = service.PersistenceStatus(project.Name);
            Assert.IsFalse((bool)persistence.ProjectSaved);
            Assert.IsFalse((bool)persistence.HostAvailable);
            Assert.ThrowsException<ArgumentException>(() => service.SaveHostDocument(new Request {
                Project = project.Name, ExpectedHostPath = "relative.xlsm" }));
            Assert.ThrowsException<InvalidOperationException>(() => service.SaveHostDocument(new Request {
                Project = project.Name, ExpectedHostPath = @"C:\fixture\Book.xlsm" }));
            Assert.ThrowsException<ArgumentException>(() => service.SaveHostDocumentAs(new Request {
                Project = project.Name, Path = " " }));
            Assert.ThrowsException<InvalidOperationException>(() => service.SaveHostDocumentAs(new Request {
                Project = project.Name, Path = @"C:\fixture\New.xlsm", ExpectedProjectVersion = "version" }));
            dynamic saved = service.PersistExcelSignature(project.Name);
            Assert.IsFalse((bool)saved.Available);
            Assert.IsFalse((bool)saved.Saved);
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
            public int ImportAddedCount { get; set; } = 1;
            public void Add(FakeComponent component) { items.Add(component); }
            public void Remove(FakeComponent component) { RemoveAttempts++; items.Remove(component); }
            public void Import(string path)
            {
                ImportAttempts++;
                for (int index = 0; index < ImportAddedCount; index++)
                    items.Add(new FakeComponent(index == 0 ? "ImportedModule" : "ImportedModule" + index, 1));
                if (ImportThenThrow) throw new InvalidOperationException("COM error after add");
            }
            public IEnumerator<FakeComponent> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeComponent
        {
            public FakeComponent(string name, int type)
            {
                Name = name;
                Type = type;
                if (type == 3)
                {
                    Designer = new FakeDesigner();
                    Properties.Add(new FakeProperty { Name = "Caption", Value = name });
                    Properties.Add(new FakeProperty { Name = "Width", Value = 240d });
                    Properties.Add(new FakeProperty { Name = "Height", Value = 180d });
                }
            }
            public string Name { get; set; }
            public string Description { get; set; } = "Original";
            public int Type { get; set; }
            public FakePropertyCollection Properties { get; } = new FakePropertyCollection();
            public FakeDesigner Designer { get; }
            public FakeCodeModule CodeModule { get; } = new FakeCodeModule();
            public int ExportAttempts { get; private set; }
            public bool WriteExportFile { get; set; }
            public void Export(string path)
            {
                ExportAttempts++;
                if (WriteExportFile) File.WriteAllText(path, "Attribute VB_Name = \"" + Name + "\"");
            }
        }

        public sealed class FakeDesigner
        {
            public List<object> Controls { get; } = new List<object>();
        }

        public sealed class FakePropertyCollection : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> items = new List<FakeProperty> {
                new FakeProperty { Name = "Instancing", Value = 1 }
            };
            public FakeProperty Item(string name) { return items.Single(p => p.Name == name); }
            public void Add(FakeProperty property) { items.Add(property); }
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
