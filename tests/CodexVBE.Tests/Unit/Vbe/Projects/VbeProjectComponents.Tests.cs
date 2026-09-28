namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeProjectExcelHostTests
    {
        [TestMethod]
        public void SignaturePersistenceRejectsWrongProcessMissingPathsAndAmbiguousWorkbooks()
        {
            foreach (int fault in new[] { 0, 1, 2, 3, 4, 5 })
            {
                var f = Create(); f.Workbook.VBASigned = true;
                if (fault == 0) f.Project.FileName = " ";
                if (fault == 1) f.Project.FileName = "relative.xlsm";
                if (fault == 2) f.Host.WindowOwner = 123;
                if (fault == 3) f.Workbook.FullName = Path.Combine(Path.GetTempPath(), "another.xlsm");
                if (fault == 4) f.Excel.Workbooks.Clear();
                if (fault == 5) f.Excel.Workbooks.Add(new FakeWorkbook(f.Project) { FullName = f.Project.FileName, VBASigned = true });
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.PersistExcelSignature(f.Project.Name));
                Assert.AreEqual(0, f.Workbook.SaveAttempts);
            }
        }

        [TestMethod]
        public void StatusAndSaveHandleUnsavedAmbiguousAndFailedProjectPaths()
        {
            var f = Create(); f.Project.FailFileName = true;
            dynamic status = f.Service.SignatureStatus(f.Project.Name); Assert.IsTrue((bool)status.Available);
            dynamic persistence = f.Service.PersistenceStatus(f.Project.Name); Assert.IsTrue((bool)persistence.HostAvailable);
            f.Excel.Workbooks.Add(new FakeWorkbook(f.Project) { FullName = "Other" });
            status = f.Service.SignatureStatus(f.Project.Name); Assert.IsFalse((bool)status.Available);
            persistence = f.Service.PersistenceStatus(f.Project.Name); Assert.IsFalse((bool)persistence.HostAvailable);
            f.Project.FailFileName = false;
            f.Project.FileName = "relative.xlsm";
            persistence = f.Service.PersistenceStatus(f.Project.Name); Assert.IsFalse((bool)persistence.HostAvailable);
            f.Project.FileName = f.Workbook.FullName;
            f.Excel.Workbooks.Add(new FakeWorkbook(f.Project) { FullName = f.Project.FileName });
            persistence = f.Service.PersistenceStatus(f.Project.Name); Assert.IsFalse((bool)persistence.HostAvailable);
            var single = Create(""); single.Vbe.VBProjects.Add(new VbeProjectComponentsTests.FakeProject { Name = "Other" });
            status = single.Service.SignatureStatus(single.Project.Name); Assert.IsFalse((bool)status.Available);
            Assert.ThrowsException<ArgumentException>(() => single.Service.SaveHostDocument(null));
            Assert.ThrowsException<ArgumentException>(() => single.Service.SaveHostDocument(new Request { ExpectedHostPath = "relative.xlsm" }));
        }

        [TestMethod]
        public void SaveRejectsRelativeProjectPathsAndIndependentProjectSaveCancellation()
        {
            var f = Create(); dynamic before = f.Service.ProjectProperties(f.Project.Name);
            var r = new Request { Project = f.Project.Name, ExpectedHostPath = f.Project.FileName, ExpectedProjectVersion = before.Version };
            f.Project.FileName = "relative.xlsm"; before = f.Service.ProjectProperties(f.Project.Name); r.ExpectedProjectVersion = before.Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveHostDocument(r));
            f.Project.FileName = r.ExpectedHostPath; f.Project.Saved = false; f.Workbook.Saved = true; f.Workbook.CommitSave = false;
            before = f.Service.ProjectProperties(f.Project.Name); r.ExpectedProjectVersion = before.Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveHostDocument(r));
        }

        [TestMethod]
        public void SaveAsValidatesEveryPathAndEveryPostSaveReadback()
        {
            string root = Path.Combine(Path.GetTempPath(), "CodexVBE-saveas-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                var f = Create("");
                Assert.ThrowsException<ArgumentException>(() => f.Service.SaveHostDocumentAs(null));
                Assert.ThrowsException<ArgumentException>(() => f.Service.SaveHostDocumentAs(new Request { Path = " " }));
                Assert.ThrowsException<ArgumentException>(() => f.Service.SaveHostDocumentAs(new Request { Path = Path.Combine(root, "book.xlsm") }));
                Assert.ThrowsException<ArgumentException>(() => f.Service.SaveHostDocumentAs(new Request { Path = Path.Combine(root, "book.xlsx"), ExpectedProjectVersion = "version" }));
                Assert.ThrowsException<DirectoryNotFoundException>(() => f.Service.SaveHostDocumentAs(new Request { Path = Path.Combine(root, "absent", "book.xlsm"), ExpectedProjectVersion = "version" }));
                for (int fault = 0; fault < 5; fault++)
                {
                    f = Create(""); string path = Path.Combine(root, "readback" + fault + ".xlsm"); var current = f;
                    f.Workbook.AfterSaveAs = () =>
                    {
                        if (fault == 0) current.Workbook.FullName = Path.Combine(root, "wrong.xlsm");
                        if (fault == 1) current.Project.FileName = Path.Combine(root, "wrong.xlsm");
                        if (fault == 2) File.Delete(path);
                        if (fault == 3) current.Workbook.Saved = false;
                        if (fault == 4) current.Project.Saved = false;
                    };
                    dynamic before = f.Service.ProjectProperties(f.Project.Name);
                    Assert.ThrowsException<InvalidOperationException>(() => current.Service.SaveHostDocumentAs(new Request { Project = current.Project.Name, Path = path, ExpectedProjectVersion = before.Version }));
                    Assert.AreEqual(1, current.Workbook.SaveAsAttempts);
                }
                f = Create(""); f.Project.FailFileName = true; dynamic state = f.Service.ProjectProperties(f.Project.Name);
                f.Workbook.AfterSaveAs = () => f.Project.FailFileName = false;
                dynamic result = f.Service.SaveHostDocumentAs(new Request { Project = f.Project.Name, Path = Path.Combine(root, "first.xlsm"), ExpectedProjectVersion = state.Version });
                Assert.IsTrue((bool)result.SaveAsInvoked);
            }
            finally { Directory.Delete(root, true); }
        }
    }

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProjectComponentsTests
    {
        [TestMethod]
        public void ComponentProbeCoversMissingDescriptorsNullValuesErrorsAndEmptyCode()
        {
            var project = new FakeProject(); var component = new MetadataComponent(); project.VBComponents.Add(component); var service = Service(project);
            component.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new MetadataProperty("Null", null), new MetadataProperty("Empty", typeof(string)) });
            dynamic names = service.ComponentProbe(project.Name, component.Name, "descriptor_names", null); Assert.AreEqual(2, names.Length);
            dynamic empty = service.ComponentProbe(project.Name, component.Name, "descriptor_value", "Empty"); Assert.IsNull((object)empty.Value);
            dynamic nullType = service.ComponentProbe(project.Name, component.Name, "descriptor_value", "Null"); Assert.IsNull((string)nullType.Type);
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(project.Name, component.Name, "descriptor_value", "Missing"));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(project.Name, component.Name, "unsupported", null));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentPropertyValue(project.Name, component.Name, " "));
            component.CodeModule.LineCount = 0;
            dynamic hash = service.ComponentProbe(project.Name, component.Name, "code_sha", null); Assert.AreEqual(0, (int)hash.Lines);
            component.Properties.Add(new FakeProperty { Name = "Empty", Value = null });
            dynamic property = service.ComponentPropertyValue(project.Name, component.Name, "Empty"); Assert.IsNull((object)property.Value);
            component.Properties.Add(new FakeProperty { Name = "Unavailable", FailRead = true });
            dynamic snapshot = service.ComponentProperties(project.Name, component.Name);
            Assert.IsTrue(((IEnumerable)snapshot.DesignerProperties).Cast<VbePropertyInfo>().Any(p => p.Error != null));
        }

        [TestMethod]
        public void ComPropertiesAreReportedAsObjectsWithoutSerializingNativeState()
        {
            object com = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            try
            {
                Assert.IsTrue(System.Runtime.InteropServices.Marshal.IsComObject(com));
                var project = new FakeProject(); var component = new MetadataComponent(); project.VBComponents.Add(component); var service = Service(project);
                component.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new MetadataProperty("Native", typeof(object), com), new MetadataProperty("InvalidScalarMetadata", typeof(int), com) });
                component.Properties.Add(new FakeProperty { Name = "Native", Value = com });
                dynamic descriptor = service.ComponentProbe(project.Name, component.Name, "descriptor_value", "Native");
                Assert.AreEqual("object", (string)descriptor.Kind); Assert.IsNull((object)descriptor.Value);
                dynamic designer = service.ComponentPropertyValue(project.Name, component.Name, "Native");
                Assert.AreEqual("object", (string)designer.Kind); Assert.IsNull((object)designer.Value);
                dynamic snapshot = service.ComponentProperties(project.Name, component.Name);
                Assert.AreEqual("object", ((IEnumerable)snapshot.DesignerProperties).Cast<VbePropertyInfo>().Single(p => p.Name == "Native").Kind);
                var properties = (List<VbePropertyInfo>)PrivateProjectMethod("ReadProperties", component);
                Assert.AreEqual("object", properties.Single(p => p.Name == "InvalidScalarMetadata").Kind);
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
        }

        [TestMethod]
        public void ImportReadbackFailureDoesNotRetryOrClaimVerifiedSuccess()
        {
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-" + Guid.NewGuid().ToString("N") + ".bas"); File.WriteAllText(path, "Sub Run()\r\nEnd Sub");
            try
            {
                for (int fault = 0; fault < 2; fault++)
                {
                    var project = new FakeProject(); var service = Service(project); dynamic before = service.ProjectProperties(project.Name);
                    project.VBComponents.FailImportedEnumeration = fault == 0;
                    project.VBComponents.FailImportedCodeReadback = fault == 1;
                    var request = new Request { Project = project.Name, Path = path, ExpectedProjectVersion = before.Version };
                    if (fault == 0) StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.ImportComponent(request)).Message, "uncertain");
                    else
                    {
                        dynamic result = service.ImportComponent(request);
                        Assert.IsTrue((bool)result.Applied); Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.VerificationPending);
                        Assert.IsNotNull((string)result.ComponentReadbackError); Assert.IsNotNull((string)result.NextRead);
                    }
                    Assert.AreEqual(1, project.VBComponents.ImportAttempts);
                }
                foreach (string invalid in new[] { null, "relative.bas", path + ".absent" })
                {
                    var project = new FakeProject(); var service = Service(project); var request = new Request { Project = project.Name, Path = invalid };
                    if (invalid == path + ".absent") Assert.ThrowsException<FileNotFoundException>(() => service.ImportComponent(request));
                    else Assert.ThrowsException<ArgumentException>(() => service.ImportComponent(request));
                    Assert.AreEqual(0, project.VBComponents.ImportAttempts);
                }
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void ConstructorRejectsMissingHostProbeAndNativeProbeUsesTheCurrentProcess()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new VbeProjectComponents(null, null, null));
            var type = typeof(VbeProjectComponents).GetNestedType("NativeExcelHostProbe", System.Reflection.BindingFlags.NonPublic);
            var host = (VbeProjectComponents.IExcelHostProbe)Activator.CreateInstance(type, true);
            Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().Id, host.CurrentProcessId);
            Assert.AreEqual(string.Equals(System.Diagnostics.Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase), host.IsExcel);
            using (var form = new System.Windows.Forms.Form()) Assert.AreEqual((uint)host.CurrentProcessId, host.WindowProcessId(form.Handle));
            try { Assert.IsNotNull(host.ExcelApplication()); }
            catch (System.Runtime.InteropServices.COMException error) { Assert.AreNotEqual(0, error.ErrorCode); }
        }

        [TestMethod]
        public void ScalarPropertyEditingValidatesMetadataConversionsAndReadback()
        {
            var target = new MetadataComponent();
            var text = new MetadataProperty("Text", typeof(string), "old");
            var numeric = new MetadataProperty("Number", typeof(int), 1);
            var choice = new MetadataProperty("Choice", typeof(Choice), Choice.First);
            var decimalValue = new MetadataProperty("Amount", typeof(decimal), 1m);
            var boxed = new MetadataProperty("Boxed", typeof(object), 1);
            var empty = new MetadataProperty("Empty", typeof(object));
            var objectValue = new MetadataProperty("Object", typeof(Uri), new Uri("https://example.com"));
            target.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { text, numeric, choice, decimalValue, boxed, empty, objectValue });
            foreach (string property in new[] { null, " " }) Assert.ThrowsException<ArgumentException>(() => PrivateProjectMethod("SetScalar", target, property, "value"));
            Assert.ThrowsException<ArgumentException>(() => PrivateProjectMethod("SetScalar", target, "Text", null));
            Assert.ThrowsException<InvalidOperationException>(() => PrivateProjectMethod("SetScalar", target, "Missing", "value"));
            text.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => PrivateProjectMethod("SetScalar", target, "Text", "value")); text.ReadOnly = false;
            Assert.ThrowsException<ArgumentException>(() => PrivateProjectMethod("SetScalar", target, "Text", 1));
            Assert.ThrowsException<InvalidOperationException>(() => PrivateProjectMethod("SetScalar", target, "Object", "value"));
            PrivateProjectMethod("SetScalar", target, "Number", "2"); Assert.AreEqual(2, numeric.Value);
            PrivateProjectMethod("SetScalar", target, "Amount", "2.5"); Assert.AreEqual(2.5m, decimalValue.Value);
            PrivateProjectMethod("SetScalar", target, "Choice", "Second"); Assert.AreEqual(Choice.Second, choice.Value);
            PrivateProjectMethod("SetScalar", target, "Choice", 1); Assert.AreEqual(Choice.First, choice.Value);
            PrivateProjectMethod("SetScalar", target, "Boxed", 3L); Assert.AreEqual(3, boxed.Value);
            PrivateProjectMethod("SetScalar", target, "Empty", "new"); Assert.AreEqual("new", empty.Value);
            text.NormalizeString = true; PrivateProjectMethod("SetScalar", target, "Text", "new"); Assert.AreEqual("NEW", text.Value);
            text.IgnoreWrite = true;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => PrivateProjectMethod("SetScalar", target, "Text", "different")).Message, "did not retain");
        }

        [TestMethod]
        public void MetadataInspectionHandlesNullableTypesScalarsAndGetterFailures()
        {
            var target = new MetadataComponent();
            target.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
            {
                new MetadataProperty("NullType", null), new MetadataProperty("NullString", typeof(string)),
                new MetadataProperty("Failure", typeof(string)) { FailRead = true },
                new MetadataProperty("Date", typeof(DateTime), new DateTime(2026, 1, 1)),
                new MetadataProperty("Enum", typeof(Choice), Choice.First), new MetadataProperty("Decimal", typeof(decimal), 3m)
            });
            var properties = (List<VbePropertyInfo>)PrivateProjectMethod("ReadProperties", target);
            Assert.AreEqual("object", properties.Single(p => p.Name == "NullType").Kind);
            Assert.IsNull(properties.Single(p => p.Name == "NullString").Value);
            StringAssert.Contains(properties.Single(p => p.Name == "Failure").Error, "Unreadable");
            Assert.AreEqual("First", properties.Single(p => p.Name == "Enum").Value);
            foreach (object value in new object[] { null, "text", true, (byte)1, (short)2, 3, 4L, 5f, 6d, 7m }) Assert.AreEqual(value, PrivateProjectMethod("Scalar", value));
            Assert.AreEqual("x", PrivateProjectMethod("Scalar", 'x'));
            Assert.IsFalse((bool)PrivateProjectMethod("IsSafeScalarType", new object[] { null }));
        }

        [TestMethod]
        public void ProjectsIncludeReferenceIdentityAndComponentsRejectMissingOrAmbiguousNames()
        {
            var project = new FakeProject();
            project.References.Add(new ProjectReference { GUID = "guid", Major = 1, Minor = 2 });
            project.VBComponents.Add(new FakeComponent("Module1", 1));
            var service = Service(project);
            dynamic snapshot = service.ProjectProperties(project.Name); Assert.AreEqual(1, snapshot.References.Count);
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProperties(project.Name, null));
            Assert.ThrowsException<InvalidOperationException>(() => service.ComponentProperties(project.Name, "absent"));
            project.VBComponents.Add(new FakeComponent("module1", 1));
            Assert.ThrowsException<InvalidOperationException>(() => service.ComponentProperties(project.Name, "Module1"));
        }

        [TestMethod]
        public void MissingVersionsModesAndUnretainedNamesBlockComponentChanges()
        {
            var project = new FakeProject(); var component = new FakeComponent("Module1", 1); project.VBComponents.Add(component); var service = Service(project);
            Assert.ThrowsException<ArgumentException>(() => service.SetProjectProperty(new Request { Project = project.Name, Property = "Description", Value = "new" }));
            Assert.ThrowsException<ArgumentException>(() => service.SetComponentProperty(new Request { Project = project.Name, Module = component.Name, Property = "Description", Value = "new" }));
            dynamic before = service.ComponentProperties(project.Name, component.Name);
            var request = new Request { Project = project.Name, Module = component.Name, Property = "Name", Value = "NewName", ExpectedComponentVersion = before.Version };
            service.SetComponentProperty(request); Assert.AreEqual("NewName", component.Name);
            component.IgnoreRename = true;
            before = service.ComponentProperties(project.Name, component.Name);
            request.Module = component.Name; request.NewName = "Other"; request.ExpectedComponentVersion = before.Version;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.RenameComponent(request)).Message, "did not retain");
            request.NewName = null;
            Assert.ThrowsException<ArgumentException>(() => service.RenameComponent(request));
            request.NewName = "Other";
            project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => service.RenameComponent(request));
        }

        [TestMethod]
        public void InstancingAcceptsNumericRepresentationsAndRejectsOverflowOrIgnoredSetter()
        {
            var project = new FakeProject(); var component = new FakeComponent("Class1", 2); project.VBComponents.Add(component); var service = Service(project);
            foreach (object value in new object[] { 1L, 2d, 1m, 2 })
            {
                dynamic before = service.ComponentProperties(project.Name, component.Name);
                dynamic result = service.SetClassInstancing(new Request { Project = project.Name, Module = component.Name, Value = value, ExpectedComponentVersion = before.Version });
                Assert.AreEqual(Convert.ToInt32(value), (int)result.Instancing);
            }
            foreach (object value in new object[] { null, "1", true, 0, 3, double.MaxValue })
                Assert.ThrowsException<ArgumentException>(() => service.SetClassInstancing(new Request { Value = value }));
            component.Properties.Item("Instancing").IgnoreSet = true;
            dynamic state = service.ComponentProperties(project.Name, component.Name);
            Assert.ThrowsException<InvalidOperationException>(() => service.SetClassInstancing(new Request { Project = project.Name, Module = component.Name, Value = 1, ExpectedComponentVersion = state.Version }));
        }

        [TestMethod]
        public void ImportPreservesExistingInventoryAndReportsComFailureWithoutNewComponents()
        {
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-existing-" + Guid.NewGuid().ToString("N") + ".bas"); File.WriteAllText(path, "Option Explicit");
            try
            {
                var project = new FakeProject(); project.VBComponents.Add(new FakeComponent("Existing", 1)); var service = Service(project);
                dynamic state = service.ProjectProperties(project.Name);
                dynamic applied = service.ImportComponent(new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version });
                Assert.IsTrue((bool)applied.Verified);
                Assert.IsTrue(project.VBComponents.Any(c => c.Name == "Existing"));
                project = new FakeProject(); project.VBComponents.ImportAddedCount = 0; project.VBComponents.ImportThenThrow = true; service = Service(project);
                state = service.ProjectProperties(project.Name);
                var error = Assert.ThrowsException<InvalidOperationException>(() => service.ImportComponent(new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version }));
                StringAssert.Contains(error.Message, "COM error after add"); Assert.AreEqual(1, project.VBComponents.ImportAttempts);
            }
            finally { File.Delete(path); }
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
            var request = new Request
            {
                Project = "missing",
                Property = "Name",
                Value = "Renamed"
            };
            var error = Assert.ThrowsException<InvalidOperationException>(() => service.SetProjectProperty(request));
            StringAssert.Contains(error.Message, "rename is disabled");
        }

        [TestMethod]
        public void ProjectPropertyRequiresFreshVersionAndPersistsReadback()
        {
            var project = new FakeProject();
            var service = Service(project);
            var request = new Request
            {
                Project = project.Name,
                Property = "Description",
                Value = "New description",
                ExpectedProjectVersion = "stale"
            };
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
            var request = new Request
            {
                Project = project.Name,
                Module = module.Name,
                Property = "Description",
                Value = "Changed",
                ExpectedComponentVersion = "stale"
            };
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
            var request = new Request
            {
                Project = project.Name,
                Module = module.Name,
                NewName = "../escape"
            };
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
            var request = new Request
            {
                Project = project.Name,
                Module = document.Name,
                ExpectedProjectVersion = projectState.Version,
                ExpectedComponentVersion = componentState.Version
            };
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
            var request = new Request
            {
                Project = project.Name,
                Module = module.Name,
                ExpectedProjectVersion = "stale",
                ExpectedComponentVersion = componentState.Version
            };
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
                dynamic result = service.ImportComponent(new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version });
                Assert.IsTrue((bool)result.Applied);
                Assert.IsTrue((bool)result.Verified);
                Assert.AreEqual("ImportedModule", (string)result.ImportedName);
                StringAssert.Contains((string)result.ImportError, "after add");
                Assert.AreEqual(1, project.VBComponents.ImportAttempts);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void ImportWithoutExactlyOneNewComponentReportsUncertainOutcome()
        {
            foreach (int addedCount in new[]
            {
                0,
                2
            }

            )
            {
                var project = new FakeProject();
                project.VBComponents.ImportAddedCount = addedCount;
                var service = Service(project);
                dynamic state = service.ProjectProperties(project.Name);
                string path = Path.Combine(Path.GetTempPath(), "CodexVBE-import-" + Guid.NewGuid().ToString("N") + ".bas");
                File.WriteAllText(path, "Attribute VB_Name = \"ImportedModule\"");
                try
                {
                    var error = Assert.ThrowsException<InvalidOperationException>(() => service.ImportComponent(new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version }));
                    StringAssert.Contains(error.Message, "new components: " + addedCount);
                    Assert.AreEqual(1, project.VBComponents.ImportAttempts);
                }
                finally
                {
                    File.Delete(path);
                }
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
                dynamic result = service.ImportComponent(new Request { Project = project.Name, Path = path, ExpectedProjectVersion = state.Version });
                Assert.IsTrue((bool)result.Applied);
                Assert.IsTrue((bool)result.Verified);
                Assert.IsFalse((bool)result.VerificationPending);
                Assert.IsNull((object)result.ImportError);
                Assert.IsNull((object)result.NextRead);
                Assert.AreEqual("ImportedModule", (string)result.ImportedName);
                Assert.AreEqual(1, project.VBComponents.ImportAttempts);
                Assert.AreEqual(1, project.VBComponents.Count());
            }
            finally
            {
                File.Delete(path);
            }
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
                var error = Assert.ThrowsException<IOException>(() => service.ExportComponent(new Request { Project = project.Name, Module = module.Name, Path = path, ExpectedComponentVersion = state.Version }));
                StringAssert.Contains(error.Message, "did not create");
                Assert.AreEqual(1, module.ExportAttempts);
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [TestMethod]
        public void ExportCreatesFileAndNeverOverwritesIt()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Module1", 1)
            {
                WriteExportFile = true
            };
            project.VBComponents.Add(module);
            var service = Service(project);
            dynamic state = service.ComponentProperties(project.Name, module.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-export-" + Guid.NewGuid().ToString("N") + ".bas");
            try
            {
                var request = new Request
                {
                    Project = project.Name,
                    Module = module.Name,
                    Path = path,
                    ExpectedComponentVersion = state.Version
                };
                dynamic result = service.ExportComponent(request);
                Assert.AreEqual(path, (string)result.Path);
                Assert.IsTrue((long)result.Bytes > 0);
                Assert.AreEqual(1, module.ExportAttempts);
                Assert.ThrowsException<IOException>(() => service.ExportComponent(request));
                Assert.AreEqual(1, module.ExportAttempts);
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [TestMethod]
        public void FormExportPreservesExistingFrxCompanion()
        {
            var project = new FakeProject();
            var form = new FakeComponent("Form1", 3)
            {
                WriteExportFile = true
            };
            project.VBComponents.Add(form);
            var service = Service(project);
            dynamic state = service.ComponentProperties(project.Name, form.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-export-" + Guid.NewGuid().ToString("N") + ".frm");
            string companion = Path.ChangeExtension(path, ".frx");
            File.WriteAllText(companion, "existing binary companion");
            try
            {
                var error = Assert.ThrowsException<IOException>(() => service.ExportComponent(new Request { Project = project.Name, Module = form.Name, Path = path, ExpectedComponentVersion = state.Version }));
                StringAssert.Contains(error.Message, "FRX companion");
                Assert.AreEqual(0, form.ExportAttempts);
                Assert.IsFalse(File.Exists(path));
                Assert.AreEqual("existing binary companion", File.ReadAllText(companion));
            }
            finally
            {
                File.Delete(companion);
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [TestMethod]
        public void ClassInstancingRejectsFractionAndNonClassThenVerifiesValue()
        {
            var project = new FakeProject();
            var module = new FakeComponent("Class1", 2);
            project.VBComponents.Add(module);
            var service = Service(project);
            var request = new Request
            {
                Project = project.Name,
                Module = module.Name,
                Value = 1.5
            };
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
            var descriptors = ((IEnumerable)service.ComponentProbe(project.Name, module.Name, "descriptor_names", null)).Cast<object>().ToArray();
            Assert.IsTrue(descriptors.Any(item => (string)((dynamic)item).Name == "Description"));
            var propertyNames = ((IEnumerable)service.ComponentProbe(project.Name, module.Name, "designer_property_names", null)).Cast<string>().ToArray();
            CollectionAssert.Contains(propertyNames, "Instancing");
            Assert.AreEqual(1, (int)((dynamic)service.ComponentProbe(project.Name, module.Name, "code_count", null)).Lines);
            Assert.AreEqual(64, ((string)((dynamic)service.ComponentProbe(project.Name, module.Name, "code_sha", null)).Sha256).Length);
            dynamic descriptor = service.ComponentProbe(project.Name, module.Name, "descriptor_value", "Description");
            Assert.AreEqual("Original", (string)descriptor.Value);
            dynamic property = service.ComponentPropertyValue(project.Name, module.Name, "Instancing");
            Assert.AreEqual(1, (int)property.Value);
            Assert.ThrowsException<ArgumentException>(() => service.ComponentPropertyValue(project.Name, module.Name, " "));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(project.Name, module.Name, "descriptor_value", "Absent"));
            Assert.ThrowsException<ArgumentException>(() => service.ComponentProbe(project.Name, module.Name, "unknown", null));
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
            Assert.ThrowsException<InvalidOperationException>(() => service.ComponentPropertyValue(project.Name, document.Name, "MailEnvelope"));
        }

        [TestMethod]
        public void NonExcelHostReportsSignatureAndPersistenceLimitsAndRejectsSaving()
        {
            var project = new FakeProject
            {
                Saved = false
            };
            var service = Service(project);
            dynamic signature = service.SignatureStatus(project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsNull((object)signature.Signed);
            dynamic persistence = service.PersistenceStatus(project.Name);
            Assert.IsFalse((bool)persistence.ProjectSaved);
            Assert.IsFalse((bool)persistence.HostAvailable);
            Assert.ThrowsException<ArgumentException>(() => service.SaveHostDocument(new Request { Project = project.Name, ExpectedHostPath = "relative.xlsm" }));
            Assert.ThrowsException<InvalidOperationException>(() => service.SaveHostDocument(new Request { Project = project.Name, ExpectedHostPath = @"C:\fixture\Book.xlsm" }));
            Assert.ThrowsException<ArgumentException>(() => service.SaveHostDocumentAs(new Request { Project = project.Name, Path = " " }));
            Assert.ThrowsException<InvalidOperationException>(() => service.SaveHostDocumentAs(new Request { Project = project.Name, Path = @"C:\fixture\New.xlsm", ExpectedProjectVersion = "version" }));
            dynamic saved = service.PersistExcelSignature(project.Name);
            Assert.IsFalse((bool)saved.Available);
            Assert.IsFalse((bool)saved.Saved);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProjectExcelHostTests
    {
        [TestMethod]
        public void StatusReadsTheMatchingWorkbookAndReportsUnavailableStates()
        {
            var fixture = Create();
            fixture.Workbook.VBASigned = true;
            fixture.Project.Saved = false;
            fixture.Workbook.Saved = false;
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            Assert.IsTrue((bool)signature.Available);
            Assert.AreEqual(true, (bool? )signature.Signed);
            dynamic persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsTrue((bool)persistence.HostAvailable);
            Assert.IsFalse((bool)persistence.ProjectSaved);
            Assert.AreEqual(false, (bool? )persistence.HostSaved);
            Assert.AreEqual(fixture.Project.FileName, (string)persistence.HostPath);
            fixture.Workbook.FullName = Path.Combine(Path.GetTempPath(), "different.xlsm");
            signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsFalse((bool)persistence.HostAvailable);
            StringAssert.Contains((string)persistence.Reason, "No workbook matches");
        }

        [TestMethod]
        public void StatusRejectsAnotherExcelProcessAndHandlesComLookupFailure()
        {
            var fixture = Create();
            fixture.Host.WindowOwner = 99;
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            dynamic persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)signature.Available);
            Assert.IsFalse((bool)persistence.HostAvailable);
            StringAssert.Contains((string)signature.Reason, "not this VBE host");
            fixture.Host.WindowOwner = (uint)fixture.Host.CurrentProcessId;
            fixture.Host.LookupFailure = new InvalidOperationException("ROT unavailable");
            signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            persistence = fixture.Service.PersistenceStatus(fixture.Project.Name);
            StringAssert.Contains((string)signature.Reason, "ROT unavailable");
            StringAssert.Contains((string)persistence.Reason, "ROT unavailable");
        }

        [TestMethod]
        public void SingleUnsavedWorkbookCanBeMatchedWithoutInventingAHostPath()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic status = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsTrue((bool)status.HostAvailable);
            Assert.AreEqual(false, (bool? )status.HostHasPath);
            Assert.IsNull((object)status.HostPath);
            dynamic signature = fixture.Service.SignatureStatus(fixture.Project.Name);
            Assert.IsTrue((bool)signature.Available);
            fixture.Excel.Workbooks.Add(new FakeWorkbook(fixture.Project) { FullName = "Book2", Path = "" });
            status = fixture.Service.PersistenceStatus(fixture.Project.Name);
            Assert.IsFalse((bool)status.HostAvailable);
            StringAssert.Contains((string)status.Reason, "no saved workbook path");
        }

        [TestMethod]
        public void SaveRequiresExactCurrentPathAndVerifiedSavedReadback()
        {
            var fixture = Create();
            fixture.Project.Saved = false;
            fixture.Workbook.Saved = false;
            dynamic projectState = fixture.Service.ProjectProperties(fixture.Project.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                ExpectedProjectVersion = projectState.Version,
                ExpectedHostPath = fixture.Project.FileName
            };
            dynamic saved = fixture.Service.SaveHostDocument(request);
            Assert.IsTrue((bool)saved.SaveInvoked);
            Assert.IsFalse((bool)saved.HostSavedBefore);
            Assert.IsFalse((bool)saved.ProjectSavedBefore);
            Assert.IsTrue((bool)saved.HostSaved);
            Assert.IsTrue((bool)saved.ProjectSaved);
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
            request.ExpectedHostPath = Path.Combine(Path.GetTempPath(), "different.xlsm");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SaveRefusesReadOnlyAndCancelledHostSave()
        {
            var fixture = Create();
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                ExpectedProjectVersion = state.Version,
                ExpectedHostPath = fixture.Project.FileName
            };
            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            Assert.AreEqual(0, fixture.Workbook.SaveAttempts);
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.CommitSave = false;
            fixture.Workbook.Saved = false;
            fixture.Project.Saved = false;
            request.ExpectedProjectVersion = ((dynamic)fixture.Service.ProjectProperties(fixture.Project.Name)).Version;
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocument(request));
            StringAssert.Contains(error.Message, "did not mark");
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SaveAsFirstTimeWritesMacroWorkbookAndVerifiesBothPaths()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-SaveAs-" + Guid.NewGuid().ToString("N") + ".xlsm");
            try
            {
                dynamic result = fixture.Service.SaveHostDocumentAs(new Request { Project = fixture.Project.Name, ExpectedProjectVersion = state.Version, Path = path });
                Assert.AreEqual(path, (string)result.HostPath);
                Assert.AreEqual(path, (string)result.ProjectPath);
                Assert.IsTrue((bool)result.SaveAsInvoked);
                Assert.IsTrue((long)result.Bytes > 0);
                Assert.AreEqual(52, fixture.Workbook.LastSaveAsFormat);
                Assert.AreEqual(1, fixture.Workbook.SaveAsAttempts);
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [TestMethod]
        public void SaveAsRefusesExistingDestinationReadOnlyWorkbookAndSecondSave()
        {
            var fixture = Create("");
            fixture.Workbook.FullName = "Book1";
            fixture.Workbook.Path = "";
            dynamic state = fixture.Service.ProjectProperties(fixture.Project.Name);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-SaveAs-" + Guid.NewGuid().ToString("N") + ".xlsm");
            var request = new Request
            {
                Project = fixture.Project.Name,
                ExpectedProjectVersion = state.Version,
                Path = path
            };
            request.Path = Path.ChangeExtension(path, ".xlsx");
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SaveHostDocumentAs(request));
            request.Path = path;
            File.WriteAllText(path, "existing");
            try
            {
                Assert.ThrowsException<IOException>(() => fixture.Service.SaveHostDocumentAs(request));
            }
            finally
            {
                File.Delete(path);
            }

            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocumentAs(request));
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.Path = Path.GetTempPath();
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SaveHostDocumentAs(request));
            Assert.AreEqual(0, fixture.Workbook.SaveAsAttempts);
        }

        [TestMethod]
        public void SignedWorkbookPersistenceRequiresSignatureBeforeAndAfterSave()
        {
            var fixture = Create();
            fixture.Workbook.VBASigned = true;
            dynamic result = fixture.Service.PersistExcelSignature(fixture.Project.Name);
            Assert.IsTrue((bool)result.Available);
            Assert.IsTrue((bool)result.Saved);
            Assert.AreEqual(1, fixture.Workbook.SaveAttempts);
            fixture.Workbook.ReadOnly = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PersistExcelSignature(fixture.Project.Name));
            fixture.Workbook.ReadOnly = false;
            fixture.Workbook.VBASigned = false;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PersistExcelSignature(fixture.Project.Name));
            fixture.Workbook.VBASigned = true;
            fixture.Workbook.DropSignatureOnSave = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PersistExcelSignature(fixture.Project.Name));
            StringAssert.Contains(error.Message, "no longer reports");
            Assert.AreEqual(2, fixture.Workbook.SaveAttempts);
        }
    }
}
