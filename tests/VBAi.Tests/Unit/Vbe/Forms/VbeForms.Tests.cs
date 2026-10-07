namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed partial class VbeFormsCoverageTests
    {
        [TestMethod]
        public void FormCreationReportsRollbackAbsentRetainedOrUninspectableComponents()
        {
            for (int fault = 0; fault < 6; fault++)
            {
                var project = new VbeFormsTests.FakeProject(); var host = new VbeFormsTests.FakeVbe(); host.VBProjects.Add(project); var service = new VbeForms(host);
                project.VBComponents.RejectedCreatedName = "NewForm";
                if (fault == 0) project.VBComponents.BeforeAdd = () => { throw new InvalidOperationException("Add unavailable"); };
                if (fault == 1) project.VBComponents.IgnoreRemove = true;
                if (fault == 2 || fault == 4) project.VBComponents.BeforeRemove = () => { throw new InvalidOperationException("Remove unavailable"); };
                if (fault == 3 || fault == 4) project.VBComponents.BeforeEnumeration = () => { if (project.VBComponents.AddCalls > 0) throw new InvalidOperationException("Inspection unavailable"); };
                if (fault == 5)
                {
                    project.VBComponents.RejectedCreatedName = null; project.VBComponents.IgnoreRemove = true;
                    project.VBComponents.ConfigureAdded = created => created.FailDesignerWindow = true;
                }
                var error = Assert.ThrowsException<InvalidOperationException>(() => service.Create(new Request { Project = project.Name, Form = "NewForm" }));
                if (fault == 0) StringAssert.Contains(error.Message, "absent after rollback");
                else StringAssert.Contains(error.Message, "rollback could not be verified");
                Assert.IsNotNull(error.InnerException);
            }
        }

        [TestMethod]
        public void DirectControlCreationVerifiesRollbackAfterEveryNativeBoundaryFailure()
        {
            for (int fault = 0; fault < 6; fault++)
            {
                var form = new VbeFormsTests.FakeForm("Form1"); var project = new VbeFormsTests.FakeProject(); project.VBComponents.Add(form);
                var host = new VbeFormsTests.FakeVbe(); host.VBProjects.Add(project); var service = new VbeForms(host);
                var controls = form.Designer.Controls;
                var request = new Request { Project = project.Name, Form = form.Name, Control = "Label1", ControlType = "Forms.Label.1", Caption = "Label", Width = 20, Height = 10, ExpectedFormVersion = (string)((dynamic)service.State(project.Name, form.Name)).Version };
                controls.FailNextCaption = true;
                if (fault == 0) controls.FailAdd = true;
                if (fault == 1) controls.HideAdd = true;
                if (fault == 2) controls.IgnoreRemove = true;
                if (fault == 3 || fault == 5) controls.BeforeRemove = () => { throw new InvalidOperationException("Remove unavailable"); };
                if (fault == 4 || fault == 5) controls.BeforeEnumeration = () => { if (controls.AddCalls > 0) throw new InvalidOperationException("Inspection unavailable"); };
                var error = Assert.ThrowsException<InvalidOperationException>(() => service.AddControl(request));
                if (fault <= 1) StringAssert.Contains(error.Message, "no control with the requested name remains");
                else StringAssert.Contains(error.Message, "rollback could not be verified");
                Assert.IsNotNull(error.InnerException);
            }
        }

        [TestMethod]
        public void ControlDescriptionsReportFailedGettersAndNullTypesWithoutBlockingInspection()
        {
            var form = new Form(); var node = new Node { Name = "Label1", ClassName = "Label" };
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
            {
                new VbeProjectComponentsTests.MetadataProperty("NullType", null),
                new VbeProjectComponentsTests.MetadataProperty("Failed", typeof(string)) { FailRead = true },
                new VbeProjectComponentsTests.MetadataProperty("_Font_Reserved", typeof(object))
            });
            form.Designer.Controls = new object[] { node }; var service = Service(form);
            dynamic result = service.ControlProperties("VBAProject", form.Name, node.Name);
            Assert.AreEqual(3, result.Count);
            Assert.IsNull((string)result[0].Type); Assert.IsNotNull((string)result[1].Error); Assert.AreEqual("GetterUnavailable", (string)result[2].SetterStatus);
            using (var font = new System.Drawing.Font("Arial", 10))
            {
                node.FailFont = true;
                node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("Font", typeof(System.Drawing.Font), font) });
                var properties = (List<VbePropertyInfo>)Call("ReadObjectProperties", node);
                StringAssert.Contains(properties[0].Error, "Font unavailable");
            }
        }

        [TestMethod]
        public void FormLookupAndPropertyNamesRequireValidIdentityAndFreshVersions()
        {
            var project = new VbeFormsTests.FakeProject(); project.VBComponents.Add(new VbeFormsTests.FakeModule("Module1"));
            Assert.ThrowsException<ArgumentException>(() => Call("GetForm", project, null));
            Assert.ThrowsException<InvalidOperationException>(() => Call("GetForm", project, "Module1"));
            Assert.ThrowsException<InvalidOperationException>(() => Call("GetForm", project, "Absent"));
            foreach (string name in new[] { null, "0invalid" }) Assert.ThrowsException<ArgumentException>(() => Call("ValidateName", name, "Control"));
            var form = new Form(); var service = Service(form);
            var request = PropertyRequest(service, form, "Caption", "new"); request.ExpectedFormVersion = null;
            Assert.ThrowsException<ArgumentException>(() => service.SetProperty(request));
            foreach (object value in new object[] { null, 1, " ", "0invalid" })
                Assert.ThrowsException<ArgumentException>(() => service.SetProperty(PropertyRequest(service, form, "Name", value)));
            form.Designer.ClassName = "";
            Assert.IsNull(Call("ContainerIdentity", form.Designer, form.Name));
            form.Properties.Add(new Property { Name = "Custom", Stored = "string" });
            form.Designer.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("Custom", null, "string") });
            var props = (List<VbePropertyInfo>)service.Properties("VBAProject", form.Name);
            Assert.AreEqual(typeof(string).FullName, props.Single(p => p.Name == "Custom").Type);
            service = Service(form, new VbeFormsTests.FakeModule("Other"));
            Assert.ThrowsException<InvalidOperationException>(() => service.SetProperty(PropertyRequest(service, form, "Name", "Other")));
            object com = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            try
            {
                form.Properties.Add(new Property { Name = "Native", Stored = com });
                Assert.ThrowsException<InvalidOperationException>(() => service.SetProperty(PropertyRequest(service, form, "Native", "new")));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
        }

        [TestMethod]
        public void SnapshotToleratesFontReadFailuresAndControlCreationValidatesItsArguments()
        {
            var form = new VbeFormsTests.FakeForm("Form1"); var project = new VbeFormsTests.FakeProject(); project.VBComponents.Add(form);
            var host = new VbeFormsTests.FakeVbe(); host.VBProjects.Add(project); var service = new VbeForms(host);
            form.Designer.Controls.AddExisting("Label1").FailFontRead = true;
            dynamic state = service.State(project.Name, form.Name); Assert.IsNull((string)state.Controls[0].FontName);
            Assert.ThrowsException<ArgumentException>(() => service.Create(new Request { Form = null }));
            foreach (string type in new[] { null, "Definitely.NotInstalled.1" })
                Assert.ThrowsException<ArgumentException>(() => service.AddControl(new Request { Project = project.Name, Form = form.Name, Control = "Child", ControlType = type, Width = 20, Height = 10, ExpectedFormVersion = state.Version }));
            Assert.ThrowsException<InvalidOperationException>(() => service.AddControl(new Request { Project = project.Name, Form = form.Name, Control = "Label1", ControlType = "Forms.Label.1", Width = 20, Height = 10, ExpectedFormVersion = state.Version }));
            var controls = form.Designer.Controls; controls.FailNextCaption = true;
            controls.BeforeEnumeration = () => { if (controls.RemoveCalls > 0) throw new InvalidOperationException("Inspection unavailable after removal"); };
            Assert.ThrowsException<InvalidOperationException>(() => service.AddControl(new Request { Project = project.Name, Form = form.Name, Control = "Child", Caption = "Child", ControlType = "Forms.Label.1", Width = 20, Height = 10, ExpectedFormVersion = state.Version }));
        }

        [TestMethod]
        public void NodePictureRejectsReadOnlyImageMetadataAndHierarchyCollectionsCanDisappear()
        {
            var node = new Node();
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new LiveProperty("Picture", typeof(System.Drawing.Bitmap), () => null) });
            Form form; var service = NodeService(node, out form);
            Assert.ThrowsException<InvalidOperationException>(() => service.SetNodePicture(NodeRequest(service, form, node, "Picture", "unused")));
            var request = NodeRequest(service, form, node, "Caption", "unused");
            int reads = 0; var metadata = form.Designer.Metadata;
            form.Designer.MetadataReader = () => ++reads >= 2 ? new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[0]) : metadata;
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveControl(request));
            Assert.AreEqual(1, ((object[])form.Designer.Controls).Length);
            var root = new Node(); var child = new Node { Name = "Child" };
            root.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new LiveProperty("Controls", typeof(object[]), () => new object[] { child }) });
            Assert.ThrowsException<InvalidOperationException>(() => Call("ResolveNestedControls", root, "Controls/Child/Controls/Grandchild"));
        }

        [TestMethod]
        public void PageCollectionDescriptorDisappearingAfterInspectionPreventsRemoval()
        {
            var node = new Node { Name = "MultiPage1", ClassName = "MultiPage" };
            var page = new Node { Name = "Page1", Parent = node, ClassName = "Page" };
            var pages = new List<Node> { page };
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new LiveProperty("Pages", typeof(List<Node>), () => pages) });
            Form form; var service = NodeService(node, out form);
            var request = NodeRequest(service, form, node, null, null); request.ControlPath += "/Pages/Page1";
            var metadata = node.Metadata; int reads = 0;
            node.MetadataReader = () => ++reads >= 3 ? new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[0]) : metadata;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.RemovePageOrTab(request)).Message, "has no Pages collection");
            Assert.AreEqual(1, pages.Count);
        }

        [TestMethod]
        public void ScalarConversionsCoverEverySupportedTypeAndRejectUnsafeValues()
        {
            foreach (Type unsupported in new[] { null, typeof(object), typeof(int[]), typeof(Uri) })
                Assert.ThrowsException<InvalidOperationException>(() => Call("ConvertScalar", 1, unsupported));
            Assert.AreEqual(VbeProjectComponentsTests.Choice.Second, Call("ConvertScalar", "Second", typeof(VbeProjectComponentsTests.Choice)));
            Assert.AreEqual(VbeProjectComponentsTests.Choice.First, Call("ConvertScalar", 1, typeof(VbeProjectComponentsTests.Choice)));
            Assert.AreEqual(true, Call("ConvertScalar", "true", typeof(bool)));
            Assert.AreEqual(false, Call("ConvertScalar", false, typeof(bool)));
            Assert.ThrowsException<ArgumentException>(() => Call("ConvertScalar", "invalid", typeof(bool)));
            Assert.AreEqual("value", Call("ConvertScalar", "value", typeof(string)));
            Assert.ThrowsException<ArgumentException>(() => Call("ConvertScalar", 1, typeof(string)));
            foreach (Type type in new[] { typeof(float), typeof(double), typeof(decimal) })
            {
                Assert.AreEqual(Convert.ChangeType(2, type), Call("ConvertScalar", 2, type));
                foreach (double value in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity })
                    Assert.ThrowsException<ArgumentException>(() => Call("ConvertScalar", value, type));
            }
            foreach (object value in new object[] { null, "text", true, (byte)1, (sbyte)-1, (short)2, (ushort)3, 4, 5u, 6L, 7UL, 8f, 9d, 10m })
                Assert.AreEqual(value, Call("NormalizeScalar", value));
            Assert.AreEqual("x", Call("NormalizeScalar", 'x'));
        }

        [TestMethod]
        public void DescriptorConversionAndComparisonRespectColorsVariantsAndNulls()
        {
            Assert.AreEqual(2, Call("ConvertDescriptorValue", 2, typeof(object), 1));
            Assert.AreEqual("text", Call("ConvertDescriptorValue", "text", null, null));
            Assert.AreEqual(System.Drawing.ColorTranslator.FromHtml("#112233"), Call("ConvertDescriptorValue", "#112233", typeof(System.Drawing.Color), null));
            Assert.AreEqual(System.Drawing.ColorTranslator.FromOle(255), Call("ConvertDescriptorValue", 255, typeof(System.Drawing.Color), null));
            Assert.AreEqual(System.Drawing.ColorTranslator.FromOle(255), Call("ConvertDescriptorValue", "255", typeof(System.Drawing.Color), null));
            Assert.IsTrue((bool)Call("SameDescriptorValue", 1, 1));
            Assert.IsFalse((bool)Call("SameDescriptorValue", null, 1));
            Assert.IsFalse((bool)Call("SameDescriptorValue", 1, null));
            Assert.IsTrue((bool)Call("SameDescriptorValue", System.Drawing.Color.Red, System.Drawing.Color.FromArgb(System.Drawing.Color.Red.ToArgb())));
            Assert.IsFalse((bool)Call("SameDescriptorValue", System.Drawing.Color.Red, System.Drawing.Color.Blue));
            Assert.IsTrue((bool)Call("SameDescriptorValue", "TEXT", "text"));
            Assert.IsFalse((bool)Call("SameDescriptorValue", 1, 2));
            Assert.IsFalse((bool)Call("SameDescriptorValue", new object(), 1));
            Assert.IsFalse((bool)Call("SameDescriptorValue", 1, new object()));
            Assert.IsFalse((bool)Call("SameDescriptorValue", System.Drawing.Color.Red, "red"));
        }

        [TestMethod]
        public void FontMemberWritesVerifyEverySetterAndRejectEveryInvalidSize()
        {
            var font = new FaultFont(); var designer = new Node { Font = font };
            foreach (string member in new[] { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough" })
            {
                object value = member == "Name" ? (object)"Calibri" : member == "Size" ? 11d : (object)true;
                Call("SetFormFontMember", designer, member, value);
                font.Ignore = member;
                value = member == "Name" ? (object)"Arial" : member == "Size" ? 15d : (object)false;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => Call("SetFormFontMember", designer, member, value)).Message, "did not retain");
                font.Ignore = null;
            }
            foreach (object name in new object[] { null, " ", 1 }) Assert.ThrowsException<ArgumentException>(() => Call("SetFormFontMember", designer, "Name", name));
            foreach (double size in new[] { double.NaN, double.PositiveInfinity, -1, 0, 201 }) Assert.ThrowsException<ArgumentException>(() => Call("SetFormFontMember", designer, "Size", size));
            Assert.ThrowsException<InvalidOperationException>(() => Call("SetFormFontMember", designer, "Unsupported", true));
        }

        [TestMethod]
        public void ContainerIdentityHandlesFallbackNamesMissingParentsAndHierarchyBounds()
        {
            var form = new Node { Name = null, ClassName = "UserForm" };
            Assert.AreEqual("UserForm:Form1", Call("ContainerIdentity", form, "Form1"));
            Assert.IsNull(Call("ContainerIdentity", null, "Form1"));
            Assert.IsNull(Call("ContainerIdentity", new Node { ClassName = null }, "Form1"));
            Assert.IsNull(Call("ContainerIdentity", new Node { Name = null }, "Form1"));
            Assert.IsNull(Call("ContainerIdentity", new Node { FailParent = true }, "Form1"));
            var loop = new Node(); loop.Parent = loop;
            Assert.IsNull(Call("ContainerIdentity", loop, "Form1"));
            Assert.IsNull(Call("ContainerIdentity", new Node(), "Form1"));
            Assert.IsNull(Call("SafeComName", new object()));
            Assert.IsFalse((bool)Call("SameContainer", new Node(), new Node(), "Form1"));
            Assert.IsTrue((bool)Call("SameContainer", new Node { ClassName = "UserForm" }, new Node { ClassName = "UserForm" }, "Form1"));
        }

        [TestMethod]
        public void NativeIdentityComparisonReleasesPointersAndRejectsInvalidComObjects()
        {
            object left = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            object right = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            object released = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(released);
            try
            {
                Assert.IsFalse((bool)Call("SameComIdentity", null, left));
                Assert.IsFalse((bool)Call("SameComIdentity", left, null));
                Assert.IsFalse((bool)Call("SameComIdentity", left, new object()));
                Assert.IsFalse((bool)Call("SameComIdentity", new object(), left));
                Assert.IsTrue((bool)Call("SameComIdentity", left, left));
                Assert.IsFalse((bool)Call("SameComIdentity", left, right));
                Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => Call("SameComIdentity", released, left));
                Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => Call("SameComIdentity", left, released));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(left); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(right); }
        }

        [TestMethod]
        public void MemberInspectionBoundsItsResultAndHandlesNullTypesUnreadableAndNativeValues()
        {
            object com = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            try
            {
                var node = new Node();
                node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
                {
                    new VbeProjectComponentsTests.MetadataProperty("Null", null),
                    new VbeProjectComponentsTests.MetadataProperty("Native", typeof(object), com),
                    new VbeProjectComponentsTests.MetadataProperty("Unreadable", typeof(string)) { FailRead = true },
                    new VbeProjectComponentsTests.MetadataProperty("Enum", typeof(VbeProjectComponentsTests.Choice), VbeProjectComponentsTests.Choice.First)
                });
                var members = (List<VbePropertyInfo>)Call("DescribeObjectMembers", node);
                Assert.AreEqual(4, members.Count); Assert.AreEqual("object", members[1].Kind); Assert.IsNotNull(members[2].Error);
                Assert.AreEqual(2, members[3].AllowedValues.Length);
                node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(Enumerable.Range(0, 70).Select(i => (System.ComponentModel.PropertyDescriptor)new VbeProjectComponentsTests.MetadataProperty("P" + i, typeof(int), i)).ToArray());
                Assert.AreEqual(64, ((List<VbePropertyInfo>)Call("DescribeObjectMembers", node)).Count);
                using (var image = new System.Drawing.Bitmap(2, 2)) Assert.AreEqual(64, ((string)Call("ImageDigest", image)).Length);
                var disposed = new System.Drawing.Bitmap(1, 1); disposed.Dispose(); Assert.IsNull(Call("ImageDigest", disposed));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
        }

        [TestMethod]
        public void FormPropertyMutationRejectsIndexedReadOnlyUnreadableAndUnknownScalarTypes()
        {
            foreach (int fault in new[] { 0, 1, 2, 3, 4 })
            {
                var form = new Form(); var property = new Property { Name = "Custom", Stored = "old" }; form.Properties.Add(property);
                var descriptor = new VbeProjectComponentsTests.MetadataProperty("Custom", typeof(string), "old");
                if (fault == 0) property.Indices = 1;
                if (fault == 1) descriptor.ReadOnly = true;
                if (fault == 2) property.FailRead = true;
                if (fault == 3) { property.Stored = null; descriptor = new VbeProjectComponentsTests.MetadataProperty("Custom", typeof(object)); }
                if (fault == 4) { property.Stored = null; descriptor = null; }
                form.Designer.Metadata = new System.ComponentModel.PropertyDescriptorCollection(descriptor == null ? new System.ComponentModel.PropertyDescriptor[0] : new System.ComponentModel.PropertyDescriptor[] { descriptor });
                var service = Service(form); var request = PropertyRequest(service, form, "Custom", "new");
                Assert.ThrowsException<InvalidOperationException>(() => service.SetProperty(request));
            }
            var tagForm = new Form(); var tag = new Property { Name = "Tag" }; tagForm.Properties.Add(tag); var tagService = Service(tagForm);
            tagService.SetProperty(PropertyRequest(tagService, tagForm, "Tag", "new tag")); Assert.AreEqual("new tag", tag.Stored);
            var nullTypedForm = new Form(); var nullTyped = new Property { Name = "Custom" }; nullTypedForm.Properties.Add(nullTyped);
            nullTypedForm.Designer.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("Custom", typeof(string)) });
            var nullTypedService = Service(nullTypedForm); nullTypedService.SetProperty(PropertyRequest(nullTypedService, nullTypedForm, "Custom", "text")); Assert.AreEqual("text", nullTyped.Stored);
        }

        [TestMethod]
        public void FormPropertyDescriptionsHandleAllValueKindsAndIndependentGetterErrors()
        {
            object com = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            using (var image = new System.Drawing.Bitmap(2, 2))
            using (var font = new System.Drawing.Font("Arial", 10))
                try
                {
                    var form = new Form();
                    form.Properties.AddRange(new[]
                    {
                        new Property { Name = "Indexed", Stored = "value", Indices = 1 },
                        new Property { Name = "IndicesFailure", FailIndices = true },
                        new Property { Name = "RawFailure", FailRead = true },
                        new Property { Name = "ManagedFailure", Stored = "raw" },
                        new Property { Name = "BothFailure", FailRead = true },
                        new Property { Name = "Image", Stored = image },
                        new Property { Name = "Font", Stored = font },
                        new Property { Name = "Enumerable", Stored = new[] { "a", "b" } },
                        new Property { Name = "Com", Stored = com },
                        new Property { Name = "Color", Stored = System.Drawing.Color.Red },
                        new Property { Name = "Picture" }
                    });
                    form.Designer.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
                    {
                        new VbeProjectComponentsTests.MetadataProperty("ManagedFailure", typeof(string)) { FailRead = true },
                        new VbeProjectComponentsTests.MetadataProperty("BothFailure", typeof(string)) { FailRead = true },
                        new VbeProjectComponentsTests.MetadataProperty("Color", typeof(System.Drawing.Color), System.Drawing.Color.Red)
                    });
                    var properties = (List<VbePropertyInfo>)Service(form).Properties("VBAProject", form.Name);
                    Assert.AreEqual("indexed", properties.Single(p => p.Name == "Indexed").Kind);
                    foreach (string name in new[] { "IndicesFailure", "RawFailure", "ManagedFailure", "BothFailure" }) Assert.IsNotNull(properties.Single(p => p.Name == name).Error);
                    foreach (string name in new[] { "Image", "Font", "Enumerable", "Com" }) Assert.AreEqual("object", properties.Single(p => p.Name == name).Kind);
                    Assert.IsNotNull(properties.Single(p => p.Name == "Image").Digest);
                    Assert.IsNotNull(properties.Single(p => p.Name == "Font").Members);
                    Assert.AreEqual("Red", properties.Single(p => p.Name == "Color").Display);
                    Assert.AreEqual("(empty)", properties.Single(p => p.Name == "Picture").Display);
                    form.Designer.FailPictureRead = true;
                    properties = (List<VbePropertyInfo>)Service(form).Properties("VBAProject", form.Name);
                    Assert.IsNotNull(properties.Single(p => p.Name == "Picture").Error);
                }
                finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
        }

        [TestMethod]
        public void TreePropertyInspectionReadsImagesFontsScalarsAndNullableMetadata()
        {
            using (var image = new System.Drawing.Bitmap(2, 2))
            using (var font = new System.Drawing.Font("Arial", 10))
            {
                var node = new Node { Font = font };
                node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
                {
                    new VbeProjectComponentsTests.MetadataProperty("Null", null),
                    new VbeProjectComponentsTests.MetadataProperty("Image", typeof(System.Drawing.Image), image),
                    new VbeProjectComponentsTests.MetadataProperty("Font", typeof(System.Drawing.Font), font),
                    new VbeProjectComponentsTests.MetadataProperty("Unreadable", typeof(string)) { FailRead = true }
                });
                var properties = (List<VbePropertyInfo>)Call("ReadObjectProperties", node);
                Assert.IsNull(properties[0].Value); Assert.IsNotNull(properties[1].Digest); Assert.IsNotNull(properties[2].Members); Assert.IsNotNull(properties[3].Error);
            }
        }

        [TestMethod]
        public void FormPictureWriteUsesOleReadbackAndRejectsAnIgnoredSetter()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-picture-" + Guid.NewGuid().ToString("N") + ".bmp");
            using (var image = new System.Drawing.Bitmap(2, 2)) image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            var form = new Form(); form.Properties.Add(new Property { Name = "Picture" }); var service = Service(form);
            try
            {
                var request = PropertyRequest(service, form, "Picture", "unused"); request.Path = path;
                dynamic result = service.SetPicture(request); Assert.IsNotNull((string)result.Picture);
                var props = (List<VbePropertyInfo>)service.Properties("VBAProject", form.Name); Assert.IsNotNull(props.Single(p => p.Name == "Picture").Digest);
                object installed = form.Designer.Picture; form.Designer.Picture = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(installed);
                form.Designer.IgnorePicture = true;
                request = PropertyRequest(service, form, "Picture", "unused"); request.Path = path;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.SetPicture(request)).Message, "did not retain");
            }
            finally { System.IO.File.Delete(path); }
        }

        [TestMethod]
        public void EnumAndWritableDescriptorChecksHandleUnfinishedTypesAndReadOnlyMembers()
        {
            Assert.IsNull(Call("EnumChoices", new object[] { null }));
            Assert.IsNull(Call("EnumChoices", typeof(int)));
            var assembly = AppDomain.CurrentDomain.DefineDynamicAssembly(new System.Reflection.AssemblyName("VBAi-Enum-" + Guid.NewGuid().ToString("N")), System.Reflection.Emit.AssemblyBuilderAccess.Run);
            var unfinished = assembly.DefineDynamicModule("Types").DefineEnum("Unfinished", System.Reflection.TypeAttributes.Public, typeof(int));
            Assert.IsNull(Call("EnumChoices", unfinished));
            var node = new Node(); var prop = new VbeProjectComponentsTests.MetadataProperty("Left", typeof(double), 1d) { ReadOnly = true };
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { prop });
            Assert.ThrowsException<InvalidOperationException>(() => Call("RequireWritableControlProperty", node, "Width"));
            Assert.ThrowsException<InvalidOperationException>(() => Call("RequireWritableControlProperty", node, "Left"));
            prop.ReadOnly = false; Call("RequireWritableControlProperty", node, "Left");
        }

        [TestMethod]
        public void PathResolutionRejectsEveryInvalidCollectionOrHierarchyShape()
        {
            var root = new Node(); var child = new Node { Name = "Child" };
            root.Controls = new object[] { child };
            root.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[]
            {
                new VbeProjectComponentsTests.MetadataProperty("Controls", typeof(object[]), root.Controls),
                new VbeProjectComponentsTests.MetadataProperty("Other", typeof(object[]), root.Controls)
            });
            foreach (string path in new[] { "Controls", "Controls/Child/Controls", string.Join("/", Enumerable.Repeat("Controls/Child", 9)) })
                Assert.ThrowsException<ArgumentException>(() => Call("ResolveTreeItem", root, path));
            foreach (string path in new[] { "Missing/Child", "Other/Child", "Controls/Absent" })
                Assert.ThrowsException<InvalidOperationException>(() => Call("ResolveTreeItem", root, path));
            Assert.AreSame(child, Call("ResolveTreeItem", root, "Controls/Child"));
            foreach (string path in new[] { null, " ", "Controls", "Controls/Child/Pages", "Pages/Child", string.Join("/", Enumerable.Repeat("Controls/Child", 7)), "Controls/Child/Tabs/Tab", "Controls/0Child" })
                Assert.ThrowsException<ArgumentException>(() => Call("ResolveNestedControls", root, path));
            Assert.ThrowsException<InvalidOperationException>(() => Call("ResolveNestedControls", root, "Controls/Absent"));
            Assert.ThrowsException<InvalidOperationException>(() => Call("ResolveNestedControls", root, "Controls/Child"));
            root.Controls = new object[] { child, new Node { Name = "child" } };
            root.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("Controls", typeof(object[]), root.Controls) });
            Assert.ThrowsException<InvalidOperationException>(() => Call("ResolveNestedControls", root, "Controls/Child"));
        }

        [TestMethod]
        public void TreeInspectionEnforcesDepthAndNodeLimitsAndAcceptsNullableNestedCollections()
        {
            var node = new Node(); node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("Controls", typeof(object[])) });
            var result = Call("ReadTreeNode", node, "Control", "Form1", "Controls", 0, 0); Assert.IsNotNull(result);
            Assert.ThrowsException<InvalidOperationException>(() => Call("ReadTreeNode", node, "Control", "Form1", "Controls", 0, 512));
            Assert.ThrowsException<InvalidOperationException>(() => Call("ReadTreeNode", node, "Control", "Form1", "Controls", 17, 0));
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[0]);
            Assert.IsNotNull(Call("ReadTreeNode", node, "Page", "Form1", "Controls", 0, 0));
        }

        [TestMethod]
        public void NodePropertyMutationValidatesInputsMetadataAndPostWriteState()
        {
            var node = new Node(); var property = new VbeProjectComponentsTests.MetadataProperty("Caption", typeof(string), "old");
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { property });
            Form form; var service = NodeService(node, out form);
            for (int fault = 0; fault < 4; fault++)
            {
                var request = NodeRequest(service, form, node, "Caption", "new");
                if (fault == 0) request.ControlPath = null;
                if (fault == 1) request.Property = null;
                if (fault == 2) request.Value = null;
                if (fault == 3) request.ExpectedTreeVersion = null;
                Assert.ThrowsException<ArgumentException>(() => service.SetNodeProperty(request));
            }
            foreach (string name in new[] { "Caption.X.Y", ".Caption", "Caption." }) Assert.ThrowsException<ArgumentException>(() => service.SetNodeProperty(NodeRequest(service, form, node, name, "new")));
            Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Missing", "new")));
            var stale = NodeRequest(service, form, node, "Caption", "new"); stale.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(stale));
            stale = NodeRequest(service, form, node, "Caption", "new"); stale.ControlPath = "Controls/Absent";
            Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(stale));
            property.ReadOnly = true; Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Caption", "new")));
            property.ReadOnly = false; property.IgnoreWrite = true;
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Caption", "new"))).Message, "did not retain");
            node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("_Font_Reserved", typeof(object)) });
            Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "_Font_Reserved", "new")));
        }

        [TestMethod]
        public void NodeObjectMembersUseNativeComDescriptorsAndRejectUnsafeOwners()
        {
            object dictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", true));
            try
            {
                var node = new Node { Font = dictionary };
                node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new LiveProperty("Font", typeof(object), () => dictionary), new LiveProperty("Dictionary", typeof(object), () => dictionary), new LiveProperty("Null", typeof(object), () => null) });
                Form form; var service = NodeService(node, out form);
                Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Null.Name", "new")));
                Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Font.Missing", "new")));
                Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Dictionary.Count", 2)));
                service.SetNodeProperty(NodeRequest(service, form, node, "Dictionary.CompareMode", 1));
                Assert.AreEqual(1, (int)((dynamic)dictionary).CompareMode);
                var metadata = new Node();
                metadata.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new VbeProjectComponentsTests.MetadataProperty("CompareMode", typeof(int), 1) { IgnoreWrite = true } });
                var provider = new FixedProvider(metadata); System.ComponentModel.TypeDescriptor.AddProvider(provider, dictionary);
                try { StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.SetNodeProperty(NodeRequest(service, form, node, "Dictionary.CompareMode", 2))).Message, "did not retain"); }
                finally { System.ComponentModel.TypeDescriptor.RemoveProvider(provider, dictionary); }
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dictionary); }
        }

        [TestMethod]
        public void NodePictureWritesVerifyNativeReadbackForBitmapAndIconDescriptors()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-node-picture-" + Guid.NewGuid().ToString("N") + ".bmp");
            using (var image = new System.Drawing.Bitmap(2, 2)) image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp);
            try
            {
                foreach (Type type in new[] { typeof(System.Drawing.Bitmap), typeof(System.Drawing.Icon) })
                {
                    var node = new Node { Name = "Image1", ClassName = "Image" };
                    node.Metadata = new System.ComponentModel.PropertyDescriptorCollection(new System.ComponentModel.PropertyDescriptor[] { new LiveProperty("Picture", type, () => node.Picture, value => node.Picture = value) });
                    Form form; var service = NodeService(node, out form); var request = NodeRequest(service, form, node, "Picture", "unused"); request.Path = path;
                    dynamic result = service.SetNodePicture(request); Assert.AreEqual("Picture", (string)result.Property); Assert.IsNotNull(node.Picture);
                    object installed = node.Picture; node.Picture = null; System.Runtime.InteropServices.Marshal.FinalReleaseComObject(installed);
                    node.IgnorePicture = true; request = NodeRequest(service, form, node, "Picture", "unused"); request.Path = path;
                    Assert.ThrowsException<InvalidOperationException>(() => service.SetNodePicture(request));
                }
            }
            finally { System.IO.File.Delete(path); }
        }

    }

    public sealed partial class VbeFormsCoreBranchTests
    {
        [TestMethod]
        public void PageCreationReportsInspectionFailureAfterSuccessfulNativeRollback()
        {
            var f = Create("MultiPage"); var pages = f.Control.Pages;
            var r = f.Request("Controls/MultiPage1"); r.ParentPath = r.ControlPath; r.NewName = "Page1";
            pages.FailAfterAdd = true;
            pages.BeforeEnumeration = () => { if (pages.RemoveCalls > 0) throw new InvalidOperationException("Inspection unavailable after removal"); };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(r, "Pages")).Message, "rollback could not be verified");
            Assert.AreEqual(0, pages.Count); Assert.AreEqual(1, pages.RemoveCalls);
        }
        [TestMethod]
        public void PageReadbackDetectsUnexpectedNameAndConcurrentRemovalBeforeTheMutation()
        {
            var f = Create("MultiPage"); var r = f.Request("Controls/MultiPage1"); r.ParentPath = r.ControlPath; r.NewName = "Page1";
            f.Control.Pages.AfterAdd = item => item.Name = "Unexpected";
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(r, "Pages")).Message, "rollback could not be verified");
            Assert.AreEqual("Unexpected", f.Control.Pages.Single().Name);
            f = Create("MultiPage"); f.Control.Pages.Add("Page1", "Page");
            var remove = f.Request("Controls/MultiPage1/Pages/Page1"); int reads = 0; var pages = f.Control.Pages;
            pages.BeforeEnumeration = () => { if (++reads >= 2) pages.Clear(); };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemovePageOrTab(remove)).Message, "no longer exists");
            f = Create("Frame"); var children = f.Control.Controls;
            r = f.Request("Controls/Frame1"); r.ParentPath = r.ControlPath; r.Control = "Child"; r.ControlType = "Forms.Label.1"; r.Caption = "Label"; r.Width = 20; r.Height = 10;
            children.FailNextCaption = true;
            children.BeforeEnumeration = () => { if (children.RemoveCount > 0) throw new InvalidOperationException("Inspection unavailable after removal"); };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddNestedControl(r));
        }
        [TestMethod]
        [DoNotParallelize]
        public void MutationsKeepTheIndependentVersionGuardEvenWhenTheHasherCollides()
        {
            using (var scope = new VbeFormsCoverageTests.TreeHashScope())
            {
                var f = Create("Frame");
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemoveControl(f.Request("Controls/Frame1")));
                Assert.AreEqual(0, f.Form.Designer.Controls.Count);
                f = Create("MultiPage"); var add = f.Request("Controls/MultiPage1"); add.ParentPath = add.ControlPath; add.NewName = "Page1";
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(add, "Pages")).Message, "absent after rollback");
                Assert.AreEqual(0, f.Control.Pages.Count);
                f.Control.Pages.Add("Page1", "Page");
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemovePageOrTab(f.Request("Controls/MultiPage1/Pages/Page1")));
                Assert.AreEqual(0, f.Control.Pages.Count);
            }
        }

        [TestMethod]
        public void PageCreationVerifiesRollbackAfterInspectionFailureOrIgnoredNativeAdd()
        {
            for (int fault = 0; fault < 5; fault++)
            {
                var f = Create("MultiPage"); var pages = f.Control.Pages;
                var r = f.Request("Controls/MultiPage1"); r.ParentPath = r.ControlPath; r.NewName = "Page1";
                if (fault == 0) pages.HideAdd = true;
                if (fault == 1) { pages.FailAfterAdd = true; pages.IgnoreRemove = true; }
                if (fault >= 2)
                {
                    pages.FailAfterAdd = true; if (fault == 3) pages.FailRemove = true;
                    int failures = 0;
                    pages.BeforeEnumeration = () => { if (pages.AddCalls > 0 && (fault != 4 || failures++ == 0)) throw new InvalidOperationException("Inspection unavailable"); };
                }
                var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(r, "Pages"));
                if (fault == 0) StringAssert.Contains(error.Message, "absent after rollback");
                else StringAssert.Contains(error.Message, "rollback could not be verified");
            }
        }

        [TestMethod]
        public void HierarchyMutationsRejectEveryMissingVersionStalePathAndInvalidShape()
        {
            var f = Create("MultiPage"); f.Control.Pages.Add("Page1", "Page");
            for (int operation = 0; operation < 5; operation++)
                for (int fault = 0; fault < 4; fault++)
                {
                    var r = f.Request(operation == 3 ? "Controls/MultiPage1/Pages/Page1" : "Controls/MultiPage1"); r.ParentPath = r.ControlPath; r.NewName = "NewPage"; r.Property = "Picture";
                    if (fault == 0) { r.ControlPath = null; r.ParentPath = null; }
                    if (fault == 1) r.ExpectedTreeVersion = null;
                    if (fault == 2) r.ExpectedTreeVersion = "stale";
                    if (fault == 3) { r.ControlPath = operation == 3 ? "Controls/Absent/Pages/Page1" : "Controls/Absent"; r.ParentPath = r.ControlPath; }
                    Action execute = () =>
                    {
                        if (operation == 0) f.Service.RemoveControl(r);
                        if (operation == 1) f.Service.ZOrderControl(r);
                        if (operation == 2) f.Service.AddPageOrTab(r, "Pages");
                        if (operation == 3) f.Service.RemovePageOrTab(r);
                        if (operation == 4) f.Service.SetNodePicture(r);
                    };
                    if (fault <= 1) Assert.ThrowsException<ArgumentException>(execute);
                    else Assert.ThrowsException<InvalidOperationException>(execute);
                }
            foreach (string path in new[] { "Controls", "Controls/MultiPage1/Controls", "Controls/MultiPage1/Pages/Page1" })
            {
                var r = f.Request(path); Assert.ThrowsException<ArgumentException>(() => f.Service.RemoveControl(r));
            }
            foreach (string path in new[] { "Controls", "Controls/MultiPage1/Pages/Page1" }) Assert.ThrowsException<ArgumentException>(() => f.Service.ZOrderControl(f.Request(path)));
            foreach (string path in new[] { "Controls/MultiPage1", "Controls/MultiPage1/Pages", "Controls/MultiPage1/Controls/Child" }) Assert.ThrowsException<ArgumentException>(() => f.Service.RemovePageOrTab(f.Request(path)));
            var picture = f.Request("Controls/MultiPage1"); picture.Property = null;
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetNodePicture(picture));
            var pagesRequest = f.Request("Controls/MultiPage1"); pagesRequest.ParentPath = pagesRequest.ControlPath; pagesRequest.NewName = "Page2";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(pagesRequest, "Missing"));
            pagesRequest.InsertIndex = -1; Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.AddPageOrTab(pagesRequest, "Pages"));
        }

        [TestMethod]
        public void RemovalAndZOrderDetectIgnoredDeletesAndDisappearingControls()
        {
            var f = Create("Frame"); f.Form.Designer.Controls.IgnoreRemove = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemoveControl(f.Request("Controls/Frame1")));
            f.Form.Designer.Controls.IgnoreRemove = false;
            f.Control.AfterZOrder = () => f.Form.Designer.Controls.Remove(f.Control.Name);
            var r = f.Request("Controls/Frame1");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ZOrderControl(r));
            f = Create("MultiPage"); f.Control.Pages.Add("Page1", "Page"); f.Control.Pages.IgnoreRemove = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemovePageOrTab(f.Request("Controls/MultiPage1/Pages/Page1")));
            f = Create("MultiPage"); f.Control.Pages.Add("Page1", "Page"); f.Control.Pages.Add("Page1", "Duplicate");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemovePageOrTab(f.Request("Controls/MultiPage1/Pages/Page1")));
            f = Create("MultiPage"); f.Control.Pages.Add("Other", "Other"); f.Control.Pages.Add("Page1", "Page");
            f.Service.RemovePageOrTab(f.Request("Controls/MultiPage1/Pages/Page1")); Assert.AreEqual("Other", f.Control.Pages.Single().Name);
        }
        [TestMethod]
        public void NestedControlCreationChecksRollbackAtEveryNativeCollectionBoundary()
        {
            for (int fault = 0; fault < 7; fault++)
            {
                var f = Create("Frame"); var controls = f.Control.Controls;
                var r = f.Request("Controls/Frame1"); r.ParentPath = r.ControlPath; r.Control = "Child"; r.ControlType = "Forms.Label.1"; r.Width = 20; r.Height = 10; r.Caption = "Label";
                if (fault == 0) controls.FailAdd = true;
                if (fault >= 1 && fault <= 5) controls.FailNextCaption = true;
                if (fault == 2) controls.FailNextRemove = true;
                if (fault == 3) controls.IgnoreRemove = true;
                if (fault == 4 || fault == 5)
                {
                    int failures = 0;
                    controls.BeforeEnumeration = () => { if (controls.Count > 0 && (fault == 4 || failures++ == 0)) throw new InvalidOperationException("Inspection unavailable"); };
                }
                if (fault == 6) { controls.HideAdded = true; r.Caption = null; }
                var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddNestedControl(r));
                if (fault == 0 || fault == 1 || fault == 6) StringAssert.Contains(error.Message, "no control with the requested name remains");
                else StringAssert.Contains(error.Message, "rollback could not be verified");
            }
        }

        [TestMethod]
        public void NestedCreationRejectsMissingVersionsUnknownParentsAndDuplicateChildren()
        {
            var f = Create("Frame");
            for (int fault = 0; fault < 4; fault++)
            {
                var r = f.Request("Controls/Frame1"); r.ParentPath = r.ControlPath; r.Control = "Child"; r.ControlType = "Forms.Label.1"; r.Width = 20; r.Height = 10;
                if (fault == 0) r.ControlType = null;
                if (fault == 1) r.ExpectedTreeVersion = null;
                if (fault == 2) r.ExpectedTreeVersion = "stale";
                if (fault == 3) r.ParentPath = "Controls/Absent";
                if (fault <= 1) Assert.ThrowsException<ArgumentException>(() => f.Service.AddNestedControl(r));
                else Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddNestedControl(r));
            }
            f.Control.Controls.AddExisting("Label", "Existing");
            var duplicate = f.Request("Controls/Frame1"); duplicate.ParentPath = duplicate.ControlPath; duplicate.Control = "Existing"; duplicate.ControlType = "Forms.Label.1"; duplicate.Width = 20; duplicate.Height = 10;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddNestedControl(duplicate));
        }
    }

    public sealed partial class VbeFormsContractTests
    {
        [TestMethod]
        public void FormListingIncludesOnlyUserFormsAndTheirDesignerState()
        {
            var project = new FakeProject
            {
                Name = "Projet"
            };
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
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsCoreBranchTests
    {
        [TestMethod]
        public void MultiPageAddAndRemoveUseCanonicalPagePathAndIndex()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.Caption = "First";
            dynamic first = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("Controls/MultiPage1/Pages/PageA", (string)first.AddedPath);
            Assert.AreEqual("First", f.Control.Pages.Single().Caption);
            request.ExpectedTreeVersion = first.Tree.TreeVersion;
            request.NewName = "PageB";
            request.InsertIndex = 0;
            dynamic second = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("PageB", f.Control.Pages.First().Name);
            var remove = f.Request("Controls/MultiPage1/Pages/PageA");
            dynamic removed = f.Service.RemovePageOrTab(remove);
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual("PageB", f.Control.Pages.Single().Name);
            Assert.AreNotEqual((string)second.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void TabStripAddAndRemoveTrackNativeTabCollection()
        {
            var f = Create("TabStrip");
            var request = f.Request("Controls/TabStrip1");
            request.ParentPath = request.ControlPath;
            request.NewName = "TabA";
            dynamic added = f.Service.AddPageOrTab(request, "Tabs");
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Name);
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Caption);
            dynamic removed = f.Service.RemovePageOrTab(f.Request("Controls/TabStrip1/Tabs/TabA"));
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(0, f.Control.Tabs.Count);
            Assert.AreNotEqual((string)added.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void PageAddRejectsStaleVersionDuplicateIndexAndWrongParent()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.ExpectedTreeVersion = f.Version;
            request.InsertIndex = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.InsertIndex = null;
            f.Control.Pages.Add("PageA", "Existing");
            request.ExpectedTreeVersion = f.Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            Assert.AreEqual(1, f.Control.Pages.Count);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Tabs"));
        }

        [TestMethod]
        public void PageAddRollsBackAfterNativeAddFailureAndReportsFailedRollback()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            f.Control.Pages.FailAfterAdd = true;
            var clean = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(clean.Message, "absent after rollback");
            Assert.AreEqual(0, f.Control.Pages.Count);
            f.Control.Pages.FailRemove = true;
            var incomplete = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(incomplete.Message, "rollback could not be verified");
            Assert.AreEqual(1, f.Control.Pages.Count);
        }

        [TestMethod]
        public void NodeScalarMutationUsesVersionAndConvertsToDeclaredType()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Left";
            request.Value = "12.5";
            dynamic changed = f.Service.SetNodeProperty(request);
            Assert.AreEqual(12.5d, f.Control.Left);
            request.ExpectedTreeVersion = changed.Tree.TreeVersion;
            request.Property = "Caption";
            request.Value = "Updated";
            dynamic captioned = f.Service.SetNodeProperty(request);
            Assert.AreEqual("Updated", f.Control.Caption);
            Assert.AreNotEqual((string)changed.Tree.TreeVersion, (string)captioned.Tree.TreeVersion);
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
        }

        [TestMethod]
        public void NodeMutationRefusesCrashProneNativeSettersBeforeWriting()
        {
            foreach (var caseInfo in new[]
            {
                Tuple.Create("ComboBox", "ColumnCount", (object)2),
                Tuple.Create("TextBox", "ScrollBars", (object)2),
                Tuple.Create("SpinButton", "Min", (object)1),
                Tuple.Create("ToggleButton", "Value", (object)true),
                Tuple.Create("Label", "Cancel", (object)true)
            }

            )
            {
                var f = Create(caseInfo.Item1);
                var request = f.Request("Controls/" + caseInfo.Item1 + "1");
                request.Property = caseInfo.Item2;
                request.Value = caseInfo.Item3;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request), caseInfo.Item1 + "." + caseInfo.Item2);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
        }

        [TestMethod]
        public void NodeMutationRejectsManagedObjectMemberAndUnknownProperty()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Missing";
            request.Value = "x";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            request.Property = "Font.Name";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            Assert.AreEqual("Arial", f.Control.Font.Name);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using VBAi;

    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListFiltersNonFormsAndTreeTracksCanonicalNestedPaths()
        {
            var fixture = NewFixture();
            fixture.Project.VBComponents.Add(new FakeModule("Module1"));
            var frame = fixture.Form.Designer.Controls.AddExisting("Frame1");
            frame.Controls.AddExisting("Child1");
            var forms = ((IEnumerable)fixture.Service.List(fixture.Project.Name)).Cast<object>().ToArray();
            Assert.AreEqual(1, forms.Length);
            Assert.AreEqual("Form1", (string)((dynamic)forms[0]).Name);
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual(2, (int)tree.NodeCount);
            var root = ((IEnumerable)tree.Controls).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1", (string)((dynamic)root).Path);
            var child = ((IEnumerable)((dynamic)root).Children).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1/Controls/Child1", (string)((dynamic)child).Path);
            Assert.AreEqual((string)tree.FormVersion, (string)tree.TreeVersion);
        }

        [TestMethod]
        public void SnapshotReportsOnlyDirectControlsAndParentProbeConfirmsOwnership()
        {
            var fixture = NewFixture();
            var frame = fixture.Form.Designer.Controls.AddExisting("Frame1");
            frame.Controls.AddExisting("Child1");
            dynamic state = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual("Original caption", (string)state.Caption);
            Assert.AreEqual(1, ((IEnumerable)state.Controls).Cast<object>().Count());
            Assert.AreEqual("Frame1", (string)((dynamic)((IEnumerable)state.Controls).Cast<object>().Single()).Name);
            dynamic parents = fixture.Service.ParentProbe(fixture.Project.Name, fixture.Form.Name);
            var row = ((IEnumerable)parents.Rows).Cast<object>().Single();
            Assert.IsTrue((bool)((dynamic)row).SameDesigner);
            Assert.IsFalse((bool)((dynamic)row).SameComponent);
        }

        [TestMethod]
        public void FormPropertiesDescribeScalarValueAndRejectStaleMutation()
        {
            var fixture = NewFixture();
            var properties = ((IEnumerable)fixture.Service.Properties(fixture.Project.Name, fixture.Form.Name)).Cast<VbePropertyInfo>().ToArray();
            var caption = properties.Single(p => p.Name == "Caption");
            Assert.AreEqual("scalar", caption.Kind);
            Assert.AreEqual("Original caption", caption.Value);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Property = "Caption",
                Value = "Edited",
                ExpectedFormVersion = "stale"
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            Assert.AreEqual("Original caption", fixture.Form.Designer.Caption);
        }

        [TestMethod]
        public void FormCaptionMutationCanBeReversedWithFreshVersion()
        {
            var fixture = NewFixture();
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Property = "Caption",
                Value = "Edited",
                ExpectedFormVersion = before.Version
            };
            dynamic changed = fixture.Service.SetProperty(request);
            Assert.AreEqual("Edited", fixture.Form.Designer.Caption);
            Assert.AreEqual("Edited", (string)changed.State.Caption);
            Assert.AreNotEqual((string)before.Version, (string)changed.State.Version);
            request.Value = "Original caption";
            request.ExpectedFormVersion = changed.State.Version;
            dynamic restored = fixture.Service.SetProperty(request);
            Assert.AreEqual("Original caption", (string)restored.State.Caption);
            Assert.AreEqual((string)before.Version, (string)restored.State.Version);
        }

        [TestMethod]
        public void AddAndRemoveControlRoundTripChangesTreeWithoutTouchingOtherControls()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Keep");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Added",
                ControlType = "Forms.Label.1",
                Caption = "New",
                Left = 2,
                Top = 3,
                Width = 60,
                Height = 20,
                ExpectedFormVersion = before.TreeVersion
            };
            dynamic added = fixture.Service.AddControl(request);
            Assert.AreEqual(2, fixture.Form.Designer.Controls.Count);
            Assert.AreEqual("New", fixture.Form.Designer.Controls.Item("Added").Caption);
            Assert.AreNotEqual((string)before.TreeVersion, (string)added.Version);
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            dynamic removed = fixture.Service.RemoveControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Added", ExpectedTreeVersion = tree.TreeVersion });
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Count);
            Assert.AreEqual("Keep", fixture.Form.Designer.Controls.Single().Name);
            Assert.AreEqual((string)before.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void FailedPostAddSetterRollsBackTheControl()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.FailNextCaption = true;
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, Control = "Broken", ControlType = "Forms.Label.1", Caption = "Fails", Left = 1, Top = 1, Width = 40, Height = 20, ExpectedFormVersion = before.Version }));
            StringAssert.Contains(error.Message, "no control with the requested name remains");
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Count);
            dynamic after = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual((string)before.Version, (string)after.Version);
        }

        [TestMethod]
        public void NestedControlCanBeAddedAndRemovedByCanonicalPath()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Frame1");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            dynamic added = fixture.Service.AddNestedControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ParentPath = "Controls/Frame1", Control = "Nested", ControlType = "Forms.Label.1", Caption = "Inside", Left = 1, Top = 1, Width = 30, Height = 10, ExpectedTreeVersion = before.TreeVersion });
            Assert.AreEqual(2, (int)added.NodeCount);
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Item("Frame1").Controls.Count);
            dynamic removed = fixture.Service.RemoveControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Frame1/Controls/Nested", ExpectedTreeVersion = added.TreeVersion });
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Item("Frame1").Controls.Count);
            Assert.AreEqual((string)before.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void NodeCaptionMutationRequiresFreshTreeAndCanBeReversed()
        {
            var fixture = NewFixture();
            var label = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Label1",
                Property = "Caption",
                Value = "After",
                ExpectedTreeVersion = before.TreeVersion
            };
            dynamic changed = fixture.Service.SetNodeProperty(request);
            Assert.AreEqual("After", label.Caption);
            Assert.AreNotEqual((string)before.TreeVersion, (string)changed.Tree.TreeVersion);
            request.Value = "Original control";
            request.ExpectedTreeVersion = changed.Tree.TreeVersion;
            dynamic restored = fixture.Service.SetNodeProperty(request);
            Assert.AreEqual("Original control", label.Caption);
            Assert.AreEqual((string)before.TreeVersion, (string)restored.Tree.TreeVersion);
        }

        [TestMethod]
        public void GeometryRejectsNonFiniteInputBeforeVbideMutation()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                Left = double.NaN,
                Top = 1,
                Width = 20,
                Height = 10
            };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlGeometry(request));
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Item("Label1").Left);
        }

        [TestMethod]
        public void CreateAndOpenFormRequireValidIdentityAndDesignMode()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = "NewForm"
            };
            fixture.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(request));
            fixture.Project.Mode = 2;
            request.Form = "2Invalid";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.Create(request));
            request.Form = "form1";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(request));
            request.Form = "NewForm";
            dynamic created = fixture.Service.Create(request);
            Assert.AreEqual("NewForm", (string)created.Form);
            var form = fixture.Project.VBComponents.OfType<FakeForm>().Single(item => item.Name == "NewForm");
            Assert.IsTrue(form.DesignerWindow().Visible);
            form.DesignerWindow().Visible = false;
            dynamic opened = fixture.Service.Open(fixture.Project.Name, "newform");
            Assert.IsTrue(form.DesignerWindow().Visible);
            Assert.AreEqual("NewForm", (string)opened.Form);
        }

        [TestMethod]
        public void FailedFormRenameRollsBackOnlyTheCreatedComponent()
        {
            var fixture = NewFixture();
            fixture.Project.VBComponents.RejectedCreatedName = "DeniedForm";
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(new Request { Project = fixture.Project.Name, Form = "DeniedForm" }));
            StringAssert.Contains(error.Message, "absent after rollback");
            Assert.AreEqual(1, fixture.Project.VBComponents.Count);
            Assert.AreEqual("Form1", fixture.Project.VBComponents.OfType<FakeForm>().Single().Name);
        }

        [TestMethod]
        public void ControlGeometryNameCaptionAndFontRoundTripWithFreshVersion()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                ExpectedFormVersion = before.Version,
                Left = 12,
                Top = 7,
                Width = 80,
                Height = 21
            };
            dynamic moved = fixture.Service.SetControlGeometry(request);
            Assert.AreEqual(12d, control.Left);
            Assert.AreEqual(80d, control.Width);
            request.ExpectedFormVersion = moved.Version;
            request.NewName = "Heading";
            dynamic renamed = fixture.Service.RenameControl(request);
            Assert.AreEqual("Heading", control.Name);
            request.Control = "Heading";
            request.ExpectedFormVersion = renamed.Version;
            request.Caption = "New heading";
            dynamic captioned = fixture.Service.SetControlCaption(request);
            Assert.AreEqual("New heading", control.Caption);
            request.ExpectedFormVersion = captioned.Version;
            request.FontName = "Consolas";
            request.FontSize = 13;
            request.FontBold = true;
            dynamic font = fixture.Service.SetControlFont(request);
            Assert.AreEqual("Consolas", control.Font.Name);
            Assert.AreEqual(13d, control.Font.Size);
            Assert.IsTrue(control.Font.Bold);
            Assert.AreNotEqual((string)before.Version, (string)font.Version);
        }

        [TestMethod]
        public void ControlEditsRejectDuplicateNameMissingCaptionInvalidFontAndStaleState()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            fixture.Form.Designer.Controls.AddExisting("Keep");
            dynamic state = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                ExpectedFormVersion = state.Version,
                NewName = "Keep"
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RenameControl(request));
            request.NewName = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.RenameControl(request));
            request.Caption = null;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlCaption(request));
            request.FontName = "Arial";
            request.FontSize = double.NaN;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlFont(request));
            request.FontSize = 11;
            request.ExpectedFormVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetControlFont(request));
            Assert.AreEqual("Label1", fixture.Form.Designer.Controls.Item("Label1").Name);
        }

        [TestMethod]
        public void ControlPropertiesIncludeCurrentNativeValuesAndRejectUnknownControl()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1").Caption = "Visible text";
            var properties = ((IEnumerable)fixture.Service.ControlProperties(fixture.Project.Name, fixture.Form.Name, "Label1")).Cast<object>().ToArray();
            var caption = properties.Single(property => (string)((dynamic)property).Name == "Caption");
            Assert.AreEqual("Visible text", (string)((dynamic)caption).Value);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ControlProperties(fixture.Project.Name, fixture.Form.Name, "Missing"));
        }

        [TestMethod]
        public void FormPropertiesWriteScalarIdentityAndAllSupportedFontMembers()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ExpectedFormVersion = ((dynamic)fixture.Service.State(fixture.Project.Name, fixture.Form.Name)).Version
            };
            request.Property = "Width";
            request.Value = "345.5";
            dynamic resized = fixture.Service.SetProperty(request);
            Assert.AreEqual(345.5d, fixture.Form.Designer.Width);
            request.ExpectedFormVersion = resized.State.Version;
            request.Property = "Name";
            request.Value = "RenamedForm";
            dynamic renamed = fixture.Service.SetProperty(request);
            Assert.AreEqual("RenamedForm", fixture.Form.Name);
            request.Form = fixture.Form.Name;
            request.ExpectedFormVersion = renamed.State.Version;
            foreach (var edit in new[]
            {
                Tuple.Create("Font.Name", (object)"Consolas"),
                Tuple.Create("Font.Size", (object)12.5d),
                Tuple.Create("Font.Bold", (object)true),
                Tuple.Create("Font.Italic", (object)true),
                Tuple.Create("Font.Underline", (object)true),
                Tuple.Create("Font.Strikethrough", (object)true)
            }

            )
            {
                request.Property = edit.Item1;
                request.Value = edit.Item2;
                dynamic result = fixture.Service.SetProperty(request);
                request.ExpectedFormVersion = result.State.Version;
            }

            Assert.AreEqual("Consolas", fixture.Form.Designer.Font.Name);
            Assert.AreEqual(12.5d, fixture.Form.Designer.Font.Size);
            Assert.IsTrue(fixture.Form.Designer.Font.Bold);
            Assert.IsTrue(fixture.Form.Designer.Font.Italic);
            Assert.IsTrue(fixture.Form.Designer.Font.Underline);
            Assert.IsTrue(fixture.Form.Designer.Font.Strikethrough);
        }

        [TestMethod]
        public void FormPropertyWritesRejectUnknownUnsafeAndMalformedMembers()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ExpectedFormVersion = ((dynamic)fixture.Service.State(fixture.Project.Name, fixture.Form.Name)).Version,
                Value = "value"
            };
            foreach (string property in new[]
            {
                "Missing",
                "Caption.Other"
            }

            )
            {
                request.Property = property;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            }

            request.Property = "Font..Size";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            request.Property = "Font.Unknown";
            request.Value = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            request.Property = "Font.Size";
            request.Value = double.PositiveInfinity;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            request.Property = "Name";
            request.Value = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            Assert.AreEqual("Form1", fixture.Form.Name);
        }

        [TestMethod]
        public void EventCatalogUsesExactFormOrCanonicalControlPath()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic formEvents = fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, null);
            Assert.AreEqual("UserForm", (string)formEvents.ObjectName);
            dynamic controlEvents = fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, "Controls/Label1");
            Assert.AreEqual("Label1", (string)controlEvents.ObjectName);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, "Label1"));
        }

        [TestMethod]
        public void ZOrderInvokesNativeControlButReportsUnverifiedEffect()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Label1",
                ExpectedTreeVersion = tree.TreeVersion,
                ZPosition = 0
            };
            dynamic moved = fixture.Service.ZOrderControl(request);
            Assert.AreEqual(1, control.ZOrderCount);
            Assert.AreEqual(0, control.LastZPosition);
            Assert.AreEqual("Unverified", (string)moved.Verification);
            Assert.IsTrue((bool)moved.VerificationPending);
            request.ZPosition = 2;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.ZOrderControl(request));
            request.ZPosition = 1;
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ZOrderControl(request));
            Assert.AreEqual(1, control.ZOrderCount);
        }

        [TestMethod]
        public void PageTabAndPictureCommandsRejectUnverifiedTargetsBeforeNativeWrites()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Frame1");
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var page = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ParentPath = "Controls/Frame1",
                NewName = "Page1",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddPageOrTab(page, "Pages"));
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddPageOrTab(page, "Tabs"));
            var remove = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Frame1",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.RemovePageOrTab(remove));
            var picture = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Frame1",
                Property = "Caption",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetNodePicture(picture));
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Count);
        }
    }
}
