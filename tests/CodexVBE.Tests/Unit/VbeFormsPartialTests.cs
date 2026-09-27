using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsPartialTests
    {
        private static Fixture Create(string type = "Label")
        {
            var form = new FakeForm();
            var control = form.Designer.Controls.AddExisting(type, type + "1");
            var project = new FakeProject();
            project.VBComponents.Add(form);
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture { Project = project, Form = form, Control = control,
                Service = new VbeForms(vbe) };
        }

        [TestMethod]
        public void DuplicateLabelCopiesSupportedPropertiesAndChangesTree()
        {
            var f = Create();
            f.Control.Caption = "Source";
            f.Control.Left = 4; f.Control.Top = 7;
            f.Control.Width = 48; f.Control.Height = 19;
            f.Control.BackColor = 255;
            f.Control.Font.Name = "Calibri";
            f.Control.Font.Size = 11;
            f.Control.Font.Bold = true;
            dynamic before = f.Service.Tree(f.Project.Name, f.Form.Name);
            dynamic result = f.Service.DuplicateLabel(new Request {
                Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/Label1",
                NewName = "LabelCopy", ExpectedTreeVersion = before.TreeVersion });
            var copy = f.Form.Designer.Controls.Item("LabelCopy");
            Assert.AreEqual("Controls/LabelCopy", (string)result.NewPath);
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual(4d, copy.Left);
            Assert.AreEqual(48d, copy.Width);
            Assert.AreEqual(255, copy.BackColor);
            Assert.AreEqual("Calibri", copy.Font.Name);
            Assert.IsTrue(copy.Font.Bold);
            Assert.AreNotEqual((string)before.TreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void DuplicateLabelRollsBackNewControlWhenNativeSetterFails()
        {
            var f = Create();
            f.Form.Designer.Controls.FailNextCaption = true;
            dynamic before = f.Service.Tree(f.Project.Name, f.Form.Name);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(new Request {
                Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/Label1",
                NewName = "BrokenCopy", ExpectedTreeVersion = before.TreeVersion }));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)before.TreeVersion,
                (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void DuplicateLabelRejectsStaleTreeAndNameCollisionBeforeAdd()
        {
            var f = Create();
            var request = new Request { Project = f.Project.Name, Form = f.Form.Name,
                ControlPath = "Controls/Label1", NewName = "LabelCopy", ExpectedTreeVersion = "stale" };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(request));
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            request.ExpectedTreeVersion = tree.TreeVersion;
            request.NewName = "Label1";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void AppendListItemUsesCanonicalTreeAndVerifiesCount()
        {
            var f = Create("ComboBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            dynamic result = f.Service.AppendListItem(new Request {
                Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/ComboBox1",
                ExpectedTreeVersion = tree.TreeVersion, Text = "first" });
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(0, (int)result.CountBefore);
            Assert.AreEqual(1, (int)result.CountAfter);
            Assert.AreEqual("first", f.Control.Items.Single());
        }

        [TestMethod]
        public void AppendListItemRejectsBoundMulticolumnAndStaleListsWithoutMutation()
        {
            var f = Create("ListBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            var request = new Request { Project = f.Project.Name, Form = f.Form.Name,
                ControlPath = "Controls/ListBox1", ExpectedTreeVersion = tree.TreeVersion,
                Text = "value" };
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            request.ExpectedTreeVersion = tree.TreeVersion;
            f.Control.RowSource = "Sheet1!A1:A3";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            f.Control.RowSource = "";
            f.Control.ColumnCount = 2;
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            Assert.AreEqual(0, f.Control.ListCount);
        }

        [TestMethod]
        public void AddListItemRequiresVersionAndRejectsBoundControlBeforeNativeAdd()
        {
            var f = Create("ComboBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            var request = new Request { Project = f.Project.Name, Form = f.Form.Name,
                ControlPath = "Controls/ComboBox1", ExpectedTreeVersion = tree.TreeVersion,
                Text = "value" };
            Assert.ThrowsException<ArgumentException>(() => f.Service.AddListItem(request));
            request.ExpectedListVersion = "current";
            f.Control.RowSource = "Sheet1!A1:A3";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddListItem(request));
            Assert.AreEqual(0, f.Control.ListCount);
        }

        [TestMethod]
        public void ListInitializerRejectsInvalidItemsAndNestedPathsBeforeVbideAccess()
        {
            var service = new VbeForms(new FakeVbe());
            var request = new Request { Project = "VBAProject", Form = "Form1",
                ControlPath = "Controls/ComboBox1", ExpectedTreeVersion = "tree",
                ExpectedSha256 = "sha", Items = new[] { "first\nsecond" } };
            Assert.ThrowsException<ArgumentException>(() => service.SetListInitializer(request));
            request.Items = new[] { "valid" };
            request.ControlPath = "Controls/Frame1/Controls/ComboBox1";
            Assert.ThrowsException<ArgumentException>(() => service.SetListInitializer(request));
        }

        [TestMethod]
        public void ManagedListBlockEscapesQuotesAndRejectsEditedBody()
        {
            const string prefix = "' CodexVBE BEGIN LIST Controls/ComboBox1 SHA256=";
            const string end = "' CodexVBE END LIST Controls/ComboBox1";
            var generate = typeof(VbeForms).GetMethod("GenerateListBlock",
                BindingFlags.NonPublic | BindingFlags.Static);
            var validate = typeof(VbeForms).GetMethod("ValidateManagedBlock",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(generate);
            Assert.IsNotNull(validate);
            var lines = (string[])generate.Invoke(null,
                new object[] { "ComboBox1", new[] { "a\"b", "second" }, prefix, end });
            Assert.AreEqual("    Me.ComboBox1.AddItem \"a\"\"b\"", lines[2]);
            validate.Invoke(null, new object[] { lines, 0, lines.Length - 1, "ComboBox1", prefix });
            lines[2] = "    Me.ComboBox1.AddItem \"changed\"";
            var error = Assert.ThrowsException<TargetInvocationException>(() =>
                validate.Invoke(null, new object[] { lines, 0, lines.Length - 1, "ComboBox1", prefix }));
            Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
        }

        private sealed class Fixture
        {
            public FakeProject Project;
            public FakeForm Form;
            public FakeControl Control;
            public VbeForms Service;
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; } = "VBAProject";
            public int Mode { get; set; } = 2;
            public List<FakeForm> VBComponents { get; } = new List<FakeForm>();
        }

        public sealed class FakeForm
        {
            public string Name { get; set; } = "Form1";
            public int Type => 3;
            public FakeDesigner Designer { get; } = new FakeDesigner();
            public FakeProperties Properties { get; }
            public FakeForm() { Properties = new FakeProperties(Designer); }
        }

        public sealed class FakeDesigner
        {
            public FakeControls Controls { get; }
            public string Caption { get; set; } = "Form";
            public double Width { get; set; } = 300;
            public double Height { get; set; } = 200;
            public FakeDesigner() { Controls = new FakeControls(this); }
        }

        public sealed class FakeProperties : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> items;
            public FakeProperties(FakeDesigner designer)
            {
                items = new List<FakeProperty> { new FakeProperty("Caption", designer.Caption) };
            }
            public IEnumerator<FakeProperty> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeProperty
        {
            public string Name { get; }
            public object Value { get; }
            public int NumIndices => 0;
            public FakeProperty(string name, object value) { Name = name; Value = value; }
        }

        public sealed class FakeControls : IEnumerable<FakeControl>
        {
            private readonly object owner;
            private readonly List<FakeControl> items = new List<FakeControl>();
            public FakeControls(object owner) { this.owner = owner; }
            public int Count => items.Count;
            public int RemoveCount { get; private set; }
            public bool FailNextCaption { get; set; }
            public FakeControl Add(string progId, string name, bool visible)
            {
                var control = AddExisting(progId == "Forms.Label.1" ? "Label" : "Other", name);
                control.FailCaption = FailNextCaption;
                FailNextCaption = false;
                return control;
            }
            public FakeControl AddExisting(string type, string name)
            {
                var control = new FakeControl(type, name, owner);
                items.Add(control);
                return control;
            }
            public FakeControl Item(string name) { return items.Single(x => x.Name == name); }
            public void Remove(string name) { RemoveCount++; items.Remove(Item(name)); }
            public IEnumerator<FakeControl> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeControl
        {
            private string caption = "Original";
            public FakeControl(string type, string name, object parent)
            {
                TypeDescriptor.AddProvider(new NamedProvider(TypeDescriptor.GetProvider(this), type), this);
                Name = name; Parent = parent; Controls = new FakeControls(this);
            }
            public string Name { get; set; }
            public object Parent { get; }
            public FakeControls Controls { get; }
            public FakeFont Font { get; } = new FakeFont();
            public bool FailCaption { get; set; }
            public string Caption
            {
                get { return caption; }
                set { if (FailCaption) throw new InvalidOperationException("Native Caption setter failed"); caption = value; }
            }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; } = 20;
            public double Height { get; set; } = 10;
            public int BackColor { get; set; }
            public string RowSource { get; set; } = "";
            public int ColumnCount { get; set; } = 1;
            public List<string> Items { get; } = new List<string>();
            public int ListCount => Items.Count;
            public void AddItem(string text) { Items.Add(text); }
        }

        public sealed class FakeFont
        {
            public string Name { get; set; } = "Arial";
            public double Size { get; set; } = 10;
            public bool Bold { get; set; }
        }

        private sealed class NamedProvider : TypeDescriptionProvider
        {
            private readonly TypeDescriptionProvider parent;
            private readonly string name;
            public NamedProvider(TypeDescriptionProvider parent, string name) { this.parent = parent; this.name = name; }
            public override ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object instance)
            {
                return new NamedDescriptor(parent.GetTypeDescriptor(objectType, instance), name);
            }
        }

        private sealed class NamedDescriptor : CustomTypeDescriptor
        {
            private readonly string name;
            public NamedDescriptor(ICustomTypeDescriptor parent, string name) : base(parent) { this.name = name; }
            public override string GetClassName() { return name; }
        }
    }
}
