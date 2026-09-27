using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsTests
    {
        private static Fixture NewFixture()
        {
            var project = new FakeProject();
            var form = new FakeForm("Form1");
            project.VBComponents.Add(form);
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture { Project = project, Form = form, Service = new VbeForms(vbe) };
        }

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
            var properties = ((IEnumerable)fixture.Service.Properties(fixture.Project.Name, fixture.Form.Name))
                .Cast<VbePropertyInfo>().ToArray();
            var caption = properties.Single(p => p.Name == "Caption");
            Assert.AreEqual("scalar", caption.Kind);
            Assert.AreEqual("Original caption", caption.Value);

            var request = new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                Property = "Caption", Value = "Edited", ExpectedFormVersion = "stale" };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            Assert.AreEqual("Original caption", fixture.Form.Designer.Caption);
        }

        [TestMethod]
        public void FormCaptionMutationCanBeReversedWithFreshVersion()
        {
            var fixture = NewFixture();
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                Property = "Caption", Value = "Edited", ExpectedFormVersion = before.Version };
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
            var request = new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                Control = "Added", ControlType = "Forms.Label.1", Caption = "New",
                Left = 2, Top = 3, Width = 60, Height = 20, ExpectedFormVersion = before.TreeVersion };
            dynamic added = fixture.Service.AddControl(request);
            Assert.AreEqual(2, fixture.Form.Designer.Controls.Count);
            Assert.AreEqual("New", fixture.Form.Designer.Controls.Item("Added").Caption);
            Assert.AreNotEqual((string)before.TreeVersion, (string)added.Version);

            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            dynamic removed = fixture.Service.RemoveControl(new Request {
                Project = fixture.Project.Name, Form = fixture.Form.Name,
                ControlPath = "Controls/Added", ExpectedTreeVersion = tree.TreeVersion
            });
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
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddControl(
                new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                    Control = "Broken", ControlType = "Forms.Label.1", Caption = "Fails",
                    Left = 1, Top = 1, Width = 40, Height = 20, ExpectedFormVersion = before.Version }));
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
            dynamic added = fixture.Service.AddNestedControl(new Request {
                Project = fixture.Project.Name, Form = fixture.Form.Name,
                ParentPath = "Controls/Frame1", Control = "Nested",
                ControlType = "Forms.Label.1", Caption = "Inside",
                Left = 1, Top = 1, Width = 30, Height = 10,
                ExpectedTreeVersion = before.TreeVersion
            });
            Assert.AreEqual(2, (int)added.NodeCount);
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Item("Frame1").Controls.Count);
            dynamic removed = fixture.Service.RemoveControl(new Request {
                Project = fixture.Project.Name, Form = fixture.Form.Name,
                ControlPath = "Controls/Frame1/Controls/Nested",
                ExpectedTreeVersion = added.TreeVersion
            });
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
            var request = new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                ControlPath = "Controls/Label1", Property = "Caption", Value = "After",
                ExpectedTreeVersion = before.TreeVersion };
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
            var request = new Request { Project = fixture.Project.Name, Form = fixture.Form.Name,
                Control = "Label1", Left = double.NaN, Top = 1, Width = 20, Height = 10 };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlGeometry(request));
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Item("Label1").Left);
        }

        private sealed class Fixture
        {
            internal FakeProject Project { get; set; }
            internal FakeForm Form { get; set; }
            internal VbeForms Service { get; set; }
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; } = "VBAProject";
            public int Mode { get; set; } = 2;
            public List<object> VBComponents { get; } = new List<object>();
        }

        public sealed class FakeModule
        {
            public FakeModule(string name) { Name = name; }
            public string Name { get; set; }
            public int Type { get { return 1; } }
        }

        public sealed class FakeForm
        {
            private readonly FakeWindow window = new FakeWindow();
            public FakeForm(string name)
            {
                Name = name;
                Designer = new FakeDesigner();
                Properties = new FakePropertyCollection(Designer);
            }
            public string Name { get; set; }
            public int Type { get { return 3; } }
            public bool HasOpenDesigner { get { return window.Visible; } }
            public FakeDesigner Designer { get; }
            public FakePropertyCollection Properties { get; }
            public FakeWindow DesignerWindow() { return window; }
        }

        public sealed class FakeWindow
        {
            public bool Visible { get; set; }
        }

        public sealed class FakeDesigner
        {
            public FakeDesigner() { Controls = new FakeControls(this); }
            public string Caption { get; set; } = "Original caption";
            public double Width { get; set; } = 300;
            public double Height { get; set; } = 200;
            public FakeControls Controls { get; }
        }

        public sealed class FakePropertyCollection : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> properties;
            public FakePropertyCollection(FakeDesigner designer)
            {
                properties = new List<FakeProperty> {
                    new FakeProperty("Caption", () => designer.Caption, value => designer.Caption = (string)value),
                    new FakeProperty("Width", () => designer.Width, value => designer.Width = (double)value),
                    new FakeProperty("Height", () => designer.Height, value => designer.Height = (double)value)
                };
            }
            public FakeProperty Item(string name) { return properties.Single(p => p.Name == name); }
            public IEnumerator<FakeProperty> GetEnumerator() { return properties.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeProperty
        {
            private readonly Func<object> read;
            private readonly Action<object> write;
            public FakeProperty(string name, Func<object> read, Action<object> write)
            { Name = name; this.read = read; this.write = write; }
            public string Name { get; }
            public int NumIndices { get { return 0; } }
            public object Value { get { return read(); } set { write(value); } }
        }

        public sealed class FakeControls : IEnumerable<FakeControl>
        {
            private readonly object owner;
            private readonly List<FakeControl> items = new List<FakeControl>();
            public FakeControls(object owner) { this.owner = owner; }
            public bool FailNextCaption { get; set; }
            public int Count { get { return items.Count; } }
            public FakeControl Add(string type, string name, bool visible)
            {
                var control = new FakeControl(name, owner) { FailCaption = FailNextCaption };
                FailNextCaption = false;
                items.Add(control);
                return control;
            }
            public FakeControl AddExisting(string name)
            {
                var control = new FakeControl(name, owner);
                items.Add(control);
                return control;
            }
            public FakeControl Item(string name) { return items.Single(x => x.Name == name); }
            public void Remove(string name) { items.Remove(Item(name)); }
            public IEnumerator<FakeControl> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeControl
        {
            private string caption = "Original control";
            public FakeControl(string name, object parent)
            {
                Name = name;
                Parent = parent;
                Controls = new FakeControls(this);
            }
            public string Name { get; set; }
            public object Parent { get; }
            public FakeControls Controls { get; }
            public FakeFont Font { get; } = new FakeFont();
            public bool FailCaption { get; set; }
            public string Caption
            {
                get { return caption; }
                set { if (FailCaption) throw new InvalidOperationException("Native setter rejected Caption"); caption = value; }
            }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; } = 20;
            public double Height { get; set; } = 10;
        }

        public sealed class FakeFont
        {
            public string Name { get; set; } = "Arial";
            public double Size { get; set; } = 10;
            public bool Bold { get; set; }
        }
    }
}
