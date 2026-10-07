namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;
    using VBAi;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Source;
            public VbeForms Service;
            public Request Request(string path = null, string name = "Copy")
            {
                return new Request
                {
                    Project = Project.Name,
                    Form = Form.Name,
                    ControlPath = path ?? "Controls/" + Source.Name,
                    NewName = name,
                    ExpectedTreeVersion = ((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion
                };
            }
        }

        private static Fixture Create(string type)
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var source = form.Designer.Controls.AddExisting(type, type + "1");
            source.Caption = "Source";
            source.Left = 3;
            source.Top = 4;
            source.Width = 50;
            source.Height = 16;
            return new Fixture
            {
                Project = project,
                Form = form,
                Source = source,
                Service = new VbeForms(vbe)
            };
        }

        private static object Duplicate(Fixture f, string type, Request request)
        {
            try { return typeof(VbeForms).GetMethod("Duplicate" + type).Invoke(f.Service, new object[] { request }); }
            catch (System.Reflection.TargetInvocationException error)
            { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }

        private static void VerifyDuplicationGuards(string type)
        {
            {
                var f = Create(type);
                var provider = new VanishingControlsProvider(f.Form.Designer);
                System.ComponentModel.TypeDescriptor.AddProvider(provider, f.Form.Designer);
                try
                {
                    var request = f.Request(); provider.RemainingReads = 2;
                    var error = Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, request));
                    StringAssert.Contains(error.Message, "parent has no Controls");
                }
                finally { System.ComponentModel.TypeDescriptor.RemoveProvider(provider, f.Form.Designer); }
            }
            foreach (string path in new[] { "", "Controls", "Controls/x/Controls", "Pages/x" })
            {
                var f = Create(type); var request = f.Request(); request.ControlPath = path;
                Assert.ThrowsException<ArgumentException>(() => Duplicate(f, type, request), path);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
            foreach (string name in new[] { "", "bad-name" })
            {
                var f = Create(type); var request = f.Request(name: name);
                Assert.ThrowsException<ArgumentException>(() => Duplicate(f, type, request));
            }
            {
                var f = Create(type); var request = f.Request(); request.ExpectedTreeVersion = "";
                Assert.ThrowsException<ArgumentException>(() => Duplicate(f, type, request));
                request = f.Request(); request.ExpectedTreeVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, request));
                request = f.Request("Controls/Missing");
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, request));
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request(name: f.Source.Name)));
            }
            {
                var f = Create(type == "Label" ? "TextBox" : "Label");
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
            }
            foreach (string property in new[] { "Left", "Top", "Width", "Height" })
            {
                foreach (double value in new[] { double.NaN, double.PositiveInfinity, 0d, -1d, 32768d })
                {
                    if ((property == "Left" || property == "Top") &&
                        (value == 0 || (type == "Label" && !double.IsNaN(value) && !double.IsInfinity(value)))) continue;
                    if (type == "Label" && value == 32768) continue;
                    var f = Create(type); f.Source.ReadOverrides[property] = value;
                    Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()), property + ":" + value);
                    Assert.AreEqual(1, f.Form.Designer.Controls.Count);
                }
            }
            {
                var f = Create(type); f.Form.Designer.Controls.FailAdd = true;
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
                Assert.AreEqual(0, f.Form.Designer.Controls.RemoveCount);
            }
            {
                var f = Create(type); f.Form.Designer.Controls.HideAdded = true;
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
                Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            }
            foreach (string property in type == "ComboBox" ? new[] { "ListWidth" } :
                type == "TextBox" ? new[] { "Value" } :
                type == "CheckBox" || type == "ToggleButton" ? new[] { "Caption", "Value", "BooleanValue" } : new[] { "Caption" })
            {
                var f = Create(type);
                f.Form.Designer.Controls.ConfigureAdded = copy =>
                    copy.ReadOverrides[property == "BooleanValue" ? "Value" : property] =
                        property == "BooleanValue" ? (object)true : "mismatch";
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()), property);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
            if (type == "ToggleButton")
            {
                var f = Create(type); f.Source.Value = null;
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
            }
            {
                var f = Create(type);
                using (new VbeFormsCoverageTests.TreeHashScope())
                {
                    Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
                    Assert.AreEqual(1, f.Form.Designer.Controls.Count);
                    Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
                }
            }
            foreach (string property in new[] { "Left", "Top", "Width", "Height" })
            {
                var f = Create(type);
                f.Form.Designer.Controls.ConfigureAdded = copy => copy.ReadOverrides[property] = 123d;
                Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()), property);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
                Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            }
            {
                var f = Create(type);
                f.Form.Designer.Controls.ConfigureAdded = copy => copy.ReadOverrides["Left"] = 123d;
                f.Form.Designer.Controls.FailNextRemove = true;
                var error = Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, type, f.Request()));
                StringAssert.Contains(error.Message, "rollback also failed");
                Assert.AreEqual(2, f.Form.Designer.Controls.Count);
            }
            {
                var f = Create(type);
                var frame = f.Form.Designer.Controls.AddExisting("Frame", "Frame1");
                var source = frame.Controls.AddExisting(type, "Nested");
                source.Caption = "Nested"; source.Width = 20; source.Height = 10;
                var request = f.Request("Controls/Frame1/Controls/Nested");
                dynamic result = Duplicate(f, type, request);
                Assert.AreEqual("Controls/Frame1/Controls/Copy", (string)result.NewPath);
                Assert.AreEqual(2, frame.Controls.Count);
            }
        }

        private sealed class VanishingControlsProvider : System.ComponentModel.TypeDescriptionProvider
        {
            private readonly System.ComponentModel.TypeDescriptionProvider parent;
            internal int RemainingReads = int.MaxValue;
            internal VanishingControlsProvider(object owner) { parent = System.ComponentModel.TypeDescriptor.GetProvider(owner); }
            public override System.ComponentModel.ICustomTypeDescriptor GetTypeDescriptor(Type type, object instance)
            { return new VanishingControlsDescriptor(parent.GetTypeDescriptor(type, instance), () => RemainingReads-- > 0); }
        }

        private sealed class VanishingControlsDescriptor : System.ComponentModel.CustomTypeDescriptor
        {
            private readonly Func<bool> visible;
            internal VanishingControlsDescriptor(System.ComponentModel.ICustomTypeDescriptor parent, Func<bool> visible) : base(parent)
            { this.visible = visible; }
            public override System.ComponentModel.PropertyDescriptorCollection GetProperties()
            {
                var properties = base.GetProperties();
                return visible() ? properties : new System.ComponentModel.PropertyDescriptorCollection(
                    properties.Cast<System.ComponentModel.PropertyDescriptor>().Where(p => p.Name != "Controls").ToArray());
            }
        }
    }
}
