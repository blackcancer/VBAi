namespace VBAi.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Collections.Generic;
    using VBAi;

    public sealed partial class VbeFormsCoverageTests
    {
        public sealed class Node : CustomTypeDescriptor
        {
            public string ClassName { get; set; } = "Frame";
            public string Name { get; set; } = "Node";
            private object parent;
            public bool FailParent { get; set; }
            public object Parent { get { if (FailParent) throw new InvalidOperationException("Parent unavailable"); return parent; } set { parent = value; } }
            public object Controls { get; set; } = new object[0];
            public Action<string, object> ScalarWrite { get; set; }
            public float Width { set { ScalarWrite?.Invoke("Width", value); } }
            public float Height { set { ScalarWrite?.Invoke("Height", value); } }
            public float ScrollWidth { set { ScalarWrite?.Invoke("ScrollWidth", value); } }
            public float ScrollHeight { set { ScalarWrite?.Invoke("ScrollHeight", value); } }
            private object font;
            public bool FailFont { get; set; }
            public object Font { get { if (FailFont) throw new InvalidOperationException("Font unavailable"); return font; } set { font = value; } }
            private object picture;
            public bool IgnorePicture { get; set; }
            public bool FailPictureRead { get; set; }
            public object Picture { get { if (FailPictureRead) throw new InvalidOperationException("Picture unavailable"); return picture; } set { if (!IgnorePicture) picture = value; } }
            public PropertyDescriptorCollection Metadata { get; set; } = new PropertyDescriptorCollection(new PropertyDescriptor[0]);
            public Func<PropertyDescriptorCollection> MetadataReader { get; set; }
            public override string GetClassName() => ClassName;
            public override PropertyDescriptorCollection GetProperties() => MetadataReader == null ? Metadata : MetadataReader();
            public override PropertyDescriptorCollection GetProperties(Attribute[] attributes) => GetProperties();
        }

        public sealed class Property
        {
            public string Name { get; set; }
            public object Stored { get; set; }
            public int Indices { get; set; }
            public bool FailIndices { get; set; }
            public bool FailRead { get; set; }
            public int NumIndices { get { if (FailIndices) throw new InvalidOperationException("Indices unavailable"); return Indices; } }
            public object Value { get { if (FailRead) throw new InvalidOperationException("Value unavailable"); return Stored; } set { Stored = value; } }
        }

        public sealed class Properties : List<Property>
        {
            public Property Item(string name) => Find(p => p.Name == name);
        }

        public sealed class Form
        {
            public string Name { get; set; } = "Form1";
            public int Type => 3;
            public Node Designer { get; } = new Node { ClassName = "UserForm" };
            public Properties Properties { get; } = new Properties
            {
                new Property { Name = "Caption", Stored = "Form" },
                new Property { Name = "Width", Stored = 300d },
                new Property { Name = "Height", Stored = 200d }
            };
        }

        private static VbeForms Service(Form form, object additional = null)
        {
            var project = new VbeFormsTests.FakeProject(); project.VBComponents.Add(form);
            if (additional != null) project.VBComponents.Add(additional);
            var host = new VbeFormsTests.FakeVbe(); host.VBProjects.Add(project);
            return new VbeForms(host);
        }

        private static Request PropertyRequest(VbeForms service, Form form, string name, object value)
        {
            return new Request { Project = "VBAProject", Form = form.Name, Property = name, Value = value, ExpectedFormVersion = (string)((dynamic)service.State("VBAProject", form.Name)).Version };
        }

        public sealed class FaultFont
        {
            public string Ignore { get; set; }
            private string name = "Arial";
            private double size = 10;
            private bool bold, italic, underline, strikethrough;
            public string Name { get { return name; } set { if (Ignore != "Name") name = value; } }
            public double Size { get { return size; } set { if (Ignore != "Size") size = value; } }
            public bool Bold { get { return bold; } set { if (Ignore != "Bold") bold = value; } }
            public bool Italic { get { return italic; } set { if (Ignore != "Italic") italic = value; } }
            public bool Underline { get { return underline; } set { if (Ignore != "Underline") underline = value; } }
            public bool Strikethrough { get { return strikethrough; } set { if (Ignore != "Strikethrough") strikethrough = value; } }
        }

        public sealed class LiveProperty : PropertyDescriptor
        {
            private readonly Type type;
            private readonly Func<object> read;
            private readonly Action<object> write;
            public LiveProperty(string name, Type type, Func<object> read, Action<object> write = null) : base(name, null) { this.type = type; this.read = read; this.write = write; }
            public override Type ComponentType => typeof(Node);
            public override Type PropertyType => type;
            public override bool IsReadOnly => write == null;
            public override bool CanResetValue(object component) => false;
            public override bool ShouldSerializeValue(object component) => false;
            public override void ResetValue(object component) { }
            public override object GetValue(object component) => read();
            public override void SetValue(object component, object value) => write(value);
        }

        public sealed class FixedProvider : TypeDescriptionProvider
        {
            private readonly ICustomTypeDescriptor descriptor;
            public FixedProvider(ICustomTypeDescriptor descriptor) { this.descriptor = descriptor; }
            public override ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object instance) => descriptor;
        }

        internal sealed class TreeHashScope : IDisposable
        {
            private readonly Func<byte[], byte[]> previous = VbeForms.HashTree;
            public TreeHashScope() { VbeForms.HashTree = _ => new byte[32]; }
            public void Dispose() { VbeForms.HashTree = previous; }
        }

        private static VbeForms NodeService(Node node, out Form form)
        {
            form = new Form(); node.Parent = form.Designer;
            form.Designer.Controls = new object[] { node };
            var controls = form.Designer.Controls;
            form.Designer.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] { new LiveProperty("Controls", typeof(object[]), () => controls) });
            return Service(form);
        }

        private static Request NodeRequest(VbeForms service, Form form, Node node, string property, object value)
        {
            return new Request { Project = "VBAProject", Form = form.Name, ControlPath = "Controls/" + node.Name, Property = property, Value = value, ExpectedTreeVersion = (string)((dynamic)service.Tree("VBAProject", form.Name)).TreeVersion };
        }

        private static object Call(string method, params object[] arguments)
        {
            try { return typeof(VbeForms).GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, arguments); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
    }
}
