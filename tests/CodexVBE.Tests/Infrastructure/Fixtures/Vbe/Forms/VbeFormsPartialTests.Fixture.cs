namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsPartialTests
    {
        private static Fixture Create(string type = "Label")
        {
            var form = new FakeForm();
            var control = form.Designer.Controls.AddExisting(type, type + "1");
            var project = new FakeProject();
            project.VBComponents.Add(form);
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture
            {
                Project = project,
                Form = form,
                Control = control,
                Service = new VbeForms(vbe)
            };
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

            public FakeForm()
            {
                Properties = new FakeProperties(Designer);
            }
        }

        public sealed class FakeDesigner
        {
            public FakeControls Controls { get; }
            public string Caption { get; set; } = "Form";
            public double Width { get; set; } = 300;
            public double Height { get; set; } = 200;

            public FakeDesigner()
            {
                Controls = new FakeControls(this);
            }
        }

        public sealed class FakeProperties : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> items;
            public FakeProperties(FakeDesigner designer)
            {
                items = new List<FakeProperty>
                {
                    new FakeProperty("Caption", designer.Caption)
                };
            }

            public IEnumerator<FakeProperty> GetEnumerator()
            {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeProperty
        {
            public string Name { get; }
            public object Value { get; }
            public int NumIndices => 0;

            public FakeProperty(string name, object value)
            {
                Name = name;
                Value = value;
            }
        }

        public sealed class FakeControls : IEnumerable<FakeControl>
        {
            private readonly object owner;
            private readonly List<FakeControl> items = new List<FakeControl>();
            public FakeControls(object owner)
            {
                this.owner = owner;
            }

            public int Count => items.Count;
            public int RemoveCount { get; private set; }
            public bool FailNextCaption { get; set; }
            public bool FailNextValue { get; set; }
            public bool FailNextListWidth { get; set; }
            public bool FailNextChildCaption { get; set; }
            public bool FailNextRemove { get; set; }

            public FakeControl Add(string progId, string name, bool visible)
            {
                string type = progId.StartsWith("Forms.", StringComparison.Ordinal) && progId.EndsWith(".1", StringComparison.Ordinal) ? progId.Substring(6, progId.Length - 8) : "Other";
                var control = AddExisting(type, name);
                control.FailCaption = FailNextCaption;
                FailNextCaption = false;
                control.FailValue = FailNextValue;
                FailNextValue = false;
                control.FailListWidth = FailNextListWidth;
                FailNextListWidth = false;
                control.Controls.FailNextCaption = FailNextChildCaption;
                FailNextChildCaption = false;
                return control;
            }

            public FakeControl AddExisting(string type, string name)
            {
                var control = new FakeControl(type, name, owner);
                items.Add(control);
                return control;
            }

            public FakeControl Item(string name)
            {
                return items.Single(x => x.Name == name);
            }

            public void Remove(string name)
            {
                RemoveCount++;
                if (FailNextRemove)
                {
                    FailNextRemove = false;
                    throw new InvalidOperationException("Native Remove failed");
                }

                items.Remove(Item(name));
            }

            public IEnumerator<FakeControl> GetEnumerator()
            {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeControl
        {
            private string caption = "Original";
            private object value;
            private object listWidth = "60 pt";
            public FakeControl(string type, string name, object parent)
            {
                TypeDescriptor.AddProvider(new NamedProvider(TypeDescriptor.GetProvider(this), type), this);
                Name = name;
                Parent = parent;
                Controls = new FakeControls(this);
                Pages = new FakePageTabCollection(this);
                Tabs = new FakePageTabCollection(this);
                if (type == "CheckBox" || type == "ToggleButton" || type == "OptionButton")
                    value = false;
            }

            public string Name { get; set; }
            public object Parent { get; }
            public FakeControls Controls { get; }
            public FakePageTabCollection Pages { get; }
            public FakePageTabCollection Tabs { get; }
            public FakeFont Font { get; } = new FakeFont();
            public bool FailCaption { get; set; }
            public bool FailValue { get; set; }
            public bool FailListWidth { get; set; }

            public string Caption
            {
                get
                {
                    return caption;
                }

                set
                {
                    if (FailCaption)
                        throw new InvalidOperationException("Native Caption setter failed");
                    caption = value;
                }
            }

            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; } = 20;
            public double Height { get; set; } = 10;
            public int BackColor { get; set; }

            public object Value
            {
                get
                {
                    return value;
                }

                set
                {
                    if (FailValue)
                        throw new InvalidOperationException("Native Value setter failed");
                    this.value = value;
                }
            }

            public object ListWidth
            {
                get
                {
                    return listWidth;
                }

                set
                {
                    if (FailListWidth)
                        throw new InvalidOperationException("Native ListWidth setter failed");
                    listWidth = value;
                }
            }

            public string RowSource { get; set; } = "";
            public int ColumnCount { get; set; } = 1;
            public int ScrollBars { get; set; }
            public int Min { get; set; }
            public bool Cancel { get; set; }
            public List<string> Items { get; } = new List<string>();
            public int ListCount => Items.Count;

            public void AddItem(string text)
            {
                Items.Add(text);
            }
        }

        public sealed class FakePageTabCollection : IEnumerable<FakePageTab>
        {
            private readonly FakeControl owner;
            private readonly List<FakePageTab> items = new List<FakePageTab>();
            public bool FailAfterAdd { get; set; }
            public bool FailRemove { get; set; }
            public int Count => items.Count;

            public FakePageTabCollection(FakeControl owner)
            {
                this.owner = owner;
            }

            public FakePageTab Add(string name, string caption)
            {
                return Add(name, caption, items.Count);
            }

            public FakePageTab Add(string name, string caption, int index)
            {
                var item = new FakePageTab(name, caption, owner);
                items.Insert(index, item);
                if (FailAfterAdd)
                    throw new InvalidOperationException("Native collection failed after Add");
                return item;
            }

            public void Remove(int index)
            {
                if (FailRemove)
                    throw new InvalidOperationException("Native collection Remove failed");
                items.RemoveAt(index);
            }

            public IEnumerator<FakePageTab> GetEnumerator()
            {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakePageTab
        {
            public FakePageTab(string name, string caption, FakeControl parent)
            {
                Name = name;
                Caption = caption;
                Parent = parent;
                Controls = new FakeControls(this);
            }

            public string Name { get; set; }
            public string Caption { get; set; }
            public FakeControl Parent { get; }
            public FakeControls Controls { get; }
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
            public NamedProvider(TypeDescriptionProvider parent, string name)
            {
                this.parent = parent;
                this.name = name;
            }

            public override ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object instance)
            {
                return new NamedDescriptor(parent.GetTypeDescriptor(objectType, instance), name);
            }
        }

        private sealed class NamedDescriptor : CustomTypeDescriptor
        {
            private readonly string name;
            public NamedDescriptor(ICustomTypeDescriptor parent, string name) : base(parent)
            {
                this.name = name;
            }

            public override string GetClassName()
            {
                return name;
            }
        }
    }
}
