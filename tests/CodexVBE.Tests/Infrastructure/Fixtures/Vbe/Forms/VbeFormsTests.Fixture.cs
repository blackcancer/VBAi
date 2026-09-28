namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeFormsTests
    {
        private static Fixture NewFixture()
        {
            var project = new FakeProject();
            var form = new FakeForm("Form1");
            project.VBComponents.Add(form);
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture
            {
                Project = project,
                Form = form,
                Service = new VbeForms(vbe)
            };
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
            public FakeComponents VBComponents { get; } = new FakeComponents();
        }

        public sealed class FakeComponents : IEnumerable<object>
        {
            private readonly List<object> items = new List<object>();
            public string RejectedCreatedName { get; set; }
            public Action BeforeAdd { get; set; }
            public Action BeforeRemove { get; set; }
            public Action BeforeEnumeration { get; set; }
            public bool IgnoreRemove { get; set; }
            public int AddCalls { get; private set; }
            public int RemoveCalls { get; private set; }
            public Action<FakeForm> ConfigureAdded { get; set; }
            public int Count => items.Count;

            public void Add(object component)
            {
                items.Add(component);
            }

            public FakeForm Add(int type)
            {
                AddCalls++;
                BeforeAdd?.Invoke();
                var form = new FakeForm("Temporary")
                {
                    RejectedName = RejectedCreatedName
                };
                items.Add(form);
                ConfigureAdded?.Invoke(form);
                return form;
            }

            public void Remove(object component)
            {
                RemoveCalls++;
                BeforeRemove?.Invoke();
                if (!IgnoreRemove) items.Remove(component);
            }

            public IEnumerator<object> GetEnumerator()
            {
                BeforeEnumeration?.Invoke();
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeModule
        {
            public FakeModule(string name)
            {
                Name = name;
            }

            public string Name { get; set; }

            public int Type
            {
                get
                {
                    return 1;
                }
            }
        }

        public sealed class FakeForm
        {
            private readonly FakeWindow window = new FakeWindow();
            private string name;
            public FakeForm(string name)
            {
                Name = name;
                Designer = new FakeDesigner();
                Properties = new FakePropertyCollection(Designer);
            }

            public string RejectedName { get; set; }
            public bool FailDesignerWindow { get; set; }

            public string Name
            {
                get
                {
                    return name;
                }

                set
                {
                    if (value == RejectedName)
                        throw new InvalidOperationException("VBE rejected form name");
                    name = value;
                }
            }

            public int Type
            {
                get
                {
                    return 3;
                }
            }

            public bool HasOpenDesigner
            {
                get
                {
                    return window.Visible;
                }
            }

            public FakeDesigner Designer { get; }
            public FakePropertyCollection Properties { get; }

            public FakeWindow DesignerWindow()
            {
                if (FailDesignerWindow) throw new InvalidOperationException("Designer window unavailable");
                return window;
            }
        }

        public sealed class FakeWindow
        {
            public bool Visible { get; set; }
        }

        public sealed class FakeDesigner
        {
            public FakeDesigner()
            {
                Controls = new FakeControls(this);
            }

            public string Caption { get; set; } = "Original caption";
            public double Width { get; set; } = 300;
            public double Height { get; set; } = 200;
            public FakeFont Font { get; } = new FakeFont();
            public FakeControls Controls { get; }
        }

        public sealed class FakePropertyCollection : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> properties;
            public FakePropertyCollection(FakeDesigner designer)
            {
                properties = new List<FakeProperty>
                {
                    new FakeProperty("Caption", () => designer.Caption, value => designer.Caption = (string)value),
                    new FakeProperty("Width", () => designer.Width, value => designer.Width = (double)value),
                    new FakeProperty("Height", () => designer.Height, value => designer.Height = (double)value),
                    new FakeProperty("Font", () => designer.Font, value => throw new InvalidOperationException("Use a Font member"))
                };
            }

            public FakeProperty Item(string name)
            {
                return properties.Single(p => p.Name == name);
            }

            public IEnumerator<FakeProperty> GetEnumerator()
            {
                return properties.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeProperty
        {
            private readonly Func<object> read;
            private readonly Action<object> write;
            public FakeProperty(string name, Func<object> read, Action<object> write)
            {
                Name = name;
                this.read = read;
                this.write = write;
            }

            public string Name { get; }

            public int NumIndices
            {
                get
                {
                    return 0;
                }
            }

            public object Value
            {
                get
                {
                    return read();
                }

                set
                {
                    write(value);
                }
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

            public bool FailNextCaption { get; set; }
            public bool FailAdd { get; set; }
            public bool HideAdd { get; set; }
            public bool IgnoreRemove { get; set; }
            public Action BeforeRemove { get; set; }
            public Action BeforeEnumeration { get; set; }
            public int AddCalls { get; private set; }
            public int RemoveCalls { get; private set; }

            public int Count
            {
                get
                {
                    return items.Count;
                }
            }

            public FakeControl Add(string type, string name, bool visible)
            {
                AddCalls++;
                if (FailAdd) throw new InvalidOperationException("Add unavailable");
                var control = new FakeControl(name, owner)
                {
                    FailCaption = FailNextCaption
                };
                FailNextCaption = false;
                if (!HideAdd) items.Add(control);
                return control;
            }

            public FakeControl AddExisting(string name)
            {
                var control = new FakeControl(name, owner);
                items.Add(control);
                return control;
            }

            public FakeControl Item(string name)
            {
                return items.Single(x => x.Name == name);
            }

            public void Remove(string name)
            {
                RemoveCalls++;
                BeforeRemove?.Invoke();
                if (!IgnoreRemove) items.Remove(Item(name));
            }

            public IEnumerator<FakeControl> GetEnumerator()
            {
                BeforeEnumeration?.Invoke();
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
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
            private readonly FakeFont font = new FakeFont();
            public bool FailFontRead { get; set; }
            public FakeFont Font { get { if (FailFontRead) throw new InvalidOperationException("Font unavailable"); return font; } }
            public List<object[]> ListRows { get; } = new List<object[]>();

            public int ListCount
            {
                get
                {
                    if (ThrowListCount) throw new InvalidOperationException("ListCount unavailable"); return ListCountOverride ?? ListRows.Count;
                }
            }

            public int ColumnCount { get; set; } = 1;
            public string RowSource { get; set; }
            public bool ThrowOnCell { get; set; }
            public bool ThrowListCount { get; set; }
            public int? ListCountOverride { get; set; }
            public bool ThrowAdd { get; set; }
            public bool SkipAdd { get; set; }
            public Action<FakeControl> AfterAdd { get; set; }
            public Action<FakeControl> AfterRemove { get; set; }
            public int AddAttempts { get; private set; }
            public void AddItem(string text)
            {
                AddAttempts++;
                if (ThrowAdd) throw new InvalidOperationException("Native AddItem failed");
                if (!SkipAdd) ListRows.Add(new object[] { text });
                AfterAdd?.Invoke(this);
            }
            internal bool ThrowRemove { get; set; }
            internal bool SkipRemove { get; set; }
            internal int RemoveAttempts { get; private set; }

            [IndexerName("List")]
            public object this[int row, int column]
            {
                get
                {
                    if (ThrowOnCell)
                        throw new InvalidOperationException("Indexed cell unavailable");
                    return ListRows[row][column];
                }
            }

            public void RemoveItem(int index)
            {
                RemoveAttempts++;
                if (ThrowRemove)
                    throw new InvalidOperationException("Native RemoveItem failed");
                if (!SkipRemove)
                    ListRows.RemoveAt(index);
                AfterRemove?.Invoke(this);
            }

            public bool FailCaption { get; set; }

            public string Caption
            {
                get
                {
                    return caption;
                }

                set
                {
                    if (FailCaption)
                        throw new InvalidOperationException("Native setter rejected Caption");
                    caption = value;
                }
            }

            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; } = 20;
            public double Height { get; set; } = 10;
            public int ZOrderCount { get; private set; }
            public int LastZPosition { get; private set; }

            public void ZOrder(int position)
            {
                ZOrderCount++;
                LastZPosition = position;
            }
        }

        public sealed class FakeFont
        {
            public string Name { get; set; } = "Arial";
            public double Size { get; set; } = 10;
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public bool Underline { get; set; }
            public bool Strikethrough { get; set; }
        }

        private sealed class NamedControlProvider : TypeDescriptionProvider
        {
            private readonly string name;
            public NamedControlProvider(string name) : base(TypeDescriptor.GetProvider(typeof(FakeControl)))
            {
                this.name = name;
            }

            public override ICustomTypeDescriptor GetTypeDescriptor(Type objectType, object instance)
            {
                return new NamedControlDescriptor(base.GetTypeDescriptor(objectType, instance), name);
            }
        }

        private sealed class NamedControlDescriptor : CustomTypeDescriptor
        {
            private readonly string name;
            public NamedControlDescriptor(ICustomTypeDescriptor parent, string name) : base(parent)
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
