using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsInitializerTests
    {
        private static Fixture Create(string code = "Private Sub UserForm_Initialize()\r\n    Debug.Print \"keep\"\r\nEnd Sub")
        {
            var form = new FakeForm(code);
            var list = form.Designer.Controls.Add("ComboBox", "ComboBox1");
            var project = new FakeProject();
            project.VBComponents.Add(form);
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture { Project = project, Form = form, List = list,
                Service = new VbeForms(vbe) };
        }

        private static Request RequestFor(Fixture f, params string[] items)
        {
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            return new Request { Project = f.Project.Name, Form = f.Form.Name,
                ControlPath = "Controls/ComboBox1", ExpectedTreeVersion = tree.TreeVersion,
                ExpectedSha256 = Sha(f.Form.CodeModule.Code), Items = items };
        }

        private static string Sha(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code)))
                    .Replace("-", "").ToLowerInvariant();
        }

        [TestMethod]
        public void InitializerWritesManagedBlockInsideExistingEventAndPreservesUserCode()
        {
            var f = Create();
            var request = RequestFor(f, "first", "a\"b");
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.IsTrue((bool)result.UserCodePreserved);
            Assert.IsTrue((bool)result.RuntimeVerificationPending);
            StringAssert.Contains(f.Form.CodeModule.Code, "    Debug.Print \"keep\"");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"a\"\"b\"");
            Assert.IsTrue(f.Form.CodeModule.Code.IndexOf("CodexVBE BEGIN LIST", StringComparison.Ordinal) <
                f.Form.CodeModule.Code.IndexOf("End Sub", StringComparison.Ordinal));
            Assert.AreEqual(Sha(f.Form.CodeModule.Code), (string)result.Sha256After);
            Assert.AreEqual(0, f.List.AddCount);
        }

        [TestMethod]
        public void InitializerCreatesMissingEventWithoutTouchingDesignerList()
        {
            var f = Create("");
            var request = RequestFor(f, "one");
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(1, f.Form.CodeModule.CreateEventCount);
            StringAssert.Contains(f.Form.CodeModule.Code, "Private Sub UserForm_Initialize()");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"one\"");
            Assert.AreEqual(0, f.List.AddCount);
        }

        [TestMethod]
        public void InitializerIsIdempotentForCurrentIdenticalBlock()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "one"));
            string before = f.Form.CodeModule.Code;
            dynamic again = f.Service.SetListInitializer(RequestFor(f, "one"));
            Assert.IsFalse((bool)again.Applied);
            Assert.IsTrue((bool)again.Verified);
            Assert.AreEqual(before, f.Form.CodeModule.Code);
            Assert.AreEqual(Sha(before), (string)again.Sha256After);
        }

        [TestMethod]
        public void InitializerReplacesOnlyPreviouslyManagedBlock()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "old"));
            dynamic changed = f.Service.SetListInitializer(RequestFor(f, "new"));
            Assert.IsTrue((bool)changed.Applied);
            Assert.IsTrue((bool)changed.Verified);
            Assert.IsTrue((bool)changed.UserCodePreserved);
            StringAssert.Contains(f.Form.CodeModule.Code, "Debug.Print \"keep\"");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"new\"");
            Assert.IsFalse(f.Form.CodeModule.Code.Contains("Me.ComboBox1.AddItem \"old\""));
            Assert.AreEqual(1, f.Form.CodeModule.Code.Split(new[] { "CodexVBE BEGIN LIST" },
                StringSplitOptions.None).Length - 1);
        }

        [TestMethod]
        public void InitializerRefusesEditedManagedBlockWithoutChangingCode()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "old"));
            f.Form.CodeModule.ReplaceText("Me.ComboBox1.AddItem \"old\"",
                "Me.ComboBox1.AddItem \"user edit\"");
            string edited = f.Form.CodeModule.Code;
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.SetListInitializer(RequestFor(f, "new")));
            Assert.AreEqual(edited, f.Form.CodeModule.Code);
        }

        [TestMethod]
        public void InitializerReportsNativeInsertFailureAsPendingWithOriginalCodeIntact()
        {
            var f = Create();
            string before = f.Form.CodeModule.Code;
            f.Form.CodeModule.FailInsert = true;
            dynamic result = f.Service.SetListInitializer(RequestFor(f, "one"));
            Assert.IsNull((bool?)result.Applied);
            Assert.IsFalse((bool)result.Verified);
            Assert.IsTrue((bool)result.VerificationPending);
            StringAssert.Contains((string)result.NativeError, "insert refused");
            Assert.AreEqual(before, f.Form.CodeModule.Code);
        }

        [TestMethod]
        public void InitializerChecksTreeCodeAndBindingBeforeMutatingModule()
        {
            var f = Create();
            var request = RequestFor(f, "one");
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            request.ExpectedTreeVersion = RequestFor(f, "one").ExpectedTreeVersion;
            request.ExpectedSha256 = Sha("other code");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code);
            f.List.RowSource = "Sheet1!A1:A3";
            request.ExpectedTreeVersion = RequestFor(f, "one").ExpectedTreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            Assert.AreEqual(0, f.Form.CodeModule.InsertCount);
        }

        private sealed class Fixture
        {
            public FakeProject Project;
            public FakeForm Form;
            public FakeControl List;
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
            public FakeForm(string code)
            {
                CodeModule = new FakeCodeModule(code);
                Properties = new[] { new FakeProperty { Name = "Caption", Value = "Form", NumIndices = 0 } };
            }
            public string Name { get; set; } = "Form1";
            public int Type => 3;
            public FakeDesigner Designer { get; } = new FakeDesigner();
            public FakeCodeModule CodeModule { get; }
            public FakeProperty[] Properties { get; }
        }

        public sealed class FakeProperty
        {
            public string Name { get; set; }
            public object Value { get; set; }
            public int NumIndices { get; set; }
        }

        public sealed class FakeDesigner
        {
            public FakeDesigner() { Controls = new FakeControls(this); }
            public FakeControls Controls { get; }
            public string Caption { get; set; } = "Form";
        }

        public sealed class FakeControls : IEnumerable<FakeControl>
        {
            private readonly object owner;
            private readonly List<FakeControl> items = new List<FakeControl>();
            public FakeControls(object owner) { this.owner = owner; }
            public FakeControl Add(string type, string name)
            {
                var item = new FakeControl(type, name, owner);
                items.Add(item);
                return item;
            }
            public IEnumerator<FakeControl> GetEnumerator() { return items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeControl
        {
            public FakeControl(string type, string name, object parent)
            {
                TypeDescriptor.AddProvider(new NamedProvider(TypeDescriptor.GetProvider(this), type), this);
                Name = name; Parent = parent;
            }
            public string Name { get; }
            public object Parent { get; }
            public string RowSource { get; set; } = "";
            public int ColumnCount { get; set; } = 1;
            public int AddCount { get; private set; }
            public void AddItem(string item) { AddCount++; }
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

        public sealed class FakeCodeModule
        {
            private readonly List<string> lines;
            public FakeCodeModule(string code)
            {
                lines = code.Length == 0 ? new List<string>() :
                    code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                Lines = new LineAccessor(this);
                ProcBodyLine = new ProcedureAccessor(this, "body");
                ProcStartLine = new ProcedureAccessor(this, "start");
                ProcCountLines = new ProcedureAccessor(this, "count");
            }
            public string Code => string.Join("\r\n", lines);
            public int CountOfLines => lines.Count;
            public int CreateEventCount { get; private set; }
            public int InsertCount { get; private set; }
            public bool FailInsert { get; set; }
            public LineAccessor Lines { get; }
            public ProcedureAccessor ProcBodyLine { get; }
            public ProcedureAccessor ProcStartLine { get; }
            public ProcedureAccessor ProcCountLines { get; }
            public int CreateEventProc(string eventName, string objectName)
            {
                if (eventName != "Initialize" || objectName != "UserForm")
                    throw new InvalidOperationException("Unexpected event");
                CreateEventCount++;
                lines.Add("Private Sub UserForm_Initialize()");
                lines.Add("End Sub");
                return lines.Count - 1;
            }
            public void InsertLines(int start, string text)
            {
                InsertCount++;
                if (FailInsert) throw new InvalidOperationException("native insert refused");
                lines.InsertRange(start - 1, text.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }
            public void DeleteLines(int start, int count) { lines.RemoveRange(start - 1, count); }
            public void ReplaceText(string oldText, string newText)
            {
                for (int i = 0; i < lines.Count; i++) lines[i] = lines[i].Replace(oldText, newText);
            }
            public sealed class LineAccessor
            {
                private readonly FakeCodeModule module;
                public LineAccessor(FakeCodeModule module) { this.module = module; }
                public string this[int start, int count] =>
                    string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
            }
            public sealed class ProcedureAccessor
            {
                private readonly FakeCodeModule module;
                private readonly string kind;
                public ProcedureAccessor(FakeCodeModule module, string kind)
                { this.module = module; this.kind = kind; }
                public int this[string procedure, int procKind]
                {
                    get
                    {
                        int start = module.lines.FindIndex(line =>
                            line.IndexOf("Sub " + procedure + "(", StringComparison.OrdinalIgnoreCase) >= 0);
                        if (start < 0) throw new COMException("Procedure missing");
                        if (kind == "count") return module.lines.Count - start;
                        return start + 1;
                    }
                }
            }
        }
    }
}
