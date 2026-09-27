using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeProcedureMutationTests
    {
        [TestMethod]
        public void CreateReplaceListAndRemoveProcedureRoundTrip()
        {
            var fixture = new Fixture("Option Explicit");
            var create = fixture.Request("Sub Run()\nDebug.Print 1\nEnd Sub");
            dynamic made = fixture.Navigation.CreateProcedure(create);
            Assert.AreEqual(3, (int)made.BodyLine);
            Assert.AreEqual("Run", (string)made.Procedure);
            Assert.IsFalse((bool)made.CompilationVerified);

            dynamic listed = fixture.Navigation.Procedures("Projet", "Module1");
            Assert.AreEqual(1, (int)listed.Procedures.Count);
            Assert.AreEqual("Run", (string)listed.Procedures[0].Name);
            create.ExpectedSha256 = (string)made.Sha256;
            create.Text = "Sub Run()\nDebug.Print 2\nEnd Sub";
            dynamic replaced = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsTrue((bool)replaced.Changed);
            StringAssert.Contains(fixture.Module.Code, "Debug.Print 2");
            create.ExpectedSha256 = (string)replaced.Sha256;
            dynamic unchanged = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsFalse((bool)unchanged.Changed);

            dynamic removed = fixture.Navigation.RemoveProcedure(create);
            Assert.AreEqual(3, (int)removed.RemovedLineCount);
            Assert.IsFalse(fixture.Module.Code.Contains("Run()"));
            Assert.AreEqual(0, (int)((dynamic)fixture.Navigation.Procedures("Projet", "Module1")).Procedures.Count);
        }

        [TestMethod]
        public void StaleVersionAndDuplicateProcedureNeverChangeCode()
        {
            var fixture = new Fixture("Option Explicit");
            var request = fixture.Request("Sub Run()\nEnd Sub");
            request.ExpectedSha256 = Hash("stale");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual("Option Explicit", fixture.Module.Code);
            request.ExpectedSha256 = Hash(fixture.Module.Code);
            fixture.Navigation.CreateProcedure(request);
            var before = fixture.Module.Code;
            request.ExpectedSha256 = Hash(before);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedReplacementRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nDebug.Print 1\r\nEnd Sub");
            var request = fixture.Request("Sub Run()\nDebug.Print 2\nEnd Sub");
            var before = fixture.Module.Code;
            fixture.Module.FailNextInsert = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.ReplaceProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedRemovalRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nEnd Sub");
            var request = fixture.Request(null);
            var before = fixture.Module.Code;
            fixture.Module.PretendStillPresentAfterDelete = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.RemoveProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void EventStubUsesCurrentFormTreeAndCurrentCodeVersion()
        {
            var fixture = new Fixture("", 3);
            fixture.Component.Designer.Controls.AddExisting("Button1");
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request { Project = "Projet", Form = "Module1", ObjectName = "Button1",
                EventName = "Click", ExpectedTreeVersion = tree.TreeVersion, ExpectedSha256 = Hash("") };
            dynamic result = fixture.Navigation.CreateEventProcedure(request);
            Assert.AreEqual("Button1_Click", (string)result.Procedure);
            StringAssert.Contains(fixture.Module.Code, "Private Sub Button1_Click()");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
        }

        [TestMethod]
        public void EventStubRejectsStaleTreeAndUnknownControlBeforeEditingCode()
        {
            var fixture = new Fixture("", 3);
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request { Project = "Projet", Form = "Module1", ObjectName = "Missing",
                EventName = "Click", ExpectedTreeVersion = tree.TreeVersion, ExpectedSha256 = Hash("") };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
            request.ObjectName = "UserForm";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
        }

        private static string Hash(string source)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(source)))
                    .Replace("-", "").ToLowerInvariant();
        }

        private sealed class Fixture
        {
            public readonly FakeModule Module;
            public readonly VbeCodeNavigation Navigation;
            public readonly VbeForms Forms;
            public readonly FakeComponent Component;
            public Fixture(string code, int type = 1)
            {
                var host = new FakeVbe();
                var project = new FakeProject { Name = "Projet", Mode = 2 };
                Component = new FakeComponent { Name = "Module1", Type = type };
                Module = new FakeModule(Component, code);
                Component.CodeModule = Module;
                project.VBComponents.Add(Component);
                host.VBProjects.Add(project);
                Forms = new VbeForms(host);
                Navigation = new VbeCodeNavigation(host, Forms);
            }
            public Request Request(string text)
            {
                return new Request { Project = "Projet", Module = "Module1", Procedure = "Run", ProcKind = 0,
                    Text = text, ExpectedSha256 = Hash(Module.Code) };
            }
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public int Mode { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }
        public sealed class FakeComponent
        {
            private readonly VbeFormsTests.FakeWindow window = new VbeFormsTests.FakeWindow();
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
            public VbeFormsTests.FakeDesigner Designer { get; } = new VbeFormsTests.FakeDesigner();
            public VbeFormsTests.FakePropertyCollection Properties => new VbeFormsTests.FakePropertyCollection(Designer);
            public bool HasOpenDesigner => window.Visible;
            public VbeFormsTests.FakeWindow DesignerWindow() { return window; }
        }

        public sealed class FakeModule
        {
            private readonly List<string> lines;
            public FakeComponent Parent { get; }
            public FakeLines Lines { get; }
            public FakeProcedureIndex ProcBodyLine { get; }
            public FakeProcedureIndex ProcStartLine { get; }
            public FakeProcedureIndex ProcCountLines { get; }
            public dynamic ProcOfLine { get; }
            public bool FailNextInsert { get; set; }
            public bool PretendStillPresentAfterDelete { get; set; }
            public int CountOfLines => lines.Count;
            public int CountOfDeclarationLines => lines.Count > 0 && lines[0].StartsWith("Option ", StringComparison.Ordinal) ? 1 : 0;
            public string Code => string.Join("\r\n", lines);

            public FakeModule(FakeComponent parent, string code)
            {
                Parent = parent;
                lines = string.IsNullOrEmpty(code) ? new List<string>() :
                    code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                Lines = new FakeLines(this);
                ProcBodyLine = new FakeProcedureIndex(this, 0);
                ProcStartLine = new FakeProcedureIndex(this, 0);
                ProcCountLines = new FakeProcedureIndex(this, 1);
                ProcOfLine = new FakeProcOfLine(this);
            }

            public void InsertLines(int start, string source)
            {
                if (FailNextInsert) { FailNextInsert = false; throw new InvalidOperationException("Insert failed"); }
                lines.InsertRange(start - 1, source.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }
            public void DeleteLines(int start, int count) { lines.RemoveRange(start - 1, count); }

            public int CreateEventProc(string eventName, string objectName)
            {
                string stub = "Private Sub " + objectName + "_" + eventName + "()\r\nEnd Sub";
                int first = lines.Count + (lines.Count == 0 ? 1 : 2);
                InsertLines(lines.Count + 1, (lines.Count == 0 ? "" : "\r\n") + stub);
                return first;
            }

            private Procedure Find(string name, int kind)
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    var match = Regex.Match(lines[i], @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Let|Set|Get))\s+([A-Za-z][A-Za-z0-9_]*)\s*\(", RegexOptions.IgnoreCase);
                    if (!match.Success || !string.Equals(match.Groups[2].Value, name, StringComparison.OrdinalIgnoreCase)) continue;
                    int actualKind = match.Groups[1].Value.Equals("Property Let", StringComparison.OrdinalIgnoreCase) ? 1 :
                        match.Groups[1].Value.Equals("Property Set", StringComparison.OrdinalIgnoreCase) ? 2 :
                        match.Groups[1].Value.Equals("Property Get", StringComparison.OrdinalIgnoreCase) ? 3 : 0;
                    if (actualKind != kind) continue;
                    string endKind = actualKind == 0 ? match.Groups[1].Value : "Property";
                    int end = i;
                    while (end < lines.Count && !Regex.IsMatch(lines[end], @"^\s*End\s+" + endKind + @"\s*$", RegexOptions.IgnoreCase)) end++;
                    return new Procedure { Name = name, Kind = kind, Start = i + 1,
                        Count = Math.Min(end, lines.Count - 1) - i + 1 };
                }
                return null;
            }

            private Procedure AtLine(int line)
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    var declaration = Regex.Match(lines[i], @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function)\s+([A-Za-z][A-Za-z0-9_]*)\s*\(", RegexOptions.IgnoreCase);
                    if (!declaration.Success) continue;
                    var found = Find(declaration.Groups[1].Value, 0);
                    if (found != null && line >= found.Start && line < found.Start + found.Count) return found;
                }
                return null;
            }

            private sealed class Procedure
            {
                public string Name;
                public int Kind;
                public int Start;
                public int Count;
            }
            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module) { this.module = module; }
                public string this[int start, int count] => string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
            }
            public sealed class FakeProcedureIndex
            {
                private readonly FakeModule module;
                private readonly int selector;
                public FakeProcedureIndex(FakeModule module, int selector) { this.module = module; this.selector = selector; }
                public int this[string name, int kind]
                {
                    get
                    {
                        if (module.PretendStillPresentAfterDelete && selector == 0 && module.Find(name, kind) == null)
                        { module.PretendStillPresentAfterDelete = false; return 2; }
                        var procedure = module.Find(name, kind);
                        if (procedure == null) throw new InvalidOperationException("Procedure not found");
                        return selector == 1 ? procedure.Count : procedure.Start;
                    }
                }
            }
            public sealed class FakeProcOfLine : DynamicObject
            {
                private readonly FakeModule module;
                public FakeProcOfLine(FakeModule module) { this.module = module; }
                public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
                {
                    var procedure = module.AtLine(Convert.ToInt32(indexes[0]));
                    result = procedure?.Name;
                    return true;
                }
            }
        }
    }
}
