using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeDebugLocalScalarsTests
    {
        private const string Source = "Sub Inspect()\r\n" +
            "    Dim first As Long, skipped As Variant, second As String, third As Boolean\r\n" +
            "End Sub";

        [STATestMethod]
        public void APageEvaluatesOnlyEligibleIdentifiersAndReportsIncompleteCoverage()
        {
            var fixture = new ScalarFixture(Source);
            var expressions = new List<string>();
            fixture.Service.LocalScalarEvaluator = request =>
            {
                expressions.Add(request.Expression);
                return Task.FromResult<object>(new ScalarReadback
                {
                    Expression = request.Expression,
                    Context = fixture.Context,
                    Value = "value:" + request.Expression
                });
            };
            dynamic first = fixture.Service.InspectLocalScalarsAsync(fixture.Request(0, 2)).GetAwaiter().GetResult();
            Assert.AreEqual("DeclaredScalarCandidatesOnly", (string)first.Coverage);
            Assert.IsFalse((bool)first.RuntimeInventoryComplete);
            Assert.AreEqual(4, (int)first.TotalCandidates);
            Assert.AreEqual(3, (int)first.EligibleCandidates);
            Assert.AreEqual(2, (int)first.NextOffset);
            Assert.AreEqual("Skipped", (string)first.Items[1].Status);
            Assert.AreEqual("NonScalarOrVariantType", (string)first.Items[1].SkipReason);
            CollectionAssert.AreEqual(new[] { "first" }, expressions.ToArray());
            dynamic second = fixture.Service.InspectLocalScalarsAsync(fixture.Request(2, 2)).GetAwaiter().GetResult();
            Assert.IsNull(second.NextOffset);
            CollectionAssert.AreEqual(new[] { "first", "second", "third" }, expressions.ToArray());
            Assert.IsFalse((bool)second.Aborted);
            Assert.IsFalse(VbeDebugInspection.IsActive);
        }

        [STATestMethod]
        public void AnyChangedModeSourceOrLocalsContextStopsBeforeTheNextEvaluation()
        {
            foreach (string mutation in new[] { "mode", "source", "context", "readback" })
            {
                var fixture = new ScalarFixture(Source);
                var expressions = new List<string>();
                fixture.Service.LocalScalarEvaluator = request =>
                {
                    expressions.Add(request.Expression);
                    if (mutation == "mode") fixture.Project.Mode = 2;
                    if (mutation == "source") fixture.Module.Code += "\r\n' changed";
                    if (mutation == "context") fixture.CurrentContext = "VBAProject.Module1.Other";
                    return Task.FromResult<object>(new ScalarReadback
                    {
                        Expression = mutation == "readback" ? "other" : request.Expression,
                        Context = fixture.Context,
                        Value = "1"
                    });
                };
                dynamic result = fixture.Service.InspectLocalScalarsAsync(fixture.Request(0, 4)).GetAwaiter().GetResult();
                Assert.IsTrue((bool)result.Aborted, mutation);
                Assert.AreEqual(1, result.Items.Count, mutation);
                Assert.AreEqual("Error", (string)result.Items[0].Status, mutation);
                Assert.IsNull(result.NextOffset);
                CollectionAssert.AreEqual(new[] { "first" }, expressions.ToArray(), mutation);
                Assert.IsFalse(VbeDebugInspection.IsActive, mutation);
            }
        }

        [STATestMethod]
        public void WrongInitialContextAndStaleHashRejectWithoutEvaluating()
        {
            foreach (string fault in new[] { "context", "sha", "mode" })
            {
                var fixture = new ScalarFixture(Source);
                int calls = 0;
                fixture.Service.LocalScalarEvaluator = evaluation => { calls++; return Task.FromResult<object>(null); };
                var request = fixture.Request(0, 4);
                if (fault == "context") fixture.CurrentContext = "VBAProject.Module1.Other";
                if (fault == "sha") request.ExpectedSha256 = "stale";
                if (fault == "mode") fixture.Project.Mode = 2;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.InspectLocalScalarsAsync(request).GetAwaiter().GetResult(), fault);
                Assert.AreEqual(0, calls, fault);
                Assert.IsFalse(VbeDebugInspection.IsActive, fault);
            }
        }

        [STATestMethod]
        public void EmptyOrSkippedPageRechecksContextAfterSelectionRead()
        {
            foreach (string source in new[] { "Sub Inspect()\r\nEnd Sub", "Sub Inspect()\r\nDim unsupported As Variant\r\nEnd Sub" })
            {
                var fixture = new ScalarFixture(source);
                int evaluations = 0;
                fixture.Service.LocalScalarEvaluator = request =>
                {
                    evaluations++;
                    return Task.FromResult<object>(null);
                };
                fixture.Module.CodePane.OnGetSelection = () => fixture.CurrentContext = "VBAProject.Module1.Other";
                Assert.ThrowsException<InvalidOperationException>(() =>
                    fixture.Service.InspectLocalScalarsAsync(fixture.Request(0, 4)).GetAwaiter().GetResult());
                Assert.AreEqual(0, evaluations);
                Assert.IsFalse(VbeDebugInspection.IsActive);
            }
        }

        public sealed class ScalarReadback
        {
            public string Expression { get; set; }
            public string Context { get; set; }
            public string Value { get; set; }
        }

        private sealed class ScalarFixture
        {
            internal readonly ScalarProject Project = new ScalarProject { Name = "VBAProject", Mode = 1 };
            internal readonly ScalarModule Module;
            internal readonly VbeDebug Service;
            internal string CurrentContext;
            internal string Context => "VBAProject.Module1.Inspect";

            internal ScalarFixture(string source)
            {
                Module = new ScalarModule(source);
                var component = new ScalarComponent { Name = "Module1", CodeModule = Module };
                Module.Parent = component;
                Project.VBComponents.Add(component);
                var vbe = new ScalarVbe { ActiveVBProject = Project, ActiveCodePane = Module.CodePane };
                vbe.VBProjects.Add(Project);
                Service = new VbeDebug(vbe)
                {
                    LocalContextReader = () => CurrentContext ?? Context,
                    EnsureScalarDialogAbsent = () => { }
                };
            }

            internal Request Request(int offset, int limit) => new Request
            {
                Project = Project.Name,
                Module = "Module1",
                Procedure = "Inspect",
                ExpectedMode = 1,
                ExpectedSha256 = VbeDebugTestsHash(Module.Code),
                Offset = offset,
                Limit = limit
            };
        }

        private static string VbeDebugTestsHash(string source)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(source))).Replace("-", "").ToLowerInvariant();
        }

        public sealed class ScalarVbe
        {
            public List<ScalarProject> VBProjects { get; } = new List<ScalarProject>();
            public ScalarProject ActiveVBProject { get; set; }
            public ScalarPane ActiveCodePane { get; set; }
            public object ActiveWindow { get; set; }
        }

        public sealed class ScalarProject
        {
            public string Name { get; set; }
            public int Mode { get; set; }
            public List<ScalarComponent> VBComponents { get; } = new List<ScalarComponent>();
        }

        public sealed class ScalarComponent
        {
            public string Name { get; set; }
            public ScalarModule CodeModule { get; set; }
        }

        public sealed class ScalarModule
        {
            public ScalarComponent Parent { get; set; }
            public string Code { get; set; }
            public int CountOfLines => Code.Split(new[] { "\r\n" }, StringSplitOptions.None).Length;
            public ScalarLines Lines { get; }
            public ScalarRange ProcStartLine { get; } = new ScalarRange { Value = 1 };
            public ScalarRange ProcBodyLine { get; } = new ScalarRange { Value = 1 };
            public ScalarRange ProcCountLines { get; }
            public ScalarPane CodePane { get; } = new ScalarPane();
            public ScalarModule(string code) { Code = code; Lines = new ScalarLines(this); ProcCountLines = new ScalarRange { Value = CountOfLines }; }
        }

        public sealed class ScalarLines
        {
            private readonly ScalarModule module;
            public ScalarLines(ScalarModule module) { this.module = module; }
            public string this[int first, int count] => string.Join("\r\n", module.Code.Split(new[] { "\r\n" }, StringSplitOptions.None)
                .Skip(first - 1).Take(count));
        }

        public sealed class ScalarRange
        {
            public int Value { get; set; }
            public int this[string name, int kind] => Value;
        }

        public sealed class ScalarPane
        {
            public int Line = 1, Column = 1, EndLine = 1, EndColumn = 1;
            public Action OnGetSelection { get; set; }
            public object Window { get; } = new object();
            public void GetSelection(ref int line, ref int column, ref int endLine, ref int endColumn)
            { line = Line; column = Column; endLine = EndLine; endColumn = EndColumn; OnGetSelection?.Invoke(); }
        }
    }
}
