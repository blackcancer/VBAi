using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugTests
    {
        private const string Code = "Debug.Print 1\r\nDebug.Print 2";

        private static Fixture Create(int mode = 2)
        {
            var project = new FakeProject { Name = "VBAProject", FileName = @"C:\Temp\Debug.xlsm", Mode = mode };
            var component = new FakeComponent { Name = "Module1" };
            var module = new FakeModule(component, Code);
            component.CodeModule = module;
            project.VBComponents.Add(component);
            var vbe = new FakeVbe { ActiveVBProject = project, ActiveCodePane = module.CodePane };
            vbe.VBProjects.Add(project);
            var bar = new FakeBar { Name = "Debug" };
            vbe.CommandBars.Add(bar);
            return new Fixture { Vbe = vbe, Project = project, Module = module, Bar = bar,
                Service = new VbeDebug(vbe) };
        }

        private static Request Location(Fixture f, int line = 1)
        {
            return new Request { Project = f.Project.Name, Module = "Module1", StartLine = line,
                ExpectedSha256 = Sha(Code), ExpectedMode = f.Project.Mode };
        }

        private static string Sha(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }

        [TestMethod]
        public void StateReportsProjectModeAndCurrentEditorSelection()
        {
            var f = Create(1);
            f.Module.CodePane.SetSelection(2, 3, 2, 8);
            dynamic state = f.Service.State(f.Project.Name);
            Assert.AreEqual("VBAProject", (string)state.Project);
            Assert.AreEqual(1, (int)state.Mode);
            Assert.AreEqual("VBAProject", (string)state.SelectedProject);
            Assert.AreEqual(@"C:\Temp\Debug.xlsm", (string)state.SelectedProjectPath);
            Assert.AreEqual("Module1", (string)state.ActiveModule);
            Assert.AreEqual(2, (int)state.Selection.StartLine);
            Assert.AreEqual(8, (int)state.Selection.EndColumn);
        }

        [TestMethod]
        public void StatePreservesProjectModeWhenNativeSelectionThrows()
        {
            var f = Create();
            f.Module.CodePane.FailGetSelection = true;
            dynamic state = f.Service.State(f.Project.Name);
            Assert.AreEqual(2, (int)state.Mode);
            StringAssert.Contains((string)state.Selection.Error, "selection unavailable");
        }

        [TestMethod]
        public void CommandsFilterNestedPathsAndPageWithoutExecutingControls()
        {
            var f = Create();
            var menu = new FakeControl { Caption = "&Windows", Id = 1 };
            var immediate = new FakeControl { Caption = "&Immediate Window", Id = 2554 };
            menu.Controls.Add(immediate);
            f.Bar.Controls.Add(menu);
            f.Bar.Controls.Add(new FakeControl { Caption = "Locals Window", Id = 2555 });
            var matches = ((IEnumerable)f.Service.ListCommands("immediate", 0, 0)).Cast<object>().ToArray();
            Assert.AreEqual(1, matches.Length);
            Assert.AreEqual("Debug > &Windows > &Immediate Window", (string)((dynamic)matches[0]).Path);
            Assert.AreEqual(2554, (int)((dynamic)matches[0]).Id);
            Assert.AreEqual(1, ((IEnumerable)f.Service.ListCommands(null, 1, 1)).Cast<object>().Count());
            Assert.AreEqual(0, immediate.ExecuteCount);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.ListCommands(null, -1, 1));
        }

        [TestMethod]
        public void OpenDebugPaneRequiresMatchingEnabledNativeControl()
        {
            var f = Create();
            var wrongId = new FakeControl { Caption = "Locals Window", Id = 7 };
            var disabled = new FakeControl { Caption = "Locals Window", Id = 2555, Enabled = false };
            f.Bar.Controls.Add(wrongId);
            f.Bar.Controls.Add(disabled);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenDebugPane("locals", null));
            Assert.AreEqual(0, wrongId.ExecuteCount);
            Assert.AreEqual(0, disabled.ExecuteCount);

            var valid = new FakeControl { Caption = "&Locals Window", Id = 2555 };
            f.Bar.Controls.Add(valid);
            dynamic result = f.Service.OpenDebugPane(" LOCALS ", null);
            Assert.AreEqual(2555, (int)result.ControlId);
            Assert.IsTrue((bool)result.VerificationPending);
            Assert.AreEqual(1, valid.ExecuteCount);
            Assert.ThrowsException<ArgumentException>(() => f.Service.OpenDebugPane("stack", null));
        }

        [TestMethod]
        public void GlobalDebugCommandChecksModeAndDoesNotClaimBreakpointClearVerification()
        {
            var f = Create(2);
            var clear = new FakeControl { Caption = "Clear All Breakpoints", Id = 579 };
            f.Bar.Controls.Add(clear);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(
                new Request { Project = f.Project.Name, ExpectedMode = 1, Action = "clear_all_breakpoints" }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(
                new Request { Project = f.Project.Name, ExpectedMode = 2, Action = "reset" }));
            Assert.AreEqual(0, clear.ExecuteCount);

            dynamic result = f.Service.ExecuteGlobalDebugCommand(new Request {
                Project = f.Project.Name, ExpectedMode = 2, Action = "clear_all_breakpoints" });
            Assert.AreEqual(1, clear.ExecuteCount);
            Assert.AreEqual("Unverified", (string)result.Verification);
            StringAssert.Contains((string)result.VerificationLimit, "no breakpoint inventory");
        }

        [TestMethod]
        public void SelectCodeRejectsStaleHashAndExpressionBeforeChangingSelection()
        {
            var f = Create();
            var request = Location(f);
            request.ExpectedSha256 = Sha("different code");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCode(request));
            request.ExpectedSha256 = Sha(Code);
            request.StartColumn = 1;
            request.EndColumn = 6;
            request.Expression = "wrong";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCode(request));
            Assert.AreEqual(0, f.Module.CodePane.ShowCount);
        }

        [TestMethod]
        public void SelectCodeReturnsExactTextOnlyAfterNativePaneRetainsSelection()
        {
            var f = Create();
            var request = Location(f);
            request.StartColumn = 1;
            request.EndColumn = 6;
            request.Expression = "Debug";
            dynamic result = f.Service.SelectCode(request);
            Assert.AreEqual("Debug", (string)result.SelectedText);
            Assert.AreEqual(1, f.Module.CodePane.ShowCount);
            f.Module.CodePane.RetainSelection = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCode(request));
        }

        [TestMethod]
        public void SelectCodeRangeValidatesBoundsAndConfirmsNativeSelection()
        {
            var f = Create();
            var request = Location(f);
            request.StartColumn = 7;
            request.EndLine = 2;
            request.EndColumn = 12;
            dynamic result = f.Service.SelectCodeRange(request);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(2, f.Module.CodePane.EndLine);
            request.EndLine = 3;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.SelectCodeRange(request));
            request.EndLine = 2;
            f.Module.CodePane.RetainSelection = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCodeRange(request));
        }

        [TestMethod]
        public void BreakpointCommandRejectsNonExecutableLineAndDisabledControl()
        {
            var f = Create();
            f.Module.Code = "' comment\r\nDebug.Print 2";
            var request = Location(f);
            request.ExpectedSha256 = Sha(f.Module.Code);
            request.Action = "toggle_breakpoint";
            request.ControlId = 123;
            request.ControlCaption = "Toggle Breakpoint";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(request));

            request.StartLine = 2;
            var disabled = new FakeControl { Caption = request.ControlCaption, Id = request.ControlId, Enabled = false };
            f.Bar.Controls.Add(disabled);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(request));
            Assert.AreEqual(0, disabled.ExecuteCount);
        }

        [TestMethod]
        public void BreakpointInvocationReportsUnverifiedNativeEffect()
        {
            var f = Create();
            var request = Location(f);
            request.Action = "toggle_breakpoint";
            request.ControlId = 123;
            request.ControlCaption = "Toggle Breakpoint";
            var control = new FakeControl { Caption = request.ControlCaption, Id = request.ControlId };
            f.Bar.Controls.Add(control);
            dynamic result = f.Service.InvokeCommand(request);
            Assert.AreEqual(1, control.ExecuteCount);
            Assert.AreEqual("Unverified", (string)result.Verification);
            Assert.IsFalse((bool)result.VerificationPending);
            StringAssert.Contains((string)result.VerificationLimit, "no breakpoint inventory");
        }

        private sealed class Fixture
        {
            public FakeVbe Vbe;
            public FakeProject Project;
            public FakeModule Module;
            public FakeBar Bar;
            public VbeDebug Service;
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
            public List<FakeBar> CommandBars { get; } = new List<FakeBar>();
            public FakeProject ActiveVBProject { get; set; }
            public FakePane ActiveCodePane { get; set; }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }

        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public FakeModule CodeModule { get; set; }
        }

        public sealed class FakeModule
        {
            public FakeComponent Parent { get; }
            public FakePane CodePane { get; }
            public FakeLines Lines { get; }
            public string Code { get; set; }
            public int CountOfLines => Code.Split(new[] { "\r\n" }, StringSplitOptions.None).Length;

            public FakeModule(FakeComponent parent, string code)
            {
                Parent = parent;
                Code = code;
                CodePane = new FakePane { CodeModule = this };
                Lines = new FakeLines(this);
            }
        }

        public sealed class FakeLines
        {
            private readonly FakeModule module;
            public FakeLines(FakeModule module) { this.module = module; }
            public string this[int start, int count]
            {
                get
                {
                    var lines = module.Code.Split(new[] { "\r\n" }, StringSplitOptions.None);
                    return string.Join("\r\n", lines.Skip(start - 1).Take(count));
                }
            }
        }

        public sealed class FakePane
        {
            public FakeModule CodeModule { get; set; }
            public int ShowCount { get; private set; }
            public bool FailGetSelection { get; set; }
            public bool RetainSelection { get; set; } = true;
            public int StartLine { get; private set; } = 1;
            public int StartColumn { get; private set; } = 1;
            public int EndLine { get; private set; } = 1;
            public int EndColumn { get; private set; } = 1;
            public void Show() { ShowCount++; }
            public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
            {
                if (!RetainSelection) return;
                StartLine = startLine; StartColumn = startColumn;
                EndLine = endLine; EndColumn = endColumn;
            }
            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                if (FailGetSelection) throw new InvalidOperationException("selection unavailable");
                startLine = StartLine; startColumn = StartColumn;
                endLine = EndLine; endColumn = EndColumn;
            }
        }

        public sealed class FakeBar
        {
            public string Name { get; set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
        }

        public sealed class FakeControl
        {
            public string Caption { get; set; }
            public int Id { get; set; }
            public bool Enabled { get; set; } = true;
            public int ExecuteCount { get; private set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
            public void Execute() { ExecuteCount++; }
        }
    }
}
