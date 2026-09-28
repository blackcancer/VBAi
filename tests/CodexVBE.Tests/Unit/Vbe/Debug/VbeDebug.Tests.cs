namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class DebugCommandTests
    {
        [TestMethod]
        public void CommandCatalogSearchesNestedMenusAndPagesResults()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            dynamic all = commands.ListCommands(null, 0, 0);
            Assert.AreEqual(5, ((object[])all).Length);
            dynamic match = commands.ListCommands("WATCh", 0, 10);
            Assert.AreEqual(1, ((object[])match).Length);
            Assert.AreEqual("Debug > &Watch Window", (string)match[0].Path);
            Assert.AreEqual(2556, (int)match[0].Id);
            Assert.AreEqual(2, ((object[])commands.ListCommands(null, 1, 2)).Length);
            Assert.AreEqual(0, ((object[])commands.ListCommands("missing", 0, 10)).Length);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => commands.ListCommands(null, -1, 1));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => commands.ListCommands(null, 0, -1));
        }

        [TestMethod]
        public void CommandSearchIgnoresMenuAcceleratorsInsideWords()
        {
            var host = Host();
            host.CommandBars[0].Name = "Fe&nêtre";
            host.CommandBars[0].Controls.Add(new FakeControl { Caption = "Fr&actionner", Id = 302, Enabled = true });
            var commands = new VbeDebug(host);
            dynamic result = commands.ListCommands("Fractionner", 0, 10);
            Assert.AreEqual(1, ((object[])result).Length);
            Assert.AreEqual(302, (int)result[0].Id);
            Assert.AreEqual("Fr&actionner", (string)result[0].Caption);
            Assert.AreEqual(6, ((object[])commands.ListCommands("Fenêtre", 0, 10)).Length);
        }

        [TestMethod]
        public void OpeningDebugPanesRequiresAnEnabledMatchingNativeCommand()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            Assert.ThrowsException<ArgumentException>(() => commands.OpenDebugPane("output", null));
            Assert.ThrowsException<InvalidOperationException>(() => commands.OpenDebugPane("immediate", null));
            dynamic locals = commands.OpenDebugPane("LOCALS", null);
            Assert.AreEqual(2555, (int)locals.ControlId);
            Assert.IsTrue((bool)locals.VerificationPending);
            Assert.AreEqual(1, host.CommandBars[0].Controls[0].Controls[0].Executions);
            dynamic watches = commands.OpenDebugPane("watches", null);
            Assert.AreEqual(2556, (int)watches.ControlId);
            Assert.AreEqual(1, host.CommandBars[0].Controls[1].Executions);
        }

        [TestMethod]
        public void DebugStateReportsTheActiveSelectionWithoutMovingIt()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            dynamic idle = commands.State("Projet");
            Assert.AreEqual(2, (int)idle.Mode);
            Assert.IsNull((object)idle.Selection);
            host.ActiveVBProject = host.VBProjects.Single();
            host.ActiveCodePane = new FakeCodePane();
            dynamic selected = commands.State("Projet");
            Assert.AreEqual("Projet", (string)selected.SelectedProject);
            Assert.AreEqual("Module1", (string)selected.ActiveModule);
            Assert.AreEqual(4, (int)selected.Selection.StartLine);
            Assert.AreEqual(9, (int)selected.Selection.EndColumn);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugTests
    {
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
            var menu = new FakeControl
            {
                Caption = "&Windows",
                Id = 1
            };
            var immediate = new FakeControl
            {
                Caption = "&Immediate Window",
                Id = 2554
            };
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
            var wrongId = new FakeControl
            {
                Caption = "Locals Window",
                Id = 7
            };
            var disabled = new FakeControl
            {
                Caption = "Locals Window",
                Id = 2555,
                Enabled = false
            };
            f.Bar.Controls.Add(wrongId);
            f.Bar.Controls.Add(disabled);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenDebugPane("locals", null));
            Assert.AreEqual(0, wrongId.ExecuteCount);
            Assert.AreEqual(0, disabled.ExecuteCount);
            var valid = new FakeControl
            {
                Caption = "&Locals Window",
                Id = 2555
            };
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
            var clear = new FakeControl
            {
                Caption = "Clear All Breakpoints",
                Id = 579
            };
            f.Bar.Controls.Add(clear);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 1, Action = "clear_all_breakpoints" }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 2, Action = "reset" }));
            Assert.AreEqual(0, clear.ExecuteCount);
            dynamic result = f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 2, Action = "clear_all_breakpoints" });
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
            var disabled = new FakeControl
            {
                Caption = request.ControlCaption,
                Id = request.ControlId,
                Enabled = false
            };
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
            var control = new FakeControl
            {
                Caption = request.ControlCaption,
                Id = request.ControlId
            };
            f.Bar.Controls.Add(control);
            dynamic result = f.Service.InvokeCommand(request);
            Assert.AreEqual(1, control.ExecuteCount);
            Assert.AreEqual("Unverified", (string)result.Verification);
            Assert.IsFalse((bool)result.VerificationPending);
            StringAssert.Contains((string)result.VerificationLimit, "no breakpoint inventory");
        }

        [TestMethod]
        public void WatchDialogsScheduleOnlyTheMatchingNativeCommand()
        {
            var f = Create(1);
            var add = new FakeControl
            {
                Caption = "&Add Watch...",
                Id = 1820
            };
            var edit = new FakeControl
            {
                Caption = "Edit Watch...",
                Id = 940
            };
            var quick = new FakeControl
            {
                Caption = "Quick Watch...",
                Id = 229
            };
            f.Bar.Controls.AddRange(new[] { add, edit, quick });
            var context = new RecordingContext();
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                dynamic added = f.Service.QueueAddWatchDialog(new Request { Project = f.Project.Name, Module = "Module1", Expression = "counter", WatchType = "break_when_changed", ExpectedMode = 1 });
                dynamic edited = f.Service.QueueEditWatchDialog(new Request { Project = f.Project.Name, Expression = "counter", Context = "VBAProject.Module1", NewExpression = "counter + 1", ExpectedMode = 1 });
                var location = Location(f);
                location.Expression = "Debug";
                location.StartColumn = 1;
                location.EndColumn = 6;
                dynamic watched = f.Service.QueueQuickWatchDialog(location);
                Assert.IsTrue((bool)added.Scheduled);
                Assert.AreEqual(940, (int)edited.ControlId);
                Assert.AreEqual(229, (int)watched.ControlId);
                Assert.AreEqual(3, context.Count);
                Assert.AreEqual(0, add.ExecuteCount + edit.ExecuteCount + quick.ExecuteCount);
                context.RunAll();
                Assert.AreEqual(1, add.ExecuteCount);
                Assert.AreEqual(1, edit.ExecuteCount);
                Assert.AreEqual(1, quick.ExecuteCount);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public void WatchCommandsRejectStaleModeUnavailableControlAndBadType()
        {
            var f = Create(1);
            var add = new Request
            {
                Project = f.Project.Name,
                Module = "Module1",
                Expression = "counter",
                ExpectedMode = 1
            };
            add.WatchType = "invalid";
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueAddWatchDialog(add));
            add.WatchType = "expression";
            add.ExpectedMode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
            add.ExpectedMode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
            var edit = new Request
            {
                Project = f.Project.Name,
                Expression = "counter",
                Context = "VBAProject.Module1",
                NewExpression = "counter + 1",
                ExpectedMode = 2
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueEditWatchDialog(edit));
            edit.ExpectedMode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueEditWatchDialog(edit));
            var remove = new Request
            {
                Project = f.Project.Name,
                Expression = "counter",
                Context = "VBAProject.Module1",
                ExpectedMode = 2
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemoveSelectedWatch(remove));
            remove.ExpectedMode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemoveSelectedWatch(remove));
        }

        [TestMethod]
        public void DeleteWatchExecutesNativeCommandAndRequiresSeparateVerification()
        {
            var f = Create(1);
            var wrong = new FakeControl
            {
                Caption = "Delete Watch",
                Id = 1083,
                Enabled = false
            };
            var delete = new FakeControl
            {
                Caption = "&Supprimer un espion",
                Id = 1083
            };
            f.Bar.Controls.AddRange(new[] { wrong, delete });
            dynamic result = f.Service.RemoveSelectedWatch(new Request { Project = f.Project.Name, Expression = "counter", Context = "VBAProject.Module1", ExpectedMode = 1 });
            Assert.IsTrue((bool)result.Executed);
            Assert.IsTrue((bool)result.VerificationPending);
            Assert.AreEqual(0, wrong.ExecuteCount);
            Assert.AreEqual(1, delete.ExecuteCount);
        }

        [TestMethod]
        public void OptionsDialogRequiresNativeControlAndUiContext()
        {
            var f = Create();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueDebugOptionsDialog());
            var wrong = new FakeControl
            {
                Caption = "Options...",
                Id = 522,
                Enabled = false
            };
            var options = new FakeControl
            {
                Caption = "&Options...",
                Id = 522
            };
            f.Bar.Controls.AddRange(new[] { wrong, options });
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(null);
            try
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueDebugOptionsDialog());
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }

            var context = new RecordingContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                dynamic result = f.Service.QueueDebugOptionsDialog();
                Assert.IsTrue((bool)result.Scheduled);
                Assert.AreEqual(0, options.ExecuteCount);
                context.RunAll();
                Assert.AreEqual(1, options.ExecuteCount);
                Assert.AreEqual(0, wrong.ExecuteCount);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public void CompileRequiresOwnProjectNameAndDesignMode()
        {
            var f = Create(2);
            var compile = new FakeControl
            {
                Caption = "&Compile OtherProject",
                Id = 578
            };
            f.Bar.Controls.Add(compile);
            var request = new Request
            {
                Project = f.Project.Name,
                ExpectedMode = 2
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CompileProject(request));
            Assert.AreEqual(0, compile.ExecuteCount);
            compile.Caption = "&Compiler VBAProject";
            dynamic result = f.Service.CompileProject(request);
            Assert.AreEqual(578, (int)result.ControlId);
            Assert.AreEqual(1, compile.ExecuteCount);
            request.ExpectedMode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CompileProject(request));
        }

        [TestMethod]
        public void GlobalBreakAndResetVerifyOnlyObservedModeChanges()
        {
            var f = Create(0);
            var breaker = new FakeControl
            {
                Caption = "&Break",
                Id = 189,
                OnExecute = () => f.Project.Mode = 1
            };
            var reset = new FakeControl
            {
                Caption = "&Reset",
                Id = 228,
                OnExecute = () => f.Project.Mode = 2
            };
            f.Bar.Controls.AddRange(new[] { breaker, reset });
            dynamic broken = f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, Action = "break", ExpectedMode = 0 });
            Assert.AreEqual("Verified", (string)broken.Verification);
            Assert.IsFalse((bool)broken.VerificationPending);
            dynamic resetResult = f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, Action = "reset", ExpectedMode = 1 });
            Assert.AreEqual("Verified", (string)resetResult.Verification);
            Assert.IsFalse((bool)resetResult.VerificationPending);
            Assert.AreEqual(1, breaker.ExecuteCount);
            Assert.AreEqual(1, reset.ExecuteCount);
        }

        [TestMethod]
        public void SignatureDialogRefusesCollisionWithoutSafeDocumentComponent()
        {
            var f = Create(2);
            f.Bar.Controls.Add(new FakeControl { Caption = "Remove Module1", Id = 746 });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(new Request { Project = f.Project.Name, ExpectedMode = 2 }));
            Assert.AreEqual(0, f.Bar.Controls[0].ExecuteCount);
        }

        [TestMethod]
        public void SignatureDialogSchedulesOnlyToolsCommandWhenThereIsNoRemovalCollision()
        {
            var f = Create(2);
            f.Bar.Name = "Tools";
            var signature = new FakeControl
            {
                Caption = "Digital Signature...",
                Id = 746
            };
            f.Bar.Controls.Add(signature);
            var context = new RecordingContext();
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                dynamic result = f.Service.QueueSignatureDialog(new Request { Project = f.Project.Name, ExpectedMode = 2 });
                Assert.IsTrue((bool)result.Scheduled);
                Assert.AreEqual(0, signature.ExecuteCount);
                context.RunAll();
                Assert.AreEqual(1, signature.ExecuteCount);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public void RunSubRequiresParameterlessStandardModuleAndExactSourceHash()
        {
            var f = Create(2);
            f.Project.VBComponents[0].Type = 1;
            f.Module.Code = "Public Sub TryMe()\r\nEnd Sub";
            var run = new FakeControl
            {
                Caption = "Run Sub",
                Id = 186
            };
            f.Bar.Controls.Add(run);
            var request = new Request
            {
                Project = f.Project.Name,
                Module = "Module1",
                Procedure = "TryMe",
                ExpectedMode = 2,
                ExpectedSha256 = Sha(f.Module.Code)
            };
            dynamic result = f.Service.RunSub(request);
            Assert.IsTrue((bool)result.Executed);
            Assert.AreEqual(1, run.ExecuteCount);
            Assert.AreEqual(1, request.StartLine);
            request.ExpectedSha256 = Sha("stale source");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(request));
            request.ExpectedSha256 = Sha(f.Module.Code);
            f.Module.Code = "Public Sub TryMe(value As Long)\r\nEnd Sub";
            request.ExpectedSha256 = Sha(f.Module.Code);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(request));
            Assert.AreEqual(1, run.ExecuteCount);
        }

        [TestMethod]
        public void ObjectBrowserRequiresItsNativeControlAndReportsObservedVisibility()
        {
            var f = Create();
            var host = new VbeEditorWindowsTests.FakeVbe();
            var windows = new VbeEditorWindows(host);
            Assert.ThrowsException<ArgumentNullException>(() => f.Service.OpenObjectBrowser(null));
            f.Bar.Controls.Add(new FakeControl { Caption = "Unrelated", Id = 473 });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenObjectBrowser(windows));
            var browser = new FakeControl
            {
                Caption = "&Object Browser",
                Id = 473
            };
            f.Bar.Controls.Add(browser);
            dynamic pending = f.Service.OpenObjectBrowser(windows);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            host.Windows.Add(new VbeEditorWindowsTests.FakeWindow(host) { Caption = "Object Browser", Type = 2, Visible = true });
            dynamic visible = f.Service.OpenObjectBrowser(windows);
            Assert.AreEqual("Visible", (string)visible.Verification);
            Assert.IsTrue((bool)visible.AlreadyVisible);
            Assert.AreEqual(2, browser.ExecuteCount);
        }

        [TestMethod]
        public void StepIntoVerifiesChangedCodeLocationButNotUnchangedLocation()
        {
            var f = Create(1);
            var command = new FakeControl
            {
                Caption = "Step Into",
                Id = 303,
                OnExecute = () => f.Module.CodePane.SetSelection(2, 1, 2, 1)
            };
            f.Bar.Controls.Add(command);
            var request = Location(f);
            request.Action = "step_into";
            request.ControlId = 303;
            request.ControlCaption = "Step Into";
            dynamic verified = f.Service.InvokeCommand(request);
            Assert.AreEqual("Verified", (string)verified.Verification);
            StringAssert.Contains((string)verified.Evidence, "Module1:1 to Module1:2");
            command.OnExecute = null;
            dynamic pending = f.Service.InvokeCommand(request);
            Assert.AreEqual("Unverified", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual(2, command.ExecuteCount);
        }

        [TestMethod]
        public void DebugCaptionPolicyCoversEverySupportedActionAndMode()
        {
            var allowed = typeof(VbeDebug).GetMethod("IsAllowed", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNotNull(allowed);
            var accepted = new[]
            {
                Tuple.Create("toggle_breakpoint", "Point d'arrêt", 2),
                Tuple.Create("run", "Run Sub", 2),
                Tuple.Create("continue", "Continue", 1),
                Tuple.Create("step_into", "Step Into", 1),
                Tuple.Create("step_over", "Step Over", 1),
                Tuple.Create("step_out", "Step Out", 1),
                Tuple.Create("run_to_cursor", "Run To Cursor", 1),
                Tuple.Create("set_next_statement", "Set Next Statement", 1)
            };
            foreach (var entry in accepted)
            {
                Assert.AreEqual(true, allowed.Invoke(null, new object[] { entry.Item1, entry.Item2, entry.Item3 }), entry.Item1);
                Assert.AreEqual(false, allowed.Invoke(null, new object[] { entry.Item1, entry.Item2, 0 }), entry.Item1 + " in run mode");
            }

            Assert.AreEqual(false, allowed.Invoke(null, new object[] { "unknown", "Run Sub", 2 }));
            Assert.AreEqual(false, allowed.Invoke(null, new object[] { "run", "Reset", 2 }));
        }

        [TestMethod]
        public void SignatureCollisionWithDocumentStillRefusesIfRemovalRemainsEnabled()
        {
            var f = Create(2);
            var document = new FakeComponent
            {
                Name = "ThisWorkbook",
                Type = 100
            };
            document.CodeModule = new FakeModule(document, "");
            f.Project.VBComponents.Add(document);
            f.Bar.Controls.Add(new FakeControl { Caption = "Remove Module1", Id = 746 });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(new Request { Project = f.Project.Name, ExpectedMode = 2 }));
            Assert.AreEqual(1, document.CodeModule.CodePane.ShowCount);
        }

        [TestMethod]
        public void StateWithoutActivePanePreservesProjectModeWithoutInventingSelection()
        {
            var f = Create(2);
            f.Vbe.ActiveCodePane = null;
            dynamic state = f.Service.State(f.Project.Name);
            Assert.AreEqual(2, (int)state.Mode);
            Assert.IsNull((string)state.SelectedProject);
            Assert.IsNull((string)state.ActiveModule);
            Assert.IsNull((object)state.Selection);
        }

        [TestMethod]
        public void NativePaneRoutingUsesExactIdsForWatchAndImmediate()
        {
            var f = Create();
            var watch = new FakeControl
            {
                Caption = "&Espions",
                Id = 2556
            };
            var immediate = new FakeControl
            {
                Caption = "&Immediate Window",
                Id = 2554
            };
            f.Bar.Controls.AddRange(new[] { watch, immediate });
            dynamic watchResult = f.Service.OpenDebugPane("watches", null);
            dynamic immediateResult = f.Service.OpenDebugPane("immediate", null);
            Assert.AreEqual(2556, (int)watchResult.ControlId);
            Assert.AreEqual(2554, (int)immediateResult.ControlId);
            Assert.AreEqual(1, watch.ExecuteCount);
            Assert.AreEqual(1, immediate.ExecuteCount);
            Assert.ThrowsException<ArgumentException>(() => f.Service.OpenDebugPane(null, null));
        }

        [TestMethod]
        public void GlobalDebugRejectsInvalidActionsAndReportsPendingModeTransitions()
        {
            var f = Create(0);
            Assert.ThrowsException<ArgumentException>(() => f.Service.ExecuteGlobalDebugCommand(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 0, Action = "invalid" }));
            var breaker = new FakeControl
            {
                Caption = "Break",
                Id = 189
            };
            f.Bar.Controls.Add(breaker);
            dynamic pendingBreak = f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 0, Action = "break" });
            Assert.AreEqual("Unverified", (string)pendingBreak.Verification);
            Assert.IsTrue((bool)pendingBreak.VerificationPending);
            Assert.AreEqual(1, breaker.ExecuteCount);
            f.Project.Mode = 1;
            var reset = new FakeControl
            {
                Caption = "Reset",
                Id = 228
            };
            f.Bar.Controls.Add(reset);
            dynamic pendingReset = f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 1, Action = "reset" });
            Assert.AreEqual("Unverified", (string)pendingReset.Verification);
            Assert.IsTrue((bool)pendingReset.VerificationPending);
            Assert.AreEqual(1, reset.ExecuteCount);
            f.Project.Mode = 0;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(new Request { Project = f.Project.Name, ExpectedMode = 0, Action = "clear_all_breakpoints" }));
        }

        [TestMethod]
        public void WatchDialogPreflightRejectsMissingIdentityLengthAndUiContext()
        {
            var f = Create(1);
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueAddWatchDialog(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueAddWatchDialog(new Request { Project = f.Project.Name, Module = "Module1", Expression = new string ('x', 1025) }));
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueEditWatchDialog(new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", NewExpression = new string ('x', 1025) }));
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueQuickWatchDialog(new Request { Project = f.Project.Name, Expression = "x", StartColumn = 1, EndColumn = 1 }));
            Assert.ThrowsException<ArgumentException>(() => f.Service.RemoveSelectedWatch(null));
            var add = new FakeControl
            {
                Caption = "Add Watch...",
                Id = 1820
            };
            f.Bar.Controls.Add(add);
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(null);
            try
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(new Request { Project = f.Project.Name, Module = "Module1", Expression = "x", ExpectedMode = 1 }));
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }

            Assert.AreEqual(0, add.ExecuteCount);
        }
    }
}
