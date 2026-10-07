namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;
    using VBAi;

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
            host.VBProjects.Single().VBComponents.Add(new FakeComponent { CodeModule = host.ActiveCodePane.CodeModule });
            dynamic selected = commands.State("Projet");
            Assert.AreEqual("Projet", (string)selected.SelectedProject);
            Assert.AreEqual("Module1", (string)selected.ActiveModule);
            Assert.AreEqual(4, (int)selected.Selection.StartLine);
            Assert.AreEqual(9, (int)selected.Selection.EndColumn);
        }

        [TestMethod]
        public void DebugStateDoesNotRevealAPaneFromAnotherProjectEvenIfTheProjectTreeSelectionMatches()
        {
            var host = Host();
            host.ActiveVBProject = host.VBProjects.Single();
            host.ActiveCodePane = new FakeCodePane();
            // The tree points at Projet, but its components do not own this pane.
            dynamic state = new VbeDebug(host).State("Projet");
            Assert.IsNull((object)state.SelectedProject);
            Assert.IsNull((object)state.SelectedProjectPath);
            Assert.IsNull((object)state.ActiveModule);
            Assert.IsNull((object)state.Selection);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using VBAi;

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
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueAddWatchDialog(new Request { Project = f.Project.Name, Module = "Module1", Expression = new string('x', 1025) }));
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueEditWatchDialog(new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", NewExpression = new string('x', 1025) }));
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

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using VBAi;

    public sealed partial class VbeDebugTests
    {
        [TestMethod]
        public void EveryCaptionVariantIsAllowedOnlyForItsActionAndMode()
        {
            var method = typeof(VbeDebug).GetMethod("IsAllowed", BindingFlags.NonPublic | BindingFlags.Static);
            var policies = new Dictionary<string, string[]>
            {
                ["toggle_breakpoint"] = new[] { "Toggle Breakpoint", "Point d'arrêt", "Point d’arrêt" },
                ["run"] = new[] { "Run Sub", "Exécuter Sub", "Exécuter la macro" },
                ["continue"] = new[] { "Continue", "Continuer" },
                ["step_into"] = new[] { "Step Into", "pas à pas détaillé" },
                ["step_over"] = new[] { "Step Over", "pas à pas principal" },
                ["step_out"] = new[] { "Step Out", "pas à pas sortant" },
                ["run_to_cursor"] = new[] { "Run To Cursor", "Exécuter jusqu'au curseur", "Exécuter jusqu’au curseur" },
                ["set_next_statement"] = new[] { "Set Next Statement", "Définir l'instruction suivante", "Définir l’instruction suivante" }
            };
            foreach (var policy in policies)
                foreach (var caption in policy.Value.Concat(new[] { "unrelated" }))
                    foreach (int mode in new[] { 0, 1, 2, 3 })
                    {
                        bool validMode = policy.Key == "toggle_breakpoint" ? mode == 1 || mode == 2
                            : policy.Key == "run" ? mode == 2 : mode == 1;
                        Assert.AreEqual(validMode && caption != "unrelated",
                            (bool)method.Invoke(null, new object[] { policy.Key, " &" + caption + " ", mode }),
                            policy.Key + "/" + caption + "/" + mode);
                    }
        }

        [TestMethod]
        public void WatchIdentityAndModeValidationCoversEveryMissingFieldAndAllowedType()
        {
            var f = Create(1);
            var add = new Request { Project = f.Project.Name, Module = "Module1", Expression = "x", ExpectedMode = 1 };
            var edit = new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", NewExpression = "y", ExpectedMode = 1 };
            var remove = new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", ExpectedMode = 1 };
            foreach (var pair in new[] {
                Tuple.Create(add, new[] { "Project", "Module", "Expression" }, (Action<Request>)(r => f.Service.QueueAddWatchDialog(r))),
                Tuple.Create(edit, new[] { "Project", "Expression", "Context", "NewExpression" }, (Action<Request>)(r => f.Service.QueueEditWatchDialog(r))),
                Tuple.Create(remove, new[] { "Project", "Expression", "Context" }, (Action<Request>)(r => f.Service.RemoveSelectedWatch(r))) })
            {
                Assert.ThrowsException<ArgumentException>(() => pair.Item3(null));
                foreach (var name in pair.Item2)
                {
                    var property = typeof(Request).GetProperty(name);
                    object original = property.GetValue(pair.Item1);
                    foreach (var value in new[] { null, "", " " })
                    {
                        property.SetValue(pair.Item1, value);
                        Assert.ThrowsException<ArgumentException>(() => pair.Item3(pair.Item1), name);
                    }
                    property.SetValue(pair.Item1, original);
                }
            }
            foreach (var type in new[] { null, "", " ", "expression", "break_when_true", "break_when_changed", "bad" })
            {
                add.WatchType = edit.WatchType = type;
                if (type == "bad")
                {
                    Assert.ThrowsException<ArgumentException>(() => f.Service.QueueAddWatchDialog(add));
                    Assert.ThrowsException<ArgumentException>(() => f.Service.QueueEditWatchDialog(edit));
                }
                else
                {
                    Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
                    Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueEditWatchDialog(edit));
                }
            }
            add.WatchType = null;
            f.Vbe.ActiveVBProject = new FakeProject { Name = "Other", Mode = 1 };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
            f.Vbe.ActiveVBProject = f.Project;
            add.Project = f.Project.FileName;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
            add.Module = "OtherModule";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
            add.Module = "Module1";
            f.Project.Mode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueAddWatchDialog(add));
        }

        [TestMethod]
        public void QuickWatchRejectsEachInvalidSelectionAndBreakModeCombination()
        {
            var f = Create(1);
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueQuickWatchDialog(null));
            foreach (var expression in new[] { null, "", " ", new string('x', 1025) })
            {
                var request = Location(f); request.Expression = expression; request.StartColumn = 1; request.EndColumn = 6;
                Assert.ThrowsException<ArgumentException>(() => f.Service.QueueQuickWatchDialog(request));
            }
            foreach (var columns in new[] { new[] { 0, 1 }, new[] { 1, 0 }, new[] { 2, 2 }, new[] { 3, 2 } })
            {
                var request = Location(f); request.Expression = "Debug"; request.StartColumn = columns[0]; request.EndColumn = columns[1];
                Assert.ThrowsException<ArgumentException>(() => f.Service.QueueQuickWatchDialog(request));
            }
            foreach (var mode in new[] { new[] { 1, 2 }, new[] { 2, 1 } })
            {
                f.Project.Mode = mode[0];
                var request = Location(f); request.ExpectedMode = mode[1]; request.Expression = "Debug"; request.StartColumn = 1; request.EndColumn = 6;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueQuickWatchDialog(request));
            }
            f.Project.Mode = 1;
            var valid = Location(f); valid.Expression = "Debug"; valid.StartColumn = 1; valid.EndColumn = 6;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueQuickWatchDialog(valid));
        }

        [TestMethod]
        public void ScheduledDialogsRequireContextAndContainNativeExecutionFailures()
        {
            var f = Create(1);
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (var entry in new[] {
                    Tuple.Create(1820, "Ajouter un espion", (Action)(() => f.Service.QueueAddWatchDialog(new Request { Project = f.Project.Name, Module = "Module1", Expression = "x", ExpectedMode = 1 }))),
                    Tuple.Create(940, "Modifier un espion", (Action)(() => f.Service.QueueEditWatchDialog(new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", NewExpression = "y", ExpectedMode = 1 }))),
                    Tuple.Create(229, "Espion express", (Action)(() => { var r = Location(f); r.Expression = "Debug"; r.StartColumn = 1; r.EndColumn = 6; f.Service.QueueQuickWatchDialog(r); })),
                    Tuple.Create(522, "Options...", (Action)(() => f.Service.QueueDebugOptionsDialog())) })
                {
                    f.Bar.Controls.Clear();
                    f.Bar.Controls.AddRange(new[] {
                        new FakeControl { Id = -1, Caption = entry.Item2 },
                        new FakeControl { Id = entry.Item1, Caption = entry.Item2, Enabled = false },
                        new FakeControl { Id = entry.Item1, Caption = null },
                        new FakeControl { Id = entry.Item1, Caption = "Unrelated" } });
                    Assert.ThrowsException<InvalidOperationException>(() => entry.Item3());
                    var command = new FakeControl
                    {
                        Id = entry.Item1,
                        Caption = entry.Item2,
                        OnExecute = () => { throw new InvalidOperationException("native dialog failure"); }
                    };
                    f.Bar.Controls.Add(command);
                    SynchronizationContext.SetSynchronizationContext(null);
                    Assert.ThrowsException<InvalidOperationException>(() => entry.Item3());
                    var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
                    entry.Item3(); Assert.AreEqual(1, context.Count);
                    context.RunAll(); Assert.AreEqual(1, command.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void GlobalCommandsValidateEveryModeAndObserveEachCaptionVariant()
        {
            foreach (var policy in new[] {
                Tuple.Create("break", 0, 189, new[] { "Arrêt", "Break" }),
                Tuple.Create("reset", 1, 228, new[] { "Réinitialiser", "Reset" }),
                Tuple.Create("clear_all_breakpoints", 2, 579, new[] { "Effacer tous les points d'arrêt", "Effacer tous les points d’arrêt", "Clear All Breakpoints" }) })
            {
                foreach (var caption in policy.Item4)
                {
                    var f = Create(policy.Item2);
                    var request = new Request { Project = f.Project.Name, ExpectedMode = policy.Item2, Action = policy.Item1 };
                    f.Bar.Controls.AddRange(new[] { new FakeControl { Id = -1, Caption = caption },
                        new FakeControl { Id = policy.Item3, Caption = caption, Enabled = false },
                        new FakeControl { Id = policy.Item3, Caption = null }, new FakeControl { Id = policy.Item3, Caption = "Other" } });
                    Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
                    var native = new FakeControl { Id = policy.Item3, Caption = caption }; f.Bar.Controls.Add(native);
                    dynamic result = f.Service.ExecuteGlobalDebugCommand(request);
                    Assert.IsTrue((bool)result.Executed); Assert.AreEqual(1, native.ExecuteCount);
                    if (policy.Item1 == "clear_all_breakpoints")
                    {
                        f.Project.Mode = request.ExpectedMode = 1;
                        f.Service.ExecuteGlobalDebugCommand(request); Assert.AreEqual(2, native.ExecuteCount);
                    }
                    else
                    {
                        f.Project.Mode = request.ExpectedMode = policy.Item2 == 0 ? 2 : 0;
                        Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
                    }
                }
            }
        }

        [TestMethod]
        public void ShowNextStatementRequiresTheExactActiveProjectComponentAndPane()
        {
            var f = Create(1);
            var request = new Request { Project = f.Project.Name, ExpectedMode = 1, Action = "show_next_statement" };
            f.Project.Mode = request.ExpectedMode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
            f.Project.Mode = request.ExpectedMode = 1;
            foreach (var project in new[] { null, new FakeProject { Name = f.Project.Name, Mode = 1 } })
            {
                f.Vbe.ActiveVBProject = project;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
            }
            f.Vbe.ActiveVBProject = f.Project;
            f.Vbe.ActiveCodePane = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
            var other = new FakeComponent { Name = "Missing" }; other.CodeModule = new FakeModule(other, Code);
            f.Vbe.ActiveCodePane = other.CodeModule.CodePane;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
            other.Name = "Module1";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ExecuteGlobalDebugCommand(request));
            f.Vbe.ActiveCodePane = f.Module.CodePane;
            foreach (var caption in new[] { "Afficher l'instruction suivante", "Afficher l’instruction suivante", "Show Next Statement" })
            {
                f.Bar.Controls.Clear(); f.Bar.Controls.Add(new FakeControl { Id = 1813, Caption = caption });
                dynamic result = f.Service.ExecuteGlobalDebugCommand(request);
                Assert.AreEqual("Unverified", (string)result.Verification); Assert.IsTrue((bool)result.VerificationPending);
                Assert.AreEqual("Active project", (string)result.Scope); Assert.AreEqual(1, f.Bar.Controls[0].ExecuteCount);
            }
        }

        [TestMethod]
        public void CompileDistinguishesAbsentDisabledAndForeignTargetsBeforeExecuting()
        {
            var f = Create(); var request = new Request { Project = f.Project.Name, ExpectedMode = 2 };
            Assert.ThrowsException<ArgumentException>(() => f.Service.CompileProject(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.CompileProject(new Request { Project = " " }));
            dynamic absent = f.Service.CompileProject(request);
            Assert.IsFalse((bool)absent.Executed); Assert.AreEqual("Absent", (string)absent.Capability);
            f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CompileProject(request));
            f.Project.Mode = 2;
            var compile = new FakeControl { Id = 578, Enabled = false, Caption = "VBAProject kompilieren" };
            f.Bar.Controls.Add(compile);
            dynamic disabled = f.Service.CompileProject(request);
            Assert.IsFalse((bool)disabled.Executed); Assert.AreEqual("Disabled", (string)disabled.Capability);
            Assert.AreEqual(0, compile.ExecuteCount);
            compile.Enabled = true;
            foreach (string foreign in new[] { null, "Unrelated", "Compile OtherVBAProject" })
            {
                compile.Caption = foreign;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.CompileProject(request));
            }
            compile.Caption = "VBAProject kompilieren";
            f.Service.CompileProject(request); Assert.AreEqual(1, compile.ExecuteCount);
            f.Vbe.ActiveVBProject = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CompileProject(request));
            Assert.AreEqual(1, compile.ExecuteCount);
        }

        [TestMethod]
        public void CodeSelectionsRejectEveryBoundsAndReadbackMismatchBeforeClaimingVerification()
        {
            foreach (var columns in new[] { new[] { 0, 1 }, new[] { 1, 0 }, new[] { -1, 1 }, new[] { 2, 1 }, new[] { 1, 100 } })
            {
                var f = Create(); var r = Location(f); r.StartColumn = columns[0]; r.EndColumn = columns[1];
                if (columns[0] == 0 || columns[1] == 0) Assert.ThrowsException<ArgumentException>(() => f.Service.SelectCode(r));
                else Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.SelectCode(r));
                Assert.AreEqual(0, f.Module.CodePane.ShowCount);
            }
            foreach (int mismatch in Enumerable.Range(0, 4))
            {
                var f = Create(); var r = Location(f, 2); r.StartColumn = 2; r.EndColumn = 6;
                f.Module.CodePane.SelectionReadback = selection => { selection[mismatch]++; return selection; };
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCode(r));
                r.EndLine = 2;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SelectCodeRange(r));
            }
            foreach (var values in new[] { new[] { 0, 1, 1 }, new[] { 1, 0, 1 }, new[] { 1, 1, 0 }, new[] { 2, 1, 1 } })
            {
                var f = Create(); var r = Location(f); r.StartColumn = values[0]; r.EndColumn = values[1]; r.EndLine = values[2];
                Assert.ThrowsException<ArgumentException>(() => f.Service.SelectCodeRange(r));
            }
            Assert.ThrowsException<ArgumentException>(() => Create().Service.SelectCodeRange(null));
            foreach (bool start in new[] { true, false })
            {
                var f = Create(); var r = Location(f); r.StartColumn = start ? 100 : 1; r.EndColumn = start ? 1 : 100; r.EndLine = 2;
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.SelectCodeRange(r));
            }
        }

        [TestMethod]
        public void InvokeCommandRejectsMissingIdentityStaleModeInactivePaneAndNonExecutableSource()
        {
            var f = Create(); var r = Location(f); r.Action = "toggle_breakpoint"; r.ControlId = 123; r.ControlCaption = "Toggle Breakpoint";
            foreach (var field in new[] { "Action", "ControlCaption", "Module", "ExpectedSha256" })
            {
                var property = typeof(Request).GetProperty(field); object original = property.GetValue(r);
                property.SetValue(r, " "); Assert.ThrowsException<ArgumentException>(() => f.Service.InvokeCommand(r)); property.SetValue(r, original);
            }
            r.ControlId = 0; Assert.ThrowsException<ArgumentException>(() => f.Service.InvokeCommand(r)); r.ControlId = 123;
            r.ExpectedMode = 1; Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r)); r.ExpectedMode = 2;
            r.Action = "run"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r)); r.Action = "toggle_breakpoint";
            r.Module = "Missing"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r)); r.Module = "Module1";
            foreach (int line in new[] { 0, 3 }) { r.StartLine = line; Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.InvokeCommand(r)); }
            r.StartLine = 1;
            f.Module.Code = " \r\nDebug.Print 2"; r.ExpectedSha256 = Sha(f.Module.Code);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
            f.Module.Code = Code; r.ExpectedSha256 = Sha(Code);
            foreach (var pane in new[] { null, new FakePane { CodeModule = f.Module } })
            {
                f.Vbe.ActiveCodePane = pane;
                f.Vbe.IgnoreActiveCodePaneAssignment = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
                f.Vbe.IgnoreActiveCodePaneAssignment = false;
            }
            f.Vbe.ActiveCodePane = f.Module.CodePane;
            f.Bar.Controls.Add(new FakeControl { Id = 123, Caption = "Unrelated" });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
        }

        [TestMethod]
        public void InvokeCommandSelectsRequestedPaneWhenShowAndFocusLeaveThePreviousPaneActive()
        {
            var f = Create(); var request = Location(f);
            request.Action = "toggle_breakpoint"; request.ControlId = 123; request.ControlCaption = "Toggle Breakpoint";
            var previous = new FakePane { CodeModule = f.Module };
            f.Vbe.ActiveCodePane = previous;
            f.Module.CodePane.OnSetSelection = () => Assert.AreSame(f.Module.CodePane, f.Vbe.ActiveCodePane);
            var command = new FakeControl { Id = 123, Caption = request.ControlCaption };
            f.Bar.Controls.Add(command);
            int assignmentsBefore = f.Vbe.ActiveCodePaneSetCount;

            dynamic result = f.Service.InvokeCommand(request);

            Assert.IsTrue((bool)result.Executed);
            Assert.AreEqual(1, command.ExecuteCount);
            Assert.AreSame(f.Module.CodePane, f.Vbe.ActiveCodePane);
            Assert.AreEqual(assignmentsBefore + 1, f.Vbe.ActiveCodePaneSetCount);
            Assert.AreEqual(1, f.Module.CodePane.ShowCount);
            Assert.AreEqual(1, f.Module.CodePane.Window.FocusCount);
        }

        [TestMethod]
        public void InvokeCommandDoesNotExecuteWhenTheHostIgnoresPaneSelection()
        {
            var f = Create(); var request = Location(f);
            request.Action = "toggle_breakpoint"; request.ControlId = 123; request.ControlCaption = "Toggle Breakpoint";
            var previous = new FakePane { CodeModule = f.Module };
            f.Vbe.ActiveCodePane = previous;
            f.Vbe.IgnoreActiveCodePaneAssignment = true;
            var command = new FakeControl { Id = 123, Caption = request.ControlCaption };
            f.Bar.Controls.Add(command);
            int assignmentsBefore = f.Vbe.ActiveCodePaneSetCount;

            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(request));

            Assert.AreEqual(0, command.ExecuteCount);
            Assert.AreSame(previous, f.Vbe.ActiveCodePane);
            Assert.AreEqual(assignmentsBefore + 1, f.Vbe.ActiveCodePaneSetCount);
        }

        [TestMethod]
        public void InvokeDebugReportsModeModuleLocationErrorsAndDeferredEffects()
        {
            foreach (var action in new[] { Tuple.Create("run", "Run Sub", 2), Tuple.Create("continue", "Continue", 1),
                Tuple.Create("step_into", "Step Into", 1), Tuple.Create("step_over", "Step Over", 1),
                Tuple.Create("step_out", "Step Out", 1), Tuple.Create("run_to_cursor", "Run To Cursor", 1) })
            {
                var f = Create(action.Item3); var r = Location(f); r.Action = action.Item1; r.ControlId = 321; r.ControlCaption = action.Item2;
                var native = new FakeControl { Id = 321, Caption = action.Item2 }; f.Bar.Controls.Add(native);
                native.OnExecute = () => f.Project.Mode = 0;
                dynamic mode = f.Service.InvokeCommand(r); Assert.AreEqual("Verified", (string)mode.Verification);
                f.Project.Mode = action.Item3; native.OnExecute = null;
                dynamic pending = f.Service.InvokeCommand(r); Assert.IsTrue((bool)pending.VerificationPending);
                if (action.Item1 != "run")
                {
                    native.OnExecute = () => f.Module.Parent.Name = "OtherModule";
                    dynamic module = f.Service.InvokeCommand(r); Assert.AreEqual("Verified", (string)module.Verification);
                    f.Module.Parent.Name = "Module1";
                    native.OnExecute = () => f.Module.CodePane.FailGetSelection = true;
                    dynamic failedSelection = f.Service.InvokeCommand(r); Assert.AreEqual("Unverified", (string)failedSelection.Verification);
                    f.Module.CodePane.FailGetSelection = false;
                }
                native.OnExecute = () => f.Vbe.VBProjects.Clear();
                dynamic failedState = f.Service.InvokeCommand(r);
                Assert.IsNull((object)failedState.StateAfter); Assert.IsNotNull((string)failedState.StateAfterError);
                Assert.AreEqual("Unverified", (string)failedState.Verification);
            }
        }

        [TestMethod]
        public void SetNextStatementRequiresMatchingPaneProcedureAndProcedureKind()
        {
            var f = Create(1); var r = Location(f, 2); r.Action = "set_next_statement"; r.ControlId = 400; r.ControlCaption = "Set Next Statement";
            f.Bar.Controls.Add(new FakeControl { Id = 400, Caption = r.ControlCaption });
            foreach (var pane in new[] { null, new FakePane { CodeModule = f.Module } })
            {
                f.Vbe.ActiveCodePane = pane;
                f.Vbe.IgnoreActiveCodePaneAssignment = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
                f.Vbe.IgnoreActiveCodePaneAssignment = false;
            }
            f.Vbe.ActiveCodePane = f.Module.CodePane;
            foreach (var names in new[] { new[] { (string)null, "TryMe" }, new[] { "TryMe", " " }, new[] { "TryMe", "Other" } })
            {
                f.Module.CodePane.SetSelection(1, 1, 1, 1);
                f.Module.ProcOfLine.NameAtLine = line => names[line == 1 ? 0 : 1];
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
            }
            f.Module.ProcOfLine.NameAtLine = line => "TryMe";
            f.Module.CodePane.SetSelection(1, 1, 1, 1);
            f.Module.ProcOfLine.KindAtLine = line => line == 1 ? 0 : 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.InvokeCommand(r));
            f.Module.ProcOfLine.KindAtLine = line => 0;
            f.Module.CodePane.SetSelection(1, 1, 1, 1);
            dynamic success = f.Service.InvokeCommand(r); Assert.IsTrue((bool)success.Executed);
            Assert.AreEqual(1, f.Bar.Controls[0].ExecuteCount);
        }

        [TestMethod]
        public void SignatureDialogValidatesModeIdentityNativeCaptionContextAndCollisionResolution()
        {
            var f = Create(2); var request = new Request { Project = f.Project.Name, ExpectedMode = 2 };
            Assert.ThrowsException<ArgumentException>(() => f.Service.QueueSignatureDialog(null));
            foreach (var invalid in new[] { new Request { Project = " " }, new Request { Project = f.Project.Name, ExpectedMode = 1 } })
                Assert.ThrowsException<ArgumentException>(() => f.Service.QueueSignatureDialog(invalid));
            f.Project.Mode = 1; Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(request)); f.Project.Mode = 2;
            foreach (var project in new[] { null, new FakeProject { Name = f.Project.Name, Mode = 2 } })
            { f.Vbe.ActiveVBProject = project; Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(request)); }
            f.Vbe.ActiveVBProject = f.Project;
            foreach (var bar in new[] { "Debug", "Tools", "Outils" })
            {
                f.Bar.Name = bar; f.Bar.Controls.Clear();
                f.Bar.Controls.AddRange(new[] { new FakeControl { Id = -1, Caption = "Signature" },
                    new FakeControl { Id = 746, Caption = "Digital Signature", Enabled = false }, new FakeControl { Id = 746, Caption = null } });
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(request));
            }
            var remove = new FakeControl { Id = 746, Caption = "Supprimer Module1" };
            var signature = new FakeControl { Id = 746, Caption = "&Signature numérique...", OnExecute = () => { throw new InvalidOperationException("native signature error"); } };
            var document = new FakeComponent { Name = "ThisWorkbook", Type = 100 }; document.CodeModule = new FakeModule(document, "");
            document.CodeModule.CodePane.OnShow = () => remove.Enabled = false;
            f.Project.VBComponents.Add(document); f.Bar.Name = "Outils"; f.Bar.Controls.Clear(); f.Bar.Controls.AddRange(new[] { remove, signature });
            var previous = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(null);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueSignatureDialog(request));
                Assert.AreEqual(1, document.CodeModule.CodePane.ShowCount);
                var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
                f.Service.QueueSignatureDialog(request); context.RunAll(); Assert.AreEqual(1, signature.ExecuteCount);
                Assert.AreEqual(0, remove.ExecuteCount);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void RunSubRejectsEachInvalidIdentityModuleModeHashAndUnavailableRunControl()
        {
            var f = Create(); f.Project.VBComponents[0].Type = 1; f.Module.Code = "Sub TryMe()\r\nEnd Sub";
            var r = new Request { Project = f.Project.Name, Module = "Module1", Procedure = "TryMe", ExpectedMode = 2, ExpectedSha256 = Sha(f.Module.Code) };
            Assert.ThrowsException<ArgumentException>(() => f.Service.RunSub(null));
            foreach (var name in new[] { null, " ", "1Bad", new string('A', 41) })
            { r.Procedure = name; Assert.ThrowsException<ArgumentException>(() => f.Service.RunSub(r)); }
            r.Procedure = "TryMe";
            r.ExpectedMode = 1; Assert.ThrowsException<ArgumentException>(() => f.Service.RunSub(r)); r.ExpectedMode = 2;
            f.Project.Mode = 1; Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(r)); f.Project.Mode = 2;
            r.Module = "Missing"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(r)); r.Module = "Module1";
            f.Project.VBComponents[0].Type = 2; Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(r)); f.Project.VBComponents[0].Type = 1;
            r.ExpectedSha256 = " "; Assert.ThrowsException<ArgumentException>(() => f.Service.RunSub(r)); r.ExpectedSha256 = Sha(f.Module.Code);
            f.Module.Code = ""; r.ExpectedSha256 = Sha(""); Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(r));
            f.Module.Code = "Sub TryMe()\r\nEnd Sub"; r.ExpectedSha256 = Sha(f.Module.Code);
            f.Bar.Controls.AddRange(new[] { new FakeControl { Id = -1, Caption = "Run Sub" },
                new FakeControl { Id = 186, Caption = "Run Sub", Enabled = false }, new FakeControl { Id = 186, Caption = "Unrelated" } });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(r));
        }

        [TestMethod]
        public void RunSubSelectsRequestedPaneBeforeCheckingNativeRunAvailability()
        {
            var f = Create(); f.Project.VBComponents[0].Type = 1;
            f.Module.Code = "Sub TryMe()\r\nEnd Sub";
            var request = new Request
            {
                Project = f.Project.Name,
                Module = "Module1",
                Procedure = "TryMe",
                ExpectedMode = 2,
                ExpectedSha256 = Sha(f.Module.Code)
            };
            f.Vbe.ActiveCodePane = new FakePane { CodeModule = f.Module };
            f.Module.CodePane.OnSetSelection = () => Assert.AreSame(f.Module.CodePane, f.Vbe.ActiveCodePane);
            var command = new FakeControl
            {
                Id = 186,
                Caption = "Run Sub",
                OnEnabledRead = () => Assert.AreSame(f.Module.CodePane, f.Vbe.ActiveCodePane)
            };
            f.Bar.Controls.Add(command);
            int assignmentsBefore = f.Vbe.ActiveCodePaneSetCount;

            dynamic result = f.Service.RunSub(request);

            Assert.IsTrue((bool)result.Executed);
            Assert.AreEqual(1, command.ExecuteCount);
            Assert.IsTrue(f.Vbe.ActiveCodePaneSetCount >= assignmentsBefore + 2);
            Assert.AreSame(f.Module.CodePane, f.Vbe.ActiveCodePane);
        }

        [TestMethod]
        public void RunSubDoesNotExecuteWhenTheHostIgnoresPaneSelection()
        {
            var f = Create(); f.Project.VBComponents[0].Type = 1;
            f.Module.Code = "Sub TryMe()\r\nEnd Sub";
            var request = new Request
            {
                Project = f.Project.Name,
                Module = "Module1",
                Procedure = "TryMe",
                ExpectedMode = 2,
                ExpectedSha256 = Sha(f.Module.Code)
            };
            var previous = new FakePane { CodeModule = f.Module };
            f.Vbe.ActiveCodePane = previous;
            f.Vbe.IgnoreActiveCodePaneAssignment = true;
            var command = new FakeControl { Id = 186, Caption = "Run Sub" };
            f.Bar.Controls.Add(command);

            Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunSub(request));

            Assert.AreEqual(0, command.ExecuteCount);
            Assert.AreSame(previous, f.Vbe.ActiveCodePane);
        }

        [TestMethod]
        public void EnumerationCapsDepthAndCountAndSkipsUnavailableBarsButtonsAndChildren()
        {
            var f = Create(); f.Vbe.CommandBars.Add(new FakeBar { FailName = true });
            f.Bar.Controls.AddRange(new[] { new FakeControl { FailCaption = true }, new FakeControl { Caption = "Leaf", Id = 1, FailChildren = true } });
            var parent = new FakeControl { Caption = "Root", Id = 2 }; f.Bar.Controls.Add(parent);
            for (int level = 0; level < 7; level++) { var child = new FakeControl { Caption = "Level" + level, Id = 10 + level }; parent.Controls.Add(child); parent = child; }
            var listed = ((IEnumerable)f.Service.ListCommands(null, 0, 200)).Cast<object>().ToArray();
            Assert.AreEqual(6, listed.Length); Assert.AreEqual("Debug > Leaf", (string)((dynamic)listed[0]).Path);
            f.Bar.Controls.Clear();
            for (int i = 0; i < 2010; i++) f.Bar.Controls.Add(new FakeControl { Caption = "Command" + i, Id = i });
            f.Vbe.CommandBars.Add(new FakeBar { Name = "Late" });
            Assert.AreEqual(200, ((IEnumerable)f.Service.ListCommands(null, 0, 999)).Cast<object>().Count());
            Assert.AreEqual(1, ((IEnumerable)f.Service.ListCommands(null, 1999, 10)).Cast<object>().Count());
            Assert.AreEqual(0, ((IEnumerable)f.Service.ListCommands(null, 2000, 10)).Cast<object>().Count());
        }

        [TestMethod]
        public void BrowserReadbackAndComIdentityHelpersHandleMissingPropertiesAndAcquireFailures()
        {
            var browser = typeof(VbeDebug).GetMethod("HasVisibleObjectBrowser", BindingFlags.NonPublic | BindingFlags.Static);
            foreach (var properties in new[] { new Dictionary<string, object>(), new Dictionary<string, object> { ["Type"] = 1 },
                new Dictionary<string, object> { ["Type"] = 2 }, new Dictionary<string, object> { ["Type"] = 2, ["Visible"] = false },
                new Dictionary<string, object> { ["Type"] = 2, ["Visible"] = true } })
            {
                var snapshot = new BrowserSnapshot { Windows = new[] { new BrowserWindowSnapshot { Properties = properties } } };
                Assert.AreEqual(properties.ContainsKey("Visible") && (bool)properties["Visible"], (bool)browser.Invoke(null, new object[] { snapshot }));
            }
            var caption = typeof(VbeDebug).GetMethod("IsObjectBrowserCaption", BindingFlags.NonPublic | BindingFlags.Static);
            foreach (var text in new[] { null, "Other", "Explorateur d'objets", "Explorateur d’objets", "Object Browser" })
                Assert.AreEqual(text != null && text != "Other", (bool)caption.Invoke(null, new object[] { text }));
            var same = typeof(VbeDebug).GetMethod("SameComObject", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.ThrowsException<TargetInvocationException>(() => same.Invoke(null, new object[] { null, new object() }));
            Assert.ThrowsException<TargetInvocationException>(() => same.Invoke(null, new object[] { new object(), null }));
            var f = Create(); f.Project.FailFileName = true;
            dynamic state = f.Service.State(f.Project.Name); Assert.IsNull((string)state.SelectedProjectPath); Assert.IsNotNull((object)state.Selection);
        }

        [TestMethod]
        public void NativeBrowserPaneAndDeleteWatchSearchSkipEveryNonMatchingCandidate()
        {
            var f = Create(1);
            var windows = new VbeEditorWindows(new VbeEditorWindowsTests.FakeVbe());
            f.Bar.Controls.AddRange(new[] { new FakeControl { Id = -1, Caption = "Object Browser" },
                new FakeControl { Id = 473, Caption = "Object Browser", Enabled = false } });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenObjectBrowser(windows));
            f.Bar.Controls.Add(new FakeControl { Id = 473, Caption = "Object Browser" });
            f.Service.OpenObjectBrowser(windows);
            f.Bar.Controls.Add(new FakeControl { Id = 2555, Caption = null });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenDebugPane("locals", windows));
            f.Bar.Controls.Add(new FakeControl { Id = 2555, Caption = "Variables locales" });
            f.Service.OpenDebugPane("locals", windows);
            foreach (var caption in new[] { "Delete Watch", "Supprimer un espion" })
            {
                f.Bar.Controls.Clear();
                f.Bar.Controls.AddRange(new[] { new FakeControl { Id = -1, Caption = caption },
                    new FakeControl { Id = 1083, Caption = caption, Enabled = false },
                    new FakeControl { Id = 1083, Caption = null }, new FakeControl { Id = 1083, Caption = "Other" } });
                var request = new Request { Project = f.Project.Name, Expression = "x", Context = "Module1", ExpectedMode = 1 };
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RemoveSelectedWatch(request));
                f.Bar.Controls.Add(new FakeControl { Id = 1083, Caption = caption });
                f.Service.RemoveSelectedWatch(request); Assert.AreEqual(1, f.Bar.Controls.Last().ExecuteCount);
            }
        }

        [TestMethod]
        public void DefaultCodeSelectionUsesAnEmptyCaretAndCommandSearchChecksBothIdAndCaption()
        {
            var f = Create(); var request = Location(f);
            dynamic selected = f.Service.SelectCode(request);
            Assert.AreEqual("", (string)selected.SelectedText);
            Assert.AreEqual(1, (int)selected.StartColumn); Assert.AreEqual(1, (int)selected.EndColumn);
            request.Action = "toggle_breakpoint"; request.ControlId = 123; request.ControlCaption = "Toggle Breakpoint";
            f.Bar.Controls.AddRange(new[] { new FakeControl { Id = 999, Caption = request.ControlCaption },
                new FakeControl { Id = request.ControlId, Caption = "Other" }, new FakeControl { Id = request.ControlId, Caption = request.ControlCaption } });
            f.Service.InvokeCommand(request); Assert.AreEqual(1, f.Bar.Controls.Last().ExecuteCount);
        }
    }
}
