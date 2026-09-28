using System;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Threading;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, DoNotParallelize]
    [TestCategory("Unit")]
    public sealed class EditorSessionDispatchScenarios
    {
        private static Response Run(VbeSession session, Request request, string command)
        { request.Command = command; var result = session.Execute(request); Assert.IsTrue(result.Ok, command + ": " + result.Error); return result; }
        private static Request WindowRequest(VbeSession session, SessionDispatchFixture.Window window, string action)
        {
            var r = new Request { WindowCaption = window.Caption, WindowType = window.Type, Action = action };
            r.ExpectedWindowVersion = (string)((dynamic)Run(session, r, "window_layout").Data).WindowVersion;
            return r;
        }

        [TestMethod]
        public void SessionRoutesWindowToolbarAndAddInActionsThroughTheirRealServices()
        {
            var host = new SessionDispatchFixture.Host(); var pane = new SessionDispatchFixture.Window();
            var frame = new SessionDispatchFixture.Window { Caption = "Frame", Type = 11 };
            host.Windows.Add(pane); host.Windows.Add(frame); pane.Focus = () => host.ActiveWindow = pane;
            var bar = new VbeToolbarCoverageTests.NativeBar(); host.CommandBars.Add(bar);
            var plugin = new VbeAddInReadbackTests.Plugin(); host.AddIns.Add(plugin);
            var session = new VbeSession(host);
            Assert.IsFalse(session.Execute(new Request()).Ok); Assert.IsFalse(session.Execute(new Request { Command = " " }).Ok);
            Run(session, WindowRequest(session, pane, "maximize"), "set_window_state");
            Run(session, WindowRequest(session, pane, "restore"), "set_window_state");
            var bounds = WindowRequest(session, pane, null); bounds.Left = 10; bounds.Top = 20; bounds.Width = 400; bounds.Height = 300;
            Run(session, bounds, "set_window_bounds"); Assert.AreEqual(400, pane.Width);
            Run(session, WindowRequest(session, pane, null), "show_vbe_window"); Assert.AreSame(pane, host.ActiveWindow);
            var link = WindowRequest(session, pane, "link"); link.TargetWindowCaption = frame.Caption; link.TargetWindowType = frame.Type;
            link.ExpectedTargetWindowVersion = (string)((dynamic)Run(session, WindowRequest(session, frame, null), "window_layout").Data).WindowVersion;
            Assert.IsTrue((bool)((dynamic)Run(session, link, "link_vbe_window").Data).Verified);
            Run(session, WindowRequest(session, pane, "unlink"), "link_vbe_window"); Assert.IsNull(pane.LinkedWindowFrame);
            foreach (string command in new[] { "set_toolbar_visibility", "set_toolbar_position", "set_toolbar_placement" })
            {
                dynamic state = ((dynamic)Run(session, new Request(), "list_toolbars").Data).Toolbars[0];
                var r = new Request { ObjectName = "Standard", ExpectedWindowVersion = state.WindowVersion, ExpectedToolbarLayoutVersion = state.ToolbarLayoutVersion,
                    Action = command == "set_toolbar_visibility" ? "hide" : "float", ToolbarLeft = 20, ToolbarTop = 30 };
                Assert.IsTrue((bool)((dynamic)Run(session, r, command).Data).Verified);
            }
            dynamic addin = ((dynamic)Run(session, new Request(), "list_addins").Data).AddIns[0];
            Run(session, new Request { Action = "connect", ProgId = plugin.ProgId, ExpectedAddInVersion = addin.AddInVersion }, "set_addin_connection");
            Assert.IsTrue(plugin.Connect);
        }

        [TestMethod]
        public void SessionRoutesCodeEditsUndoRedoAndClipboardThroughActualModuleMutations()
        {
            var host = new VbeSessionTests.FakeVbe(); var project = new VbeSessionTests.FakeProject { Name = "P", Mode = 2, FileName = "" };
            var module = new VbeSessionTests.FakeModule("abc"); project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "M", Type = 1, CodeModule = module });
            host.VBProjects.Add(project); var session = new VbeSession(host);
            Func<Request> current = () => new Request { Project = "P", Module = "M", ExpectedSha256 = VbeCodeClipboard.Hash(module.Code), Action = "comment", StartLine = 1, Count = 1 };
            Run(session, current(), "preview_code_edit"); Assert.AreEqual("abc", module.Code);
            Run(session, current(), "apply_code_edit"); Assert.AreNotEqual("abc", module.Code);
            Run(session, current(), "undo_code_edit"); Assert.AreEqual("abc", module.Code);
            Run(session, current(), "redo_code_edit"); Assert.AreNotEqual("abc", module.Code);
            var clipboard = new SessionDispatchFixture.Clipboard();
            typeof(VbeSession).GetField("codeClipboard", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(session, new VbeCodeClipboard(session.Execute, clipboard));
            Run(session, new Request(), "read_code_clipboard");
            foreach (string command in new[] { "copy_code", "cut_code", "paste_code" })
            {
                var request = current(); request.StartColumn = 1; request.EndLine = 1; request.EndColumn = command == "paste_code" ? 1 : 2;
                request.ExpectedClipboardVersion = clipboard.State.Version;
                Run(session, request, command);
            }
            Run(session, new Request { Project = "P" }, "project_symbols");
            Run(session, new Request { Project = "P", Action = "list" }, "code_bookmark");
            var sources = new SessionDispatchFixture.Host(); sources.VBProjects.Add(project);
            session = new VbeSession(sources);
            Assert.AreSame(sources.VBProjects, session.ProjectsEventSource());
            Assert.IsNull(session.ComponentsEventSource(null)); Assert.IsNull(session.ComponentsEventSource(""));
            Assert.AreSame(project.VBComponents, session.ComponentsEventSource("P"));
            Assert.IsNull(session.ReferenceEventSource(null)); Assert.AreSame(project, session.ReferenceEventSource("P"));
        }

        [STATestMethod]
        public void SessionRoutesDebuggerLayoutsViewsNativeHistoryAndQueuedNavigation()
        {
            var f = new EditorDebugFixture(); var session = new VbeSession(f.Vbe);
            var request = f.Location("module");
            dynamic row = ((IEnumerable)((dynamic)Run(session, request, "code_pane_layout").Data).Panes).Cast<object>().Single();
            request.Pane = row.Pane; request.ExpectedWindowVersion = row.WindowVersion;
            Run(session, request, "set_code_view");
            request = f.Location(); row = ((IEnumerable)((dynamic)Run(session, request, "code_pane_layout").Data).Panes).Cast<object>().Single();
            request.Pane = row.Pane; request.ExpectedWindowVersion = row.WindowVersion;
            Run(session, request, "scroll_code_pane");
            request = f.Location("unsplit"); Run(session, request, "set_code_split");
            Run(session, new Request(), "editor_layout");
            var other = f.AddModule("Other");
            f.Command(1826, "cascade", () => { other.CodePane.Window.Left = 26; other.CodePane.Window.Top = 26; });
            request = f.Location("cascade"); request.ExpectedWindowVersion = ((dynamic)Run(session, new Request(), "editor_layout").Data).WindowVersion;
            Run(session, request, "arrange_editor_windows");
            var savedEnabled = VbeDebug.HistoryWindowEnabled;
            var savedContext = SynchronizationContext.Current;
            try
            {
                VbeDebug.HistoryWindowEnabled = handle => true;
                f.Command(128, "undo"); request = f.Location("undo");
                request.ExpectedProjectVersion = ((dynamic)Run(session, request, "native_code_history_state").Data).HistoryVersion;
                Run(session, request, "native_code_history");
                var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
                var command = f.Command(1822, "last_position", () => f.Pane.Start = f.Pane.End = 2);
                request = f.Location("last_position"); request.StartColumn = request.EndColumn = 1;
                dynamic operation = Run(session, request, "native_code_navigation").Data;
                context.RunAll(); Assert.AreEqual(1, command.Executions);
                Run(session, new Request { Action = "status", Query = operation.OperationId }, "native_code_navigation");
            }
            finally { VbeDebug.HistoryWindowEnabled = savedEnabled; SynchronizationContext.SetSynchronizationContext(savedContext); }
            request = f.Location("go"); request.StartColumn = 1;
            Run(session, request, "navigate_code");
        }

        [TestMethod]
        public void SessionReportsRollbackReadbackMismatchInsteadOfClaimingRestoration()
        {
            var host = new SessionDispatchFixture.Host(); var project = new SessionDispatchFixture.CorruptProject();
            var component = new SessionDispatchFixture.CorruptComponent(); project.VBComponents.Add(component); host.VBProjects.Add(project);
            var session = new VbeSession(host);
            var result = session.Execute(new Request { Command = "replace_lines", Project = "P", Module = "M", StartLine = 1, Count = 1, Text = "replacement", ExpectedSha256 = VbeCodeClipboard.Hash("original") });
            Assert.IsFalse(result.Ok); StringAssert.Contains(result.Error, "Restored source does not match");
            Assert.AreEqual("unexpected normalization", component.CodeModule.Code);
        }
    }
}
