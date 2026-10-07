using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeFormRunWorkflowTests
    {
        private SynchronizationContext saved; private NativeNavigationContext context;
        [TestInitialize] public void IsolateContext() { saved = SynchronizationContext.Current; context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context); }
        [TestCleanup] public void RestoreContext() { SynchronizationContext.SetSynchronizationContext(saved); }
        [TestMethod]
        public void FormRunRequiresEveryPreflightFieldAndAnAvailableDispatcher()
        {
            using (var f = new FormsWorkflowFixture())
            {
                foreach (string field in new[] { "ExpectedMode", "Project", "Form", "ExpectedSha256", "ExpectedTreeVersion", "ControlCaption" })
                {
                    var r = f.Request(); typeof(Request).GetProperty(field).SetValue(r, field == "ExpectedMode" ? (object)0 : null);
                    Assert.ThrowsException<ArgumentException>(() => f.Service.RunForm(r), field);
                }
                SynchronizationContext.SetSynchronizationContext(null);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunForm(f.Request()));
            }
        }
        [TestMethod]
        public void FormRunRejectsChangedCodeTreeAndWrongFocusWithoutScheduling()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3 })
                using (var f = new FormsWorkflowFixture())
                {
                    var r = f.Request(); if (scenario == 0) r.ExpectedSha256 = "stale"; if (scenario == 1) r.ExpectedTreeVersion = "stale";
                    if (scenario == 2) f.Form.Window.OnFocus = () => f.Vbe.ActiveVBProject = null;
                    if (scenario == 3) f.Form.Window.OnFocus = () => f.Vbe.ActiveWindow = new object();
                    Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunForm(r)); Assert.AreEqual(0, context.Pending);
                }
        }
        [TestMethod]
        public void FormRunQueuesCapturesArgumentsAndReportsCommandReturnWithoutRuntimeClaims()
        {
            using (var f = new FormsWorkflowFixture())
            {
                f.Form.CodeModule.Reset(""); var r = f.Request(); dynamic queued = f.Service.RunForm(r); string id = queued.OperationId;
                Assert.AreEqual("Queued", (string)queued.State); Assert.IsFalse((bool)queued.CommandCompleted);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunForm(f.Request()));
                r.Project = "changed"; r.Form = "changed"; r.ExpectedSha256 = "changed"; r.ExpectedTreeVersion = "changed"; r.ControlCaption = "changed";
                var command = (FormsWorkflowFixture.RunControl)f.Vbe.CommandBars.Control;
                command.OnExecute = () =>
                {
                    dynamic running = f.Service.FormRunStatus(new Request { Project = f.Project.Name, Query = id }); Assert.AreEqual("Running", (string)running.State);
                    Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunForm(f.Request()));
                };
                context.RunAll(); dynamic result = f.Service.FormRunStatus(new Request { Project = f.Project.Name, Query = id });
                Assert.AreEqual("CommandReturned", (string)result.State); Assert.IsTrue((bool)result.CommandCompleted); Assert.IsFalse((bool)result.RuntimeVerified);
                Assert.ThrowsException<ArgumentException>(() => f.Service.FormRunStatus(new Request { Project = "other", Query = id }));
                Assert.ThrowsException<ArgumentException>(() => f.Service.FormRunStatus(new Request { Project = f.Project.Name, Query = "unknown" }));
            }
        }
        [TestMethod]
        public void FormRunRechecksProjectFormModeRevisionFocusAndNativeCommandBeforeExecute()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8 })
                using (var f = new FormsWorkflowFixture())
                {
                    dynamic queued = f.Service.RunForm(f.Request()); string id = queued.OperationId;
                    if (scenario == 0) f.Project.Mode = 1;
                    if (scenario == 1) { var project = new FormsWorkflowFixture.WorkflowProject(); project.VBComponents.Add(f.Form); f.Vbe.VBProjects[0] = project; }
                    if (scenario == 2) f.Project.VBComponents[0] = new FormsWorkflowFixture.WorkflowForm(f);
                    if (scenario == 3) f.Form.CodeModule.Reset("changed");
                    if (scenario == 4) f.Form.Designer.Caption = "changed";
                    if (scenario == 5) f.Form.Window.OnFocus = () => f.Vbe.ActiveWindow = null;
                    if (scenario == 6) f.Form.Window.OnFocus = () => f.Form.Designer.Caption = "changed during focus";
                    if (scenario == 7) f.Vbe.CommandBars.Control = null;
                    if (scenario == 8) ((FormsWorkflowFixture.RunControl)f.Vbe.CommandBars.Control).OnExecute = () => { throw new InvalidOperationException("run rejected"); };
                    context.RunAll(); dynamic result = f.Service.FormRunStatus(new Request { Project = f.Project.Name, Query = id });
                    Assert.AreEqual("Failed", (string)result.State); Assert.IsNotNull(result.Error); Assert.IsFalse((bool)result.CommandCompleted);
                }
        }
        [TestMethod]
        public void FormRunRemovesRejectedPostsAndRetainsTheLatestTwentyOperations()
        {
            using (var f = new FormsWorkflowFixture())
            {
                context.RejectPost = 1; Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunForm(f.Request()));
                context.RejectPost = 0; string first = null;
                for (int i = 0; i < 21; i++) { dynamic result = f.Service.RunForm(f.Request()); if (i == 0) first = result.OperationId; context.RunAll(); }
                Assert.ThrowsException<ArgumentException>(() => f.Service.FormRunStatus(new Request { Project = f.Project.Name, Query = first }));
                Assert.AreEqual(21, ((FormsWorkflowFixture.RunControl)f.Vbe.CommandBars.Control).Calls);
            }
        }
        [TestMethod]
        public void FormRunCommandAlsoRejectsNullControlAndBlankCaption()
        {
            using (var f = new FormsWorkflowFixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RequireFormRunCommand(" "));
                f.Vbe.CommandBars.Control = null; Assert.ThrowsException<InvalidOperationException>(() => f.Service.RequireFormRunCommand("Run"));
            }
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed class VbeFormRunCommandTests
    {
        public sealed class Command { public int Id = 186; public bool Enabled = true; public string Caption; }
        public sealed class Bars
        {
            public Command Command = new Command();
            public object FindControl(int type, int id) { Assert.AreEqual(1, type); Assert.AreEqual(186, id); return Command; }
        }
        public sealed class Host { public Bars CommandBars = new Bars(); }
        [DataTestMethod]
        [DataRow("&Ausführen")]
        [DataRow("&Ejecutar")]
        [DataRow("実行")]
        public void ExactNativeIdentityAcceptsOtherLanguages(string caption)
        {
            var host = new Host(); host.CommandBars.Command.Caption = caption;
            Assert.AreSame(host.CommandBars.Command, new VbeForms(host).RequireFormRunCommand(caption));
        }
        [TestMethod]
        public void ChangedCaptionDisabledOrWrongNativeCommandIsRefused()
        {
            var host = new Host(); host.CommandBars.Command.Caption = "Run";
            var forms = new VbeForms(host);
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Continue"));
            host.CommandBars.Command.Enabled = false;
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Run"));
            host.CommandBars.Command.Enabled = true; host.CommandBars.Command.Id = 999;
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Run"));
        }
    }
}
