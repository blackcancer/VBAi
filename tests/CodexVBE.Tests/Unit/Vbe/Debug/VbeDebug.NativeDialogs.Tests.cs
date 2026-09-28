namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugTests
    {
        [TestMethod]
        public void NativeDialogPreflightRejectsEveryMissingIdentityModeAndPrintBinding()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(new RecordingContext());
                var stopped = Create(1); var stoppedContext = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(stoppedContext);
                var stoppedCommand = new FakeControl { Id = 930, Caption = "Macros" }; stopped.Bar.Controls.Add(stoppedCommand);
                var stoppedRequest = Location(stopped); stoppedRequest.Action = "macros"; stoppedRequest.ControlCaption = "Macros";
                stopped.Service.QueueNativeIdeDialog(stoppedRequest); stoppedContext.RunAll(); Assert.AreEqual(1, stoppedCommand.ExecuteCount);
                foreach (string fault in new[] { "project", "caption", "active project", "mode mismatch", "invalid mode", "sha empty", "sha stale", "no pane", "wrong pane", "disabled", "caption mismatch" })
                {
                    var f = Create(); var request = Location(f); request.Action = "print"; request.ControlCaption = "Print"; var control = new FakeControl { Id = 4, Caption = "Print" }; f.Bar.Controls.Add(control);
                    if (fault == "project") request.Project = "";
                    if (fault == "caption") request.ControlCaption = null;
                    if (fault == "active project") f.Vbe.ActiveVBProject = new FakeProject { Name = "Other" };
                    if (fault == "mode mismatch") request.ExpectedMode = 1;
                    if (fault == "invalid mode") { request.ExpectedMode = 0; f.Project.Mode = 0; }
                    if (fault == "sha empty") request.ExpectedSha256 = null;
                    if (fault == "sha stale") request.ExpectedSha256 = "stale";
                    if (fault == "no pane") f.Vbe.ActiveCodePane = null;
                    if (fault == "wrong pane") f.Vbe.ActiveCodePane = Create().Vbe.ActiveCodePane;
                    if (fault == "disabled") control.Enabled = false;
                    if (fault == "caption mismatch") control.Caption = "Other";
                    if (fault == "project" || fault == "caption") Assert.ThrowsException<ArgumentException>(() => f.Service.QueueNativeIdeDialog(request), fault);
                    else Assert.ThrowsException<InvalidOperationException>(() => f.Service.QueueNativeIdeDialog(request), fault);
                    Assert.AreEqual(0, control.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void NativeDialogEmptyPrintSourceAndLateProjectPaneChangesAreGuarded()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (string fault in new[] { "none", "replaced project", "wrong pane" })
                {
                    var f = Create(); f.Module.Code = ""; var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
                    var control = new FakeControl { Id = 4, Caption = "Print" }; f.Bar.Controls.Add(control);
                    var request = Location(f); request.Action = "print"; request.ExpectedSha256 = Sha(""); request.ControlCaption = "Print";
                    f.Service.QueueNativeIdeDialog(request);
                    if (fault == "replaced project") { var replacement = Create().Project; f.Vbe.VBProjects.Clear(); f.Vbe.VBProjects.Add(replacement); f.Vbe.ActiveVBProject = replacement; }
                    if (fault == "wrong pane") f.Vbe.ActiveCodePane = Create().Vbe.ActiveCodePane;
                    context.RunAll(); Assert.AreEqual(fault == "none" ? 1 : 0, control.ExecuteCount, fault);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        /// <summary>Les surfaces reconnues sont programmées sans valider un dialogue ni imprimer.</summary>
        [TestMethod]
        public void NativeIdeDialogsQueueOnlyRecognizedCommandsAndCaptureTheirContext()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (string action in new[] { "macros", "help", "print", "additional_controls" })
                {
                    var f = Create(); var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
                    int id = action == "macros" ? 930 : action == "help" ? 984 : action == "print" ? 4 : 642;
                    var control = new FakeControl { Id = id, Caption = "Exact command" }; f.Bar.Controls.Add(control);
                    var request = Location(f); request.Action = action; request.ControlCaption = control.Caption;
                    dynamic result = f.Service.QueueNativeIdeDialog(request);
                    Assert.IsTrue((bool)result.Scheduled); Assert.IsFalse((bool)result.PrintSubmitted); Assert.AreEqual(0, control.ExecuteCount);
                    request.Action = "different"; request.Project = "different"; request.ControlCaption = "different";
                    context.RunAll(); Assert.AreEqual(1, control.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Les versions de source, états actifs et commandes sont revérifiés avant une action différée.</summary>
        [TestMethod]
        public void NativeIdeDialogsRefuseInvalidAndChangedSourceOrActiveProject()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (int fault in System.Linq.Enumerable.Range(0, 8))
                {
                    var f = Create(); var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
                    var control = new FakeControl { Id = 4, Caption = "Print..." }; f.Bar.Controls.Add(control);
                    var request = Location(f); request.Action = "print"; request.ControlCaption = control.Caption;
                    f.Service.QueueNativeIdeDialog(request);
                    if (fault == 0) f.Project.Mode = 1;
                    if (fault == 1) f.Vbe.ActiveVBProject = new FakeProject { Name = "Other" };
                    if (fault == 2) f.Vbe.ActiveCodePane = null;
                    if (fault == 3) f.Module.Code = "' external edit\r\n" + f.Module.Code;
                    if (fault == 4) control.Caption = "Different";
                    if (fault == 5) control.Enabled = false;
                    if (fault == 6) f.Bar.Controls.Clear();
                    if (fault == 7) control.OnExecute = () => { throw new InvalidOperationException("native failed"); };
                    context.RunAll(); Assert.AreEqual(fault == 7 ? 1 : 0, control.ExecuteCount);
                }
                var fixture = Create(); var r = Location(fixture); r.Action = "unknown"; r.ControlCaption = "Exact";
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.QueueNativeIdeDialog(r));
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.QueueNativeIdeDialog(null));
                r.Action = "macros"; Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.QueueNativeIdeDialog(r));
                fixture.Bar.Controls.Add(new FakeControl { Id = 930, Caption = "Exact" });
                SynchronizationContext.SetSynchronizationContext(null);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.QueueNativeIdeDialog(r));
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }

    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void ExistingProjectPropertiesDialogsAreRejectedOnlyForVisibleCurrentHost()
        {
            foreach (string fault in new[] { "none", "English", "French", "foreign", "hidden", "wrong class", "wrong caption" })
            using (var scene = new SystemScene())
            {
                var window = scene.Add(fault == "French" ? "P - Propriétés du projet" : "P - Project Properties");
                if (fault == "none") scene.Windows.Clear();
                if (fault == "foreign") window.ProcessId = 999999;
                if (fault == "hidden") window.Visible = false;
                if (fault == "wrong class") window.Class = "unrelated";
                if (fault == "wrong caption") window.Text = "Other";
                if (fault == "English" || fault == "French") Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoProjectPropertiesDialog());
                else VbeDebugWindows.EnsureNoProjectPropertiesDialog();
            }
        }
    }
}
