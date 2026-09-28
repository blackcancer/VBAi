namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugTests
    {
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
}
