namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeProjectProtectionTests
    {
        [TestMethod]
        public void QueueProjectPropertiesRequiresExactSelectionAndRechecksEveryGuardOnUiCallback()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                for (int fault = 0; fault < 6; fault++)
                {
                    var project = new Project(); var host = new Vbe { ActiveVBProject = project }; host.VBProjects.Add(project);
                    var command = new VbeDebugTests.FakeControl { Id = 2578, Caption = "Project Properties..." };
                    var bar = new VbeDebugTests.FakeBar { Name = "Tools" }; bar.Controls.Add(command); host.CommandBars.Add(bar);
                    var service = new VbeDebug(host); var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(context);
                    bool valid = true;
                    var request = new Request { Project = project.FileName, ExpectedProjectVersion = "version", ControlCaption = command.Caption };
                    dynamic result = service.QueueProjectPropertiesDialog(request, _ => valid);
                    Assert.AreEqual(project.Name, (string)result.ProjectName);
                    if (fault == 1) project.Mode = 1;
                    if (fault == 2) project.Protection = 1;
                    if (fault == 3) host.ActiveVBProject = new Project();
                    if (fault == 4) valid = false;
                    if (fault == 5) command.Caption = "Other";
                    context.Run();
                    Assert.AreEqual(fault == 0 ? 1 : 0, command.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void QueueRejectsIncompleteRequestsUnavailableCommandAndUnsafeProjectBeforePosting()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                var project = new Project(); var host = new Vbe { ActiveVBProject = project }; host.VBProjects.Add(project);
                var service = new VbeDebug(host); var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(context);
                Assert.ThrowsException<ArgumentException>(() => service.QueueProjectPropertiesDialog(null, _ => true));
                var request = new Request { Project = project.Name, ExpectedProjectVersion = "version", ControlCaption = "Exact" };
                Assert.ThrowsException<InvalidOperationException>(() => service.QueueProjectPropertiesDialog(request, _ => true));
                project.Mode = 1;
                Assert.ThrowsException<InvalidOperationException>(() => service.QueueProjectPropertiesDialog(request, _ => true));
                Assert.IsNull(context.Callback);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }
}
