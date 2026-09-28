namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;

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

        [TestMethod]
        public void DialogRequestsRejectEveryMissingFieldAndUiContextBeforePosting()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (string field in new[] { "Project", "ExpectedProjectVersion", "ControlCaption", "versionCheck", "context" })
                {
                    var project = new Project(); var host = new DialogHost { ActiveVBProject = project }; host.VBProjects.Add(project);
                    var command = new VbeDebugTests.FakeControl { Id = 2578, Caption = "Exact" }; var bar = new DialogBar(); bar.Controls.Add(command); host.CommandBars.Add(bar);
                    var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(field == "context" ? null : context);
                    var request = new Request { Project = project.Name, ExpectedProjectVersion = "version", ControlCaption = "Exact" };
                    if (field != "versionCheck" && field != "context") typeof(Request).GetProperty(field).SetValue(request, " ");
                    var service = new VbeDebug(host);
                    if (field == "context") StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => service.QueueProjectPropertiesDialog(request, _ => true)).Message, "UI context");
                    else Assert.ThrowsException<ArgumentException>(() => service.QueueProjectPropertiesDialog(request, field == "versionCheck" ? (Func<Request, bool>)null : _ => true));
                    Assert.IsNull(context.Callback); Assert.AreEqual(0, command.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void PostedProjectAndCommandIdentityChangesNeverExecuteAnOwnedCommand()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (string fault in new[] { "project", "rename", "id", "disabled", "caption", "path", "duplicate" })
                {
                    var project = new Project(); var host = new DialogHost { ActiveVBProject = project }; host.VBProjects.Add(project);
                    var command = new VbeDebugTests.FakeControl { Id = 2578, Caption = "Exact" }; var bar = new DialogBar(); bar.Controls.Add(command); host.CommandBars.Add(bar);
                    var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(context);
                    var service = new VbeDebug(host); service.QueueProjectPropertiesDialog(new Request { Project = project.FileName, ExpectedProjectVersion = "v", ControlCaption = "Exact" }, _ => true);
                    if (fault == "project") host.VBProjects[0] = new Project();
                    if (fault == "rename") project.Name = "Renamed";
                    if (fault == "id") command.Id = 1;
                    if (fault == "disabled") command.Enabled = false;
                    if (fault == "caption") command.Caption = "Other";
                    if (fault == "path") bar.Name = "Other";
                    if (fault == "duplicate") bar.Controls.Add(new VbeDebugTests.FakeControl { Id = 2578, Caption = "Exact" });
                    context.Run(); Assert.AreEqual(0, command.ExecuteCount, fault);
                }
                foreach (string fault in new[] { "id", "disabled", "caption" })
                {
                    var project = new Project(); var host = new DialogHost { ActiveVBProject = project }; host.VBProjects.Add(project);
                    var command = new VbeDebugTests.FakeControl { Id = fault == "id" ? 1 : 2578, Caption = fault == "caption" ? "Other" : "Exact", Enabled = fault != "disabled" };
                    var bar = new DialogBar(); bar.Controls.Add(command); host.CommandBars.Add(bar);
                    var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(context);
                    Assert.ThrowsException<InvalidOperationException>(() => new VbeDebug(host).QueueProjectPropertiesDialog(new Request { Project = project.Name, ExpectedProjectVersion = "v", ControlCaption = "Exact" }, _ => true));
                    Assert.IsNull(context.Callback); Assert.AreEqual(0, command.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [STATestMethod]
        public void PostedNativeCommandRequiresRealComIdentityAndAnUnchangedNativeType()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                foreach (string replacement in new[] { "same-type", "different-type", "managed" })
                using (var captured = new OwnedCommandDispatchFixture())
                using (var live = new OwnedCommandDispatchFixture())
                {
                    var project = new Project(); var host = new DialogHost { ActiveVBProject = project }; host.VBProjects.Add(project);
                    var bar = new DialogBar(); bar.Controls.Add(captured.Control); host.CommandBars.Add(bar);
                    var context = new QueueContext(); SynchronizationContext.SetSynchronizationContext(context);
                    new VbeDebug(host).QueueProjectPropertiesDialog(new Request { Project = project.Name, ExpectedProjectVersion = "v", ControlCaption = captured.Caption }, _ => true);
                    var managed = new VbeDebugTests.FakeControl { Id = live.Id, Caption = live.Caption };
                    if (replacement == "different-type") live.Type = 2;
                    bar.Controls[0] = replacement == "managed" ? (object)managed : live.Control;
                    context.Run(); Assert.AreEqual(replacement == "same-type" ? 1 : 0, live.Executions); Assert.AreEqual(0, captured.Executions); Assert.AreEqual(0, managed.ExecuteCount);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

    }
}
