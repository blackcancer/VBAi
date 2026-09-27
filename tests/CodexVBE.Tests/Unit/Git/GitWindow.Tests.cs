namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.Serialization;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class GitWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void UnboundWindowRejectsActionsAndKeepsRepositoryControlsDisabled()
        {
            using (var window = new GitWindow())
            {
                Invoke(window, "UpdateButtons");
                Assert.IsTrue(Field<Button>(window, "connect").Enabled);
                Assert.IsFalse(Field<Button>(window, "compare").Enabled);
                Assert.IsFalse(Field<Button>(window, "push").Enabled);
                Assert.IsFalse(Field<TabPage>(window, "branchesTab").Enabled);
                var ran = false;
                ((Task)Invoke(window, "Perform", new Func<Task>(() =>
                {
                    ran = true;
                    return Task.CompletedTask;
                }))).GetAwaiter().GetResult();
                Assert.IsFalse(ran);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void RunningWindowRejectsCloseAndDisablesInteractiveCommands()
        {
            using (var window = new GitWindow())
            {
                Set(window, "running", true);
                Invoke(window, "UpdateButtons");
                Assert.IsFalse(Field<Button>(window, "connect").Enabled);
                Assert.IsFalse(Field<Button>(window, "compare").Enabled);
                Assert.IsFalse(Field<TextBox>(window, "commitMessage").Enabled);
                Assert.IsTrue(Field<TextBox>(window, "remote").ReadOnly);
                var e = new FormClosingEventArgs(CloseReason.UserClosing, false);
                Invoke(window, "OnFormClosing", e);
                Assert.IsTrue(e.Cancel);
                Set(window, "running", false);
                Invoke(window, "UpdateButtons");
                Assert.IsFalse(Field<TextBox>(window, "remote").ReadOnly);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void PerformRestoresControlsAfterSuccessAndReportsActionFailure()
        {
            using (var window = new GitWindow())
            {
                Set(window, "project", (VbaGitProject)FormatterServices.GetUninitializedObject(typeof(VbaGitProject)));
                int calls = 0;
                var success = (Task)Invoke(window, "Perform", new Func<Task>(() =>
                {
                    calls++;
                    Assert.IsFalse(Field<Button>(window, "connect").Enabled);
                    Assert.IsFalse(Field<TextBox>(window, "commitMessage").Enabled);
                    return Task.CompletedTask;
                }));
                success.GetAwaiter().GetResult();
                Assert.AreEqual(1, calls);
                Assert.IsTrue(Field<Button>(window, "connect").Enabled);
                Assert.IsTrue(Field<TextBox>(window, "commitMessage").Enabled);
                var failure = (Task)Invoke(window, "Perform", new Func<Task>(() =>
                {
                    throw new InvalidOperationException("disposable Git failure");
                }));
                failure.GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "status").Text, "disposable Git failure");
                Assert.IsTrue(Field<Button>(window, "connect").Enabled);
                Set(window, "running", true);
                ((Task)Invoke(window, "Perform", new Func<Task>(() =>
                {
                    calls++;
                    return Task.CompletedTask;
                }))).GetAwaiter().GetResult();
                Assert.AreEqual(1, calls);
                Set(window, "running", false);
            }
        }
    }
}
