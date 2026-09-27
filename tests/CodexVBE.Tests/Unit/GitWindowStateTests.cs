using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class GitWindowStateTests
    {
        private const BindingFlags InstancePrivate = BindingFlags.Instance | BindingFlags.NonPublic;

        private static T Field<T>(GitWindow window, string name)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(window);
        }

        private static void Set(GitWindow window, string name, object value)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            field.SetValue(window, value);
        }

        private static object Invoke(GitWindow window, string name, params object[] args)
        {
            var method = typeof(GitWindow).GetMethod(name, InstancePrivate);
            Assert.IsNotNull(method, name);
            return method.Invoke(window, args);
        }

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
                ((Task)Invoke(window, "Perform", new Func<Task>(() => { ran = true; return Task.CompletedTask; })))
                    .GetAwaiter().GetResult();
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
        public void DiffRowsUseOneBasedLineNumbersAndMarkOnlyChangedSides()
        {
            using (var window = new GitWindow())
            {
                Invoke(window, "AddDiffRow", "before", "after", 0, 3, false);
                Invoke(window, "AddDiffRow", "removed", null, 1, 0, true);
                Invoke(window, "AddDiffRow", null, "added", 0, 4, true);
                var grid = Field<DataGridView>(window, "diff");
                Assert.AreEqual("1  before", grid.Rows[0].Cells[0].Value);
                Assert.AreEqual("4  after", grid.Rows[0].Cells[1].Value);
                Assert.AreEqual(Color.MistyRose, grid.Rows[1].Cells[0].Style.BackColor);
                Assert.AreEqual(Color.Empty, grid.Rows[1].Cells[1].Style.BackColor);
                Assert.AreEqual(Color.Empty, grid.Rows[2].Cells[0].Style.BackColor);
                Assert.AreEqual(Color.Honeydew, grid.Rows[2].Cells[1].Style.BackColor);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void SelectingChangeWithoutSnapshotClearsPreviousPreview()
        {
            using (var window = new GitWindow())
            {
                var grid = Field<DataGridView>(window, "diff");
                grid.Rows.Add("old", "new");
                Invoke(window, "Changes_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(0, grid.Rows.Count);
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
                var success = (Task)Invoke(window, "Perform", new Func<Task>(() => {
                    calls++;
                    Assert.IsFalse(Field<Button>(window, "connect").Enabled);
                    Assert.IsFalse(Field<TextBox>(window, "commitMessage").Enabled);
                    return Task.CompletedTask;
                }));
                success.GetAwaiter().GetResult();
                Assert.AreEqual(1, calls);
                Assert.IsTrue(Field<Button>(window, "connect").Enabled);
                Assert.IsTrue(Field<TextBox>(window, "commitMessage").Enabled);

                var failure = (Task)Invoke(window, "Perform", new Func<Task>(() => {
                    throw new InvalidOperationException("disposable Git failure");
                }));
                failure.GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "status").Text, "disposable Git failure");
                Assert.IsTrue(Field<Button>(window, "connect").Enabled);
                Set(window, "running", true);
                ((Task)Invoke(window, "Perform", new Func<Task>(() => {
                    calls++;
                    return Task.CompletedTask;
                }))).GetAwaiter().GetResult();
                Assert.AreEqual(1, calls);
                Set(window, "running", false);
            }
        }
    }
}
