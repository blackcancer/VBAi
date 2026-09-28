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

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class GitWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void SelectingChangeWithoutSnapshotClearsPreviousPreview()
        {
            using (var window = new GitWindow())
            {
                var view = Field<CodeDiffView>(window, "diff");
                view.ShowDiff("old", "new");
                var grid = DiffGrid(view);
                Assert.IsTrue(grid.RowCount > 0);
                Invoke(window, "Changes_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(0, grid.Rows.Count);
            }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class GitWindowCoverageTests
    {
        [WinFormsTestMethod]
        public void ReviewSnapshotsCoverAddedRemovedUnchangedReferencesAndFormResources()
        {
            using (var f = new Fixture())
            {
                var live = f.Git.Project.Capture();
                var empty = new VbaGitSnapshot(new VbaGitManifest { Components = new VbaGitComponent[0], References = "" }, new Dictionary<string, byte[]>());
                f.Call("PopulateChanges", null, null); Assert.AreEqual(0, f.Get<CheckedListBox>("changes").Items.Count);
                f.Call("PopulateChanges", live, live); Assert.AreEqual(0, f.Get<CheckedListBox>("changes").Items.Count);
                f.Call("PopulateChanges", empty, live); Assert.AreEqual(1, f.Get<CheckedListBox>("changes").Items.Count);
                f.Call("PopulateChanges", live, null); Assert.AreEqual(2, f.Get<CheckedListBox>("changes").Items.Count);
                var form = new VbaGitSnapshot(new VbaGitManifest { Components = new[] { new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } }, References = "form-reference" },
                    new Dictionary<string, byte[]> { ["Form1.frm"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Form1\"\n"), ["Form1.frx"] = new byte[] { 1 } });
                f.Call("PopulateChanges", form, live); Assert.AreEqual(3, f.Get<CheckedListBox>("changes").Items.Count);
                f.Call("PopulateChanges", live, form); Assert.AreEqual(3, f.Get<CheckedListBox>("changes").Items.Count);
                f.Set("reviewCommit", "fixture"); f.Call("PopulateChanges", form, form);
                Assert.AreEqual(1, f.Get<CheckedListBox>("changes").Items.Count);
                var source = typeof(GitWindow).GetMethod("Source", BindingFlags.Static | BindingFlags.NonPublic);
                Assert.AreEqual("", source.Invoke(null, new object[] { null, "Missing" }));
                Assert.AreEqual("", source.Invoke(null, new object[] { live, "Missing" }));
                Assert.AreEqual("", source.Invoke(null, new object[] { live, null }));
                StringAssert.Contains((string)source.Invoke(null, new object[] { live, "Module1" }), "Value = 1");
            }
        }

        [WinFormsTestMethod]
        public void HistoryCheckpointAndModuleRestoreNavigateActualCommitSnapshots()
        {
            using (var f = new Fixture())
            {
                string initial = f.Git.Seed(); string next = f.Git.Commit(f.Git.Snapshot("2"), initial);
                f.Git.Repository.SetRef(f.Git.Repository.Head, next); f.Compare();
                var history = f.Get<ListBox>("history");
                f.Select(history, 0); f.Event("HistoryChanged");
                Assert.IsNotNull(f.Get<string>("reviewCommit"));
                f.Set("running", true); history.SetSelected(1, true); f.Set("running", false);
                f.Event("HistoryChanged", f.Get<Button>("historyCompare"));
                Assert.AreEqual(f.Get<TabPage>("changesTab"), f.Get<TabControl>("tabs").SelectedTab);
                f.Event("RestoreModule_Click");
                Assert.IsTrue(f.Git.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2"));
                f.Git.Repository.Checkpoint(f.Git.Project.Capture(), "Review fixture"); f.Compare();
                f.Select(f.Get<ListBox>("checkpointList"), 0); f.Event("CheckpointChanged");
                StringAssert.Contains(f.Status, UiText.Get("Reviewing checkpoint"));
                f.Event("OpenModule_Click"); Assert.IsFalse(string.IsNullOrWhiteSpace(f.Status));
                f.Set("project", null); f.Event("OpenModule_Click"); f.Set("project", f.Git.Project);
            }
        }

        [WinFormsTestMethod]
        public void PreviewProgressCancellationAndEmptySelectionGuardsMatrix()
        {
            using (var f = new Fixture())
            {
                foreach (string method in new[] { "HistoryChanged", "CheckpointChanged", "RestoreModule_Click", "OpenModule_Click", "CheckpointRestore_Click", "BranchSwitch_Click", "MergeBegin_Click", "MergeOurs_Click", "MergeTheirs_Click", "MergeText_Click" }) f.Event(method);
                f.Event("PreviewImport_Click"); StringAssert.Contains(f.Status, UiText.Get("The target contains no VBA sources."));
                string commit = f.Git.Seed(); f.Git.Repository.Push(commit); f.Event("PreviewImport_Click");
                Assert.AreEqual(UiText.Get("Preview only. Pull imports these changes with a checkpoint."), f.Status);
                f.Call("ReportProgress", "ignored while idle"); Application.DoEvents();
                f.Set("running", true); f.Call("ReportProgress", "fixture progress"); Application.DoEvents();
                Assert.AreEqual("fixture progress", f.Status); f.Set("running", false);
                f.Pump((Task)f.Call("Perform", new Func<Task>(() => { f.Call("CancelOperation_Click", null, EventArgs.Empty); return Task.CompletedTask; }), false));
                f.Event("CancelOperation_Click");
                f.Set("repository", null);
                foreach (string method in new[] { "HistoryChanged", "CheckpointChanged", "PreviewImport_Click", "ConflictList_SelectedIndexChanged" }) f.Event(method);
                f.Set("repository", f.Git.Repository); f.Set("running", true);
                foreach (string method in new[] { "HistoryChanged", "CheckpointChanged", "ConflictList_SelectedIndexChanged" }) f.Call(method, null, EventArgs.Empty);
                f.Set("running", false); f.Window.Dispose(); f.Call("ReportProgress", "disposed");
            }
        }

        [WinFormsTestMethod]
        public void ReviewLayoutAndGitHubCallbacksCoverBoundAndUnboundDocuments()
        {
            using (var f = new Fixture(false))
            {
                f.Window.ClientSize = new System.Drawing.Size(900, 800);
                f.Get<TabControl>("tabs").SelectedTab = f.Get<TabPage>("changesTab"); f.Call("AdjustReviewLayout");
                Assert.IsTrue(f.Get<Label>("help").Visible);
                f.Window.ClientSize = new System.Drawing.Size(900, 500); f.Call("AdjustReviewLayout");
                Assert.IsFalse(f.Get<Label>("help").Visible);
                var pane = f.Get<GitHubPane>("githubPane");
                pane.RepositorySelected("https://github.com/example/review.git", "feature");
                Assert.AreEqual("feature", f.Get<TextBox>("branch").Text);
                Assert.IsNull(pane.LoadDraft());
                f.Set("project", null); pane.OpenModule("Module1", 1); f.Set("project", f.Git.Project);
                Assert.ThrowsException<Microsoft.CSharp.RuntimeBinder.RuntimeBinderException>(() => pane.OpenModule("Module1", 1));
                f.Set("repository", f.Git.Repository);
                pane.RepositorySelected("https://github.com/example/ignored.git", "ignored");
                StringAssert.Contains(f.Status, UiText.Get("This document is already linked. Reopen GitHub to choose another repository."));
                Assert.IsNull(pane.LoadDraft());
                f.Get<TabControl>("tabs").SelectedTab = f.Get<TabPage>("githubTab"); f.Call("AdjustReviewLayout");
                Assert.IsFalse(f.Get<TextBox>("commitMessage").Visible);
                f.Get<TabControl>("tabs").SelectedTab = f.Get<TabPage>("historyTab"); f.Call("AdjustReviewLayout");
                Assert.IsFalse(f.Get<TextBox>("commitMessage").Visible);
                f.Get<TabControl>("tabs").SelectedTab = f.Get<TabPage>("changesTab"); f.Call("AdjustReviewLayout");
                Assert.IsTrue(f.Get<TextBox>("commitMessage").Visible);
            }
        }
    }
}
