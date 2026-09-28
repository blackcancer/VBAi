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

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed partial class GitWindowCoverageTests
    {
        [WinFormsTestMethod]
        public void ConnectUsesRealLocalGitAndPersistsAndReplacesBinding()
        {
            using (var f = new Fixture(false))
            {
                const string url = "https://github.com/example/coverage-fixture.git";
                f.Get<TextBox>("remote").Text = url;
                f.Event("Connect_Click");
                Assert.IsNotNull(f.Get<MacroGitRepository>("repository"));
                string binding = Path.Combine(f.UiCache, "binding.json");
                StringAssert.Contains(File.ReadAllText(binding), url);
                f.Set("repository", null);
                f.Event("Connect_Click");
                Assert.IsFalse(File.Exists(Path.Combine(f.UiCache, "binding.pending")));
                using (var loaded = new GitWindow(f.Git.Project, "fixture", "Reloaded binding"))
                    Assert.AreEqual(url, ((TextBox)loaded.Controls.Find("remote", true)[0]).Text);
            }
        }

        [WinFormsTestMethod]
        public void PerformLockBranchRefreshCancellationAndClosingMatrix()
        {
            using (var f = new Fixture())
            {
                int called = 0;
                Func<Task> action = () => { called++; return Task.CompletedTask; };
                f.Pump((Task)f.Call("Perform", action, false)); Assert.AreEqual(1, called);
                using (new FileStream(Path.Combine(f.UiCache, "session.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                { f.Pump((Task)f.Call("Perform", action, false)); Assert.AreEqual(1, called); }
                const string url = "https://github.com/example/coverage-fixture.git";
                f.Git.Git("--git-dir=" + f.Git.Cache, "remote", "set-url", "origin", url);
                f.Get<TextBox>("remote").Text = url; f.Set("displayedBranch", "stale");
                f.Pump((Task)f.Call("Perform", action, false)); Assert.AreEqual(1, called);
                f.Pump((Task)f.Call("Perform", action, true)); Assert.AreEqual(2, called);
                f.Pump((Task)f.Call("Perform", new Func<Task>(() => { throw new OperationCanceledException(); }), true));
                Assert.AreEqual(UiText.Get("Operation cancelled."), f.Status);
                f.Set("running", true);
                var closing = new FormClosingEventArgs(CloseReason.UserClosing, false); f.Call("OnFormClosing", closing); Assert.IsTrue(closing.Cancel);
                f.Set("running", false); closing = new FormClosingEventArgs(CloseReason.UserClosing, false); f.Call("OnFormClosing", closing); Assert.IsFalse(closing.Cancel);
            }
        }

        [WinFormsTestMethod]
        public void CompareTracksBaselineDirtyAndInterruptedImportStates()
        {
            using (var f = new Fixture())
            {
                f.Compare(); StringAssert.Contains(f.Status, UiText.Get("First link: commit then push to publish, or pull to import the repository with a backup first."));
                f.Git.Seed(); f.Compare(); Assert.AreEqual(UiText.Get("VBA matches the last synchronized state."), f.Status);
                f.Git.Host.VBComponents.Item("Module1").CodeModule.Text += "' changed\n";
                f.Compare(); StringAssert.Contains(f.Status, UiText.Get(" file(s) changed since the last synchronization."));
                f.Git.Repository.PrepareRecovery(f.Git.Project.Capture()); f.Compare();
                Assert.AreEqual(UiText.Get("Import interrupted: restore VBA before continuing. The backup is preserved in the cache."), f.Status);
                f.Git.Repository.CompleteRecovery();
                f.Git.Repository.CreateBranch("feature"); f.Git.Repository.BeginMerge("feature"); f.Compare();
                Assert.IsNotNull(f.Git.Repository.PendingMerge);
            }
        }

        [WinFormsTestMethod]
        public void BranchCheckpointRemoteAndMergeActionsUseNativeGitAndReturnToLiveState()
        {
            using (var f = new Fixture())
            {
                f.Git.Seed(); f.Compare();
                f.Get<TextBox>("checkpointName").Text = "Fixture checkpoint"; f.Event("CheckpointCreate_Click");
                var checkpoints = f.Get<ListBox>("checkpointList"); f.Select(checkpoints, 0); f.Event("CheckpointRestore_Click");
                f.Get<ComboBox>("branchName").Text = "feature"; f.Event("BranchCreate_Click");
                f.Select(f.Get<ListBox>("branchList"), 0); f.Event("BranchSwitch_Click");
                f.Select(f.Get<ListBox>("branchList"), 0);
                f.Event("MergeBegin_Click"); f.Event("MergeAbort_Click");
                f.Git.Repository.Push(f.Git.Repository.Resolve(f.Git.Repository.Head));
                f.Event("RemoteBranches_Click"); Assert.IsTrue(f.Get<ComboBox>("branchName").Items.Count > 0);
                f.Get<ComboBox>("branchName").Text = "unknown"; f.Event("BranchTrack_Click");
                Assert.IsFalse(f.Get<bool>("running"));
                f.Action("merge_begin", "feature"); f.Get<TextBox>("commitMessage").Text = "Fixture merge"; f.Event("MergeComplete_Click");
                Assert.IsNull(f.Git.Repository.PendingMerge);
            }
        }

        [WinFormsTestMethod]
        public void ConflictPreviewAndOursTheirsTextResolutionMatrix()
        {
            foreach (string choice in new[] { "ours", "theirs", "text" })
            {
                using (var f = new Fixture())
                {
                    string initial = f.Git.Seed(); f.Git.Repository.CreateBranch("feature");
                    string incoming = f.Git.Commit(f.Git.Snapshot("2"), initial);
                    f.Git.Repository.SetRef("refs/heads/feature", incoming);
                    string local = f.Git.Commit(f.Git.Snapshot("3"), initial);
                    f.Git.Repository.SetRef(f.Git.Repository.Head, local); f.Git.Repository.SetRef(MacroGitRepository.Baseline, local);
                    f.Git.Host.VBComponents.Item("Module1").CodeModule.Text = System.Text.Encoding.UTF8.GetString(f.Git.Snapshot("3").Files["Module1.bas"]);
                    f.Action("merge_begin", "feature");
                    var conflicts = f.Get<ListBox>("conflictList"); Assert.AreEqual(1, conflicts.Items.Count);
                    f.Select(conflicts, 0); f.Event("ConflictList_SelectedIndexChanged");
                    Assert.IsTrue(f.Get<DataGridView>("conflictDiff").Rows.Count > 0);
                    f.Get<TextBox>("resolutionText").Text = System.Text.Encoding.UTF8.GetString(f.Git.Snapshot("4").Files["Module1.bas"]);
                    f.Event(choice == "ours" ? "MergeOurs_Click" : choice == "theirs" ? "MergeTheirs_Click" : "MergeText_Click");
                    Assert.AreEqual(0, f.Git.Repository.PendingMerge.Conflicts.Length);
                }
            }
        }

        [WinFormsTestMethod]
        public void CommitSynchronizationStaleStateAndOperationStatusMatrix()
        {
            using (var f = new Fixture())
            {
                f.Event("RemoteBranches_Click");
                Assert.AreEqual(0, f.Get<ComboBox>("branchName").Items.Count);
                f.Set("running", true); f.Call("RemoteBranches_Click", null, EventArgs.Empty); f.Set("running", false);
                f.Git.Seed(); f.Compare();
                f.Git.Host.VBComponents.Item("Module1").CodeModule.Text += "' stale live snapshot\n";
                f.Get<TextBox>("commitMessage").Text = "Coverage commit";
                f.Event("Commit_Click"); StringAssert.Contains(f.Status, UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                f.Call("PopulateChanges", f.Git.Project.Capture(), f.Git.Repository.Read(f.Git.Repository.Resolve(MacroGitRepository.Baseline)));
                f.Set("displayedLive", null); f.Event("Commit_Click");
                Assert.AreEqual(UiText.Get("Local commit created. Use Push to publish it."), f.Status);
                f.Action("commit", text: "Complete snapshot");
                Assert.AreEqual(UiText.Get("Local commit created. Use Push to publish it."), f.Status);
                f.Event("Push_Click"); Assert.AreEqual(UiText.Get("Push complete."), f.Status);
                f.Event("Fetch_Click"); Assert.IsFalse(f.Get<bool>("running"));
                f.Event("Pull_Click"); Assert.AreEqual(UiText.Get("Pull and import complete. Check and save the document."), f.Status);
                f.Git.Repository.PrepareRecovery(f.Git.Project.Capture());
                f.Git.Repository.SetRef(MacroGitRepository.AfterImport, f.Git.Repository.Resolve(f.Git.Repository.Head));
                f.Event("Restore_Click"); Assert.AreEqual(UiText.Get("VBA restored. Check and save the document."), f.Status);
                f.Git.Repository.PrepareRecovery(f.Git.Project.Capture()); f.Event("Commit_Click");
                Assert.IsTrue(f.Git.Repository.RecoveryPending); f.Git.Repository.CompleteRecovery();
            }
            using (var window = new GitWindow())
            {
                typeof(GitWindow).GetMethod("Dispose", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                    null, new[] { typeof(bool) }, null).Invoke(window, new object[] { false });
            }
            using (var f = new Fixture())
            {
                f.Set("cacheLock", new FileStream(Path.Combine(f.UiCache, "dispose.lock"), FileMode.Create, FileAccess.Write, FileShare.None));
                f.Get<System.ComponentModel.IContainer>("components").Dispose(); f.Set("components", null);
                f.Window.Dispose();
                using (File.OpenWrite(Path.Combine(f.UiCache, "dispose.lock"))) { }
            }
        }

        [WinFormsTestMethod]
        public void BinaryConflictPreviewAndSelectionChangesDuringReadAreHandled()
        {
            foreach (int selection in new[] { 0, -1, 1 })
            {
                using (var f = new Fixture())
                {
                    Func<byte, VbaGitSnapshot> snapshot = resource => new VbaGitSnapshot(
                        new VbaGitManifest { References = "", Components = new[] { new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } } },
                        new System.Collections.Generic.Dictionary<string, byte[]> {
                            ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Form1\"\n"), ["Form1.frx"] = new byte[] { 0, resource } });
                    string initial = f.Git.Commit(snapshot(1)); f.Git.Repository.SetRef(f.Git.Repository.Head, initial);
                    f.Git.Repository.CreateBranch("feature"); f.Git.Repository.SetRef("refs/heads/feature", f.Git.Commit(snapshot(2), initial));
                    f.Git.Repository.SetRef(f.Git.Repository.Head, f.Git.Commit(snapshot(3), initial));
                    f.Git.Repository.BeginMerge("feature"); f.Compare();
                    var conflicts = f.Get<ListBox>("conflictList"); Assert.AreEqual(1, conflicts.Items.Count);
                    conflicts.Items.Add("vba/another.frx"); f.Select(conflicts, 0);
                    using (var changed = new System.Threading.ManualResetEventSlim())
                    {
                        int called = 0;
                        if (selection != 0) f.Git.Repository.Progress = _ => {
                            if (System.Threading.Interlocked.Increment(ref called) != 1) return;
                            f.Window.BeginInvoke(new Action(() => { conflicts.SelectedIndex = selection; changed.Set(); }));
                            if (!changed.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Selection change was not processed.");
                        };
                        f.Event("ConflictList_SelectedIndexChanged");
                        Assert.AreEqual("", f.Get<TextBox>("resolutionText").Text);
                        if (selection == 0) Assert.IsTrue(f.Get<DataGridView>("conflictDiff").Rows.Count > 0);
                        else Assert.AreEqual(0, f.Get<DataGridView>("conflictDiff").Rows.Count);
                    }
                }
            }
        }
    }
}
