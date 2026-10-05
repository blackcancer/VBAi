using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class GitWindowCoverageTests
    {
        [WinFormsTestMethod]
        [DataRow("checkpoint_restore")][DataRow("branch_switch")][DataRow("module_restore")]
        [DataRow("pull")][DataRow("merge_complete")][DataRow("rollback")]
        public void ActualWindowHandsOffEachImportOnceAndKeepsTheCacheLeaseThroughReadback(string action)
        { ExerciseHandoff(action, null); }

        [WinFormsTestMethod]
        [DataRow("identity")][DataRow("path")][DataRow("mode")][DataRow("protection")]
        [DataRow("vba")][DataRow("head")][DataRow("branch")][DataRow("recovery")]
        [DataRow("native-uncertain")][DataRow("recovery-directory")][DataRow("owner-before-import")]
        public void DriftOrUncertainNativeFailureAfterActualModalReturnCannotReplayImport(string fault)
        { ExerciseHandoff("checkpoint_restore", fault); }

        private static void ExerciseHandoff(string action, string fault)
        {
            using (var f = new Fixture())
            using (var timer = new Timer { Interval = 10 })
            {
                string initial = f.Git.Seed(); var target = f.Git.Snapshot("2");
                string commit = f.Git.Commit(target, initial), name = null, path = null;
                if (action == "checkpoint_restore") name = f.Git.Repository.Checkpoint(target, "target").Id;
                if (action == "branch_switch" || action == "merge_complete")
                {
                    f.Git.Repository.CreateBranch("feature"); f.Git.Repository.SetRef("refs/heads/feature", commit);
                    name = "feature"; if (action == "merge_complete") f.Git.Repository.BeginMerge("feature");
                }
                if (action == "module_restore") { name = commit; path = "Module1"; }
                if (action == "pull") f.Git.Repository.Push(commit);
                if (action == "rollback") { f.Git.Repository.PrepareRecovery(target); f.Git.Repository.RecordImportedState(f.Git.Project.Capture()); }
                f.Compare(); f.Window.Hide();
                var resolved = f.Git.Host; int executingReads = 0;
                string lease = Path.Combine(f.UiCache, "session.lock");
                f.Set("project", new VbaGitProject(() => resolved, f.Git.Host.FileName, value => {
                    var request = f.Get<GitModalSession.Request>("modalRequest");
                    if (request != null && request.Phase == "Executing")
                    { executingReads++; Assert.ThrowsException<IOException>(() => { using (File.Open(lease, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } }); }
                    return ((global::FakeProject)value).FileName;
                }));
                Task operation = null; int shows = 0, admissions = 0; bool showReturned = false;
                timer.Tick += (sender, args) => {
                    if (!f.Window.Modal) return;
                    if (shows == 2) { f.Window.Close(); return; }
                    if (operation == null) operation = (Task)f.Call("RunGitAction", action, name, "merge", null, path);
                    else if (operation.IsCompleted && !f.Get<bool>("running")) f.Window.Close();
                };
                var session = new GitModalSession(() => {
                    if (++shows == 1)
                    {
                        timer.Start(); f.Window.DialogResult = DialogResult.None; f.Window.ShowDialog(); timer.Stop();
                        Assert.AreEqual(0, f.Git.Host.VBComponents.ImportAttempts);
                        Assert.IsTrue(f.Get<bool>("running")); showReturned = true;
                    }
                    else
                    {
                        Assert.IsTrue(operation.IsCompleted); Assert.IsFalse(f.Get<bool>("running"));
                        Assert.IsNull(f.Get<FileStream>("cacheLock"));
                        using (File.Open(lease, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                        var request = f.Get<GitModalSession.Request>("modalRequest");
                        Assert.AreEqual(fault == null ? "Succeeded" : "Failed", request.Phase);
                        StringAssert.Contains(f.Window.AccessibleDescription, "/" + request.Id + "/" + action + "/" + request.Phase);
                        f.Window.DialogResult = DialogResult.None; timer.Start(); f.Window.ShowDialog(); timer.Stop();
                        Assert.IsFalse(f.Window.IsDisposed, "The same owned form must survive the modal handoff and final modal return.");
                    }
                }, () => {
                    admissions++; Assert.IsTrue(showReturned);
                    Assert.ThrowsException<IOException>(() => { using (File.Open(lease, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } });
                    Assert.AreEqual(0, f.Git.Host.VBComponents.ImportAttempts);
                    if (admissions == 3 && fault == "owner-before-import") throw new InvalidOperationException("Original owner was disabled during asynchronous preparation.");
                    if (admissions > 1) return;
                    int checkpoints = f.Git.Repository.Checkpoints().Length;
                    Assert.IsTrue(((Task)f.Call("RunGitAction", "checkpoint_create", "forbidden duplicate", null, null, null)).IsCompleted);
                    Assert.AreEqual(checkpoints, f.Git.Repository.Checkpoints().Length);
                    if (fault == "identity") resolved = new global::FakeProject { FileName = f.Git.Host.FileName };
                    if (fault == "path") f.Git.Host.FileName += ".other";
                    if (fault == "mode") f.Git.Host.Mode = 1;
                    if (fault == "protection") f.Git.Host.Protection = 1;
                    if (fault == "vba") f.Git.Host.VBComponents.Item("Module1").CodeModule.Text += "' drift\n";
                    if (fault == "head") f.Git.Repository.SetRef(f.Git.Repository.Head, commit);
                    if (fault == "branch") { f.Git.Repository.CreateBranch("changed"); f.Git.Repository.SelectBranch("changed"); }
                    if (fault == "recovery") f.Git.Repository.PrepareRecovery(f.Git.Project.Capture());
                    if (fault == "native-uncertain") f.Git.Host.VBComponents.ThrowAfterImport = true;
                    if (fault == "recovery-directory") Directory.CreateDirectory(f.Git.Repository.RecoveryFile);
                });
                f.Window.AttachModalSession(session);
                try { f.Pump(VbeUiTask.Run(async () => { await session.RunAsync(); return true; })); }
                finally { f.Window.DetachModalSession(session); }
                Assert.AreEqual(2, shows);
                Assert.AreEqual(fault == null || fault == "native-uncertain" || fault == "owner-before-import" ? 3 : 1, admissions);
                Assert.AreEqual(fault == null || fault == "native-uncertain" ? 1 : 0, f.Git.Host.VBComponents.ImportAttempts);
                if (fault == null) { Assert.IsTrue(f.Git.Project.Capture().SameAs(target)); Assert.IsTrue(executingReads > 0); Assert.IsFalse(f.Git.Repository.RecoveryPending); }
                if (fault == "native-uncertain") { Assert.IsTrue(f.Git.Repository.RecoveryPending); Assert.IsNotNull(f.Git.Repository.Resolve(MacroGitRepository.Backup)); }
            }
        }
    }

    public sealed partial class GitWindowStateTests
    {
        private sealed class HandoffOwner : IWin32Window { internal IntPtr Value; public IntPtr Handle => Value; }

        [STATestMethod]
        [DataRow("zero")][DataRow("disabled")][DataRow("destroyed")][DataRow("valid")]
        public void ActualModalEntryValidatesTheOwnerBeforeCallingShow(string state)
        {
            using (var owner = new Form())
            using (var window = new GitWindow())
            {
                IntPtr handle = owner.Handle; owner.Enabled = state != "disabled";
                if (state == "destroyed") owner.Dispose();
                var identity = new HandoffOwner { Value = state == "zero" ? IntPtr.Zero : handle }; int shows = 0;
                var task = GitModalSession.ShowAsync(window, identity, (dialog, parent) => { shows++; return DialogResult.Cancel; });
                if (state == "valid") task.GetAwaiter().GetResult();
                else Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                Assert.AreEqual(state == "valid" ? 1 : 0, shows);
                Assert.IsNull(Field<GitModalSession>(window, "modalSession"));
            }
        }

        [STATestMethod]
        [DataRow(CloseReason.None, true, true)][DataRow(CloseReason.UserClosing, true, false)]
        [DataRow(CloseReason.ApplicationExitCall, true, false)][DataRow(CloseReason.WindowsShutDown, true, false)]
        [DataRow(CloseReason.None, false, false)]
        public void RunningCloseRequiresTheExactPreparedHandoffAndDialogResult(CloseReason reason, bool resultSet, bool permitted)
        {
            using (var window = new GitWindow())
            {
                var request = new GitModalSession.Request("pull", null, null, null, null, null, false, "revision"); request.Queue();
                Set(window, "running", true); Set(window, "modalRequest", request); Set(window, "leavingForImport", true);
                window.DialogResult = resultSet ? DialogResult.OK : DialogResult.None;
                var closing = new FormClosingEventArgs(reason, false); Invoke(window, "OnFormClosing", closing);
                Assert.AreEqual(!permitted, closing.Cancel); Set(window, "running", false);
            }
        }
    }
}
