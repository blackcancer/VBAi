namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie les préconditions, opérations et récupérations du coordinateur Git VBA.</summary>
    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed partial class MacroGitOperationsTests
    {
        [TestMethod]
        public void RevisionUsesLogicalFormDataWithoutLosingRepositoryStateOrTransportBytes()
        {
            using (var f = new Fixture())
            {
                byte[] original = FormResourcePreflightTests.Resource();
                byte[] other = (byte[])original.Clone(); other[24 + 1024 + 108] = 42;
                var baseline = VbaGitSnapshotCoverageTests.LogicalForm(original);
                var unchanged = VbaGitSnapshotCoverageTests.LogicalForm(other);
                string revision = f.Operations.Revision(baseline);
                Assert.AreEqual(revision, f.Operations.Revision(unchanged));
                other[24 + 2048]++;
                Assert.AreNotEqual(revision, f.Operations.Revision(VbaGitSnapshotCoverageTests.LogicalForm(other)));
                File.WriteAllText(f.Repository.RecoveryFile, "qualification recovery marker");
                Assert.AreNotEqual(revision, f.Operations.Revision(baseline));
                CollectionAssert.AreEqual(original, baseline.Serialize()["Form1.frx"]);
            }
        }

        /// <summary>Vérifie la validation du binding et la propriété du verrou de session.</summary>
        [TestMethod]
        public void OpenBindingValidationLockOwnershipAndReleaseMatrix()
        {
            var previous = MacroGitOperations.CacheDirectory;
            try
            {
                using (var f = new Fixture())
                {
                    string cache = Path.Combine(f.Root, "binding"); Directory.CreateDirectory(cache);
                    MacroGitOperations.CacheDirectory = _ => cache;
                    Assert.ThrowsException<InvalidOperationException>(() => MacroGitOperations.Open(f.Project, "scope", null));
                    string path = Path.Combine(cache, "binding.json");
                    File.WriteAllText(path, "{invalid");
                    Assert.ThrowsException<ArgumentException>(() => MacroGitOperations.Open(f.Project, "scope", null));
                    using (new FileStream(Path.Combine(cache, "session.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                    File.WriteAllText(path, new JavaScriptSerializer().Serialize(new Dictionary<string, string>
                    { ["Remote"] = "https://github.com/example/coverage-fixture.git", ["Branch"] = "main" }));
                    using (var opened = MacroGitOperations.Open(f.Project, "scope", null))
                    {
                        Assert.AreEqual("main", opened.Repository.Branch);
                        Assert.ThrowsException<IOException>(() => MacroGitOperations.Open(f.Project, "scope", null));
                    }
                    using (var opened = MacroGitOperations.Open(f.Project, "scope", null)) { opened.Dispose(); }
                }
            }
            finally { MacroGitOperations.CacheDirectory = previous; }
        }

        /// <summary>Vérifie les règles de commit, le repli sur la base distante, la sélection et les no-op.</summary>
        /// <returns>Tâche terminée après les assertions asynchrones.</returns>
        [TestMethod]
        public async Task CommitPolicyParentFallbackSelectionAndNoOpMatrix()
        {
            using (var f = new Fixture())
            {
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => f.Operations.ExecuteAsync("commit", text: " "));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("commit_selected", text: "Empty"));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("commit_selected", text: "Empty", modules: new string[0]));
                dynamic initial = await f.Operations.ExecuteAsync("commit_selected", text: "Selected", modules: new[] { "Module1" });
                dynamic noOp = await f.Operations.ExecuteAsync("commit", text: "No op");
                Assert.AreEqual((string)initial.Commit, (string)noOp.Commit);
                dynamic references = await f.Operations.ExecuteAsync("commit_selected", text: "Only references", references: true);
                Assert.AreEqual((string)initial.Commit, (string)references.Commit);
                string parent = f.Repository.Resolve(f.Repository.Head);
                f.Git("--git-dir=" + f.Cache, "update-ref", "-d", f.Repository.Head);
                f.Repository.SetRef("refs/remotes/origin/selected", parent);
                dynamic remoteBase = await f.Operations.ExecuteAsync("commit", text: "Remote base");
                Assert.AreEqual(parent, (string)remoteBase.Commit);
                f.Git("--git-dir=" + f.Cache, "update-ref", "-d", f.Repository.Head);
                f.Host.VBComponents.Item("Module1").CodeModule.Text += "' changed\n";
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("commit", text: "Refuse unimported"));
            }
        }

        /// <summary>Vérifie les refus pendant fusion, récupération et modifications locales.</summary>
        /// <returns>Tâche terminée après les assertions asynchrones.</returns>
        [TestMethod]
        public async Task RecoveryMergeDirtyAndUnknownActionGuardsMatrix()
        {
            using (var f = new Fixture())
            {
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => f.Operations.ExecuteAsync("unknown"));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("push"));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("pull"));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("branch_switch", name: "main"));
                f.Seed(); f.Repository.CreateBranch("feature"); f.Repository.BeginMerge("feature");
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("commit", text: "Blocked merge"));
                await f.Operations.ExecuteAsync("merge_abort");
                f.Repository.PrepareRecovery(f.Project.Capture());
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("checkpoint_create", name: "Blocked recovery"));
                f.Repository.CompleteRecovery();
                f.Host.VBComponents.Item("Module1").CodeModule.Text += "' dirty\n";
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("pull"));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("branch_switch", name: "feature"));
            }
        }

        /// <summary>Vérifie la découverte distante, fetch, push, pull et rollback.</summary>
        /// <returns>Tâche terminée après les assertions asynchrones.</returns>
        [TestMethod]
        public async Task LocalRemoteDiscoveryFetchPushAndPullMatrix()
        {
            using (var f = new Fixture())
            {
                string initial = f.Seed();
                await f.Operations.ExecuteAsync("remote_branches");
                await f.Operations.ExecuteAsync("fetch");
                await f.Operations.ExecuteAsync("push");
                Assert.AreEqual(initial, f.Repository.Fetch());
                await f.Operations.ExecuteAsync("push");
                string incoming = f.Commit(f.Snapshot("2"), initial); f.Repository.Push(incoming);
                await f.Operations.ExecuteAsync("pull");
                Assert.IsTrue(f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2"));
                Assert.AreEqual(incoming, f.Repository.Resolve(f.Repository.Head));
                await f.Operations.ExecuteAsync("rollback");
                Assert.IsTrue(f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 1"));
            }
        }

        /// <summary>Vérifie les préconditions d’import, les no-op, rollback et erreurs de récupération.</summary>
        [TestMethod]
        public void ImportPreflightNoOpRollbackAndRecoveryFailureMatrix()
        {
            using (var f = new Fixture())
            {
                var live = f.Project.Capture(); var target = f.Snapshot("2");
                Assert.ThrowsException<InvalidOperationException>(() => f.Import(null, live));
                Assert.ThrowsException<InvalidOperationException>(() => f.Import(target, target));
                f.Import(live, live);
                f.Repository.PrepareRecovery(live); f.Import(live, live, true);
                Assert.IsFalse(f.Repository.RecoveryPending);
                f.Operations.ImportPreview = _ => { }; f.Import(target, live);
                var current = f.Project.Capture(); f.Import(live, current, true);
                Assert.IsTrue(f.Project.Capture().SameAs(live));
                var wrongReferences = new VbaGitSnapshot(new VbaGitManifest
                { Components = live.Manifest.Components, References = "Different" }, live.Files);
                Assert.ThrowsException<InvalidOperationException>(() => f.Import(wrongReferences, live));
                Assert.IsFalse(f.Repository.RecoveryPending);
                f.Host.VBComponents.ThrowAfterImport = true;
                Assert.ThrowsException<Exception>(() => f.Import(target, live));
                Assert.IsTrue(f.Repository.RecoveryPending);
            }
        }

        /// <summary>Preserves import and readback errors without claiming a known recoverable after-state.</summary>
        [TestMethod]
        public async Task ImportAndReadbackFailuresRetainBothErrorsAndRequireMeasuredRecovery()
        {
            using (var f = new Fixture())
            {
                string initial = f.Seed();
                var before = f.Project.Capture();
                var target = f.Snapshot("2");
                var readbackFailure = new IOException("Simulated post-import capture failure");
                bool failReadback = true;
                int failedReadbacks = 0;
                var project = new VbaGitProject(() => f.Host, f.Host.FileName, _ => {
                    if (failReadback && f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2"))
                    {
                        failedReadbacks++;
                        throw readbackFailure;
                    }
                    return f.Host.FileName;
                });
                using (var operations = new MacroGitOperations(project, f.Repository))
                {
                    f.Host.VBComponents.ThrowAfterImport = true;
                    string checkpoint = f.Repository.Checkpoint(target, "Dual-failure target").Id;
                    var error = await Assert.ThrowsExceptionAsync<AggregateException>(() =>
                        operations.ExecuteAsync("checkpoint_restore", operations.Revision(before), name: checkpoint));

                    Assert.AreEqual(2, error.InnerExceptions.Count);
                    Assert.AreEqual("Simulated failure after applied import", error.InnerExceptions[0].Message);
                    Assert.AreSame(readbackFailure, error.InnerExceptions[1]);
                    Assert.AreEqual(1, failedReadbacks, "An uncertain native outcome must not trigger another capture or mutation.");
                    Assert.IsTrue(f.Repository.RecoveryPending);
                    Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                    Assert.IsNull(f.Repository.Resolve(MacroGitRepository.AfterImport), "The unobserved state must not be fabricated.");
                    Assert.AreEqual(initial, f.Repository.Resolve(f.Repository.Head));
                    Assert.AreEqual(initial, f.Repository.Resolve(MacroGitRepository.Baseline));
                    StringAssert.Contains(f.Host.VBComponents.Item("Module1").CodeModule.Text, "Value = 2");

                    failReadback = false;
                    f.Host.VBComponents.ThrowAfterImport = false;
                    await Assert.ThrowsExceptionAsync<InvalidOperationException>(() =>
                        operations.ExecuteAsync("rollback", operations.Revision(target)));
                    Assert.IsTrue(f.Repository.RecoveryPending);
                    Assert.IsTrue(f.Project.Capture().SameAs(target), "Refused rollback must preserve the observed project.");
                }
            }
        }

        /// <summary>Refuse le rollback sans marqueurs valides, avec code modifié ou sauvegarde absente.</summary>
        /// <returns>Tâche terminée après les assertions asynchrones.</returns>
        [TestMethod]
        public async Task RollbackRejectsMissingChangedAndUnavailableBackupMatrix()
        {
            using (var f = new Fixture())
            {
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("rollback"));
                string initial = f.Seed(); f.Repository.SetRef(MacroGitRepository.AfterImport, initial);
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("rollback"));
                f.Repository.SetRef(MacroGitRepository.Backup, initial);
                await f.Operations.ExecuteAsync("rollback");
                f.Host.VBComponents.Item("Module1").CodeModule.Text += "' changed\n";
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("rollback"));
            }
        }

        /// <summary>Vérifie brouillon PR, restauration ciblée, fusion et sélection de branche.</summary>
        /// <returns>Tâche terminée après les assertions asynchrones.</returns>
        [TestMethod]
        public async Task DraftModuleRestoreMergeDefaultMessageAndBranchSelectionMatrix()
        {
            using (var f = new Fixture())
            {
                string initial = f.Seed();
                await f.Operations.ExecuteAsync("pr_prepare", name: "main", text: "Draft title", choice: "Draft body");
                await f.Operations.ExecuteAsync("branch_create", name: "feature");
                string feature = f.Commit(f.Snapshot("2"), initial);
                f.Repository.SetRef("refs/heads/feature", feature);
                await f.Operations.ExecuteAsync("merge_begin", name: "feature");
                await f.Operations.ExecuteAsync("merge_complete");
                Assert.IsNull(f.Repository.PendingMerge);
                Assert.IsTrue(f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2"));
                await f.Operations.ExecuteAsync("module_restore", name: initial, path: "Module1");
                Assert.IsTrue(f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 1"));
                await f.Operations.ExecuteAsync("commit", text: "Restoration");
                await f.Operations.ExecuteAsync("branch_switch", name: "feature");
                Assert.AreEqual("feature", f.Repository.Branch);
            }
        }
    }
}
