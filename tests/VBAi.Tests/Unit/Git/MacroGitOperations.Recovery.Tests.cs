using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitOperationsTests
    {
        [TestMethod]
        public async Task RealDirectoryMarkerBlocksAllReadyActionsBeforeChangingRecoveryRefsOrApplyingSources()
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); string initial = SeedRecoveryRefs(f);
                string checkpoint = f.Repository.Checkpoint(f.Snapshot("2"), "directory marker target").Id;
                Directory.CreateDirectory(f.Repository.RecoveryFile);
                string retained = Path.Combine(f.Repository.RecoveryFile, "retained.txt"); File.WriteAllText(retained, "owned recovery evidence");
                object component = f.Host.VBComponents.Item("Module1");
                foreach (string action in new[] { "commit", "commit_selected", "checkpoint_create", "checkpoint_restore", "branch_create", "branch_switch",
                    "branch_track", "module_restore", "merge_begin", "merge_complete", "merge_abort", "merge_resolve", "pr_prepare", "fetch", "push", "pull", "remote_branches" })
                {
                    var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync(action, name: checkpoint, text: "Synthetic target"));
                    StringAssert.Contains(error.Message, UiText.Get("Restore the interrupted import before continuing."));
                    AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                    Assert.AreSame(component, f.Host.VBComponents.Item("Module1")); Assert.IsTrue(f.Project.Capture().SameAs(before));
                    Assert.AreEqual("owned recovery evidence", File.ReadAllText(retained));
                }
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task RealDirectoryMarkerCannotReportNoOpRollbackSuccessOrStartDifferingRollback(bool different)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); string initial = SeedRecoveryRefs(f);
                string backup = different ? f.Commit(f.Snapshot("2")) : initial;
                f.Repository.SetRef(MacroGitRepository.Backup, backup);
                Directory.CreateDirectory(f.Repository.RecoveryFile);
                var error = await Assert.ThrowsExceptionAsync<IOException>(() => f.Operations.ExecuteAsync("rollback"));
                StringAssert.Contains(error.Message, "not a regular file");
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Project.Capture().SameAs(before));
                Assert.AreEqual(initial, f.Repository.Resolve(MacroGitRepository.AfterImport)); Assert.AreEqual(backup, f.Repository.Resolve(MacroGitRepository.Backup));
                Assert.IsTrue(Directory.Exists(f.Repository.RecoveryFile));
            }
        }

        [DataTestMethod, DataRow("access"), DataRow("io")]
        public async Task ReadyAndRollbackRefuseUnobservableMetadataWithoutChangingRefsOrApplying(string kind)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); string initial = SeedRecoveryRefs(f);
                string checkpoint = f.Repository.Checkpoint(f.Snapshot("2"), "unobservable target").Id;
                File.WriteAllText(f.Repository.RecoveryFile, initial); int observations = 0;
                var original = MarkerFailure(kind);
                f.Repository.RecoveryAttributes = path => { observations++; throw original; };
                foreach (string action in new[] { "checkpoint_restore", "rollback", "commit" })
                {
                    Assert.AreSame(original, await ObserveOperationFailure(() => f.Operations.ExecuteAsync(action, name: checkpoint, text: "Synthetic commit")));
                    AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Project.Capture().SameAs(before));
                }
                Assert.AreEqual(3, observations); Assert.AreEqual(initial, File.ReadAllText(f.Repository.RecoveryFile));
            }
        }

        [DataTestMethod, DataRow("access"), DataRow("io")]
        public async Task StatusAndRevisionCannotReportAnUnreadableRecoveryMarkerAsAbsent(string kind)
        {
            using (var f = new Fixture())
            {
                string initial = SeedRecoveryRefs(f); var before = f.Project.Capture(); var original = MarkerFailure(kind);
                f.Repository.RecoveryAttributes = path => { throw original; };
                Assert.AreSame(original, ObserveSourceFailure(() => f.Operations.Revision(before)));
                Assert.AreSame(original, await ObserveOperationFailure(() => f.Operations.StatusAsync()));
                AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Project.Capture().SameAs(before));
            }
        }

        [DataTestMethod, DataRow(FileAttributes.ReparsePoint), DataRow(FileAttributes.Directory | FileAttributes.ReparsePoint)]
        public async Task RollbackRefusesInjectedInvalidMarkerTypeBeforeSourcesAndRecoveryFileAreTouched(FileAttributes attributes)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); string initial = SeedRecoveryRefs(f);
                File.WriteAllText(f.Repository.RecoveryFile, initial);
                f.Repository.RecoveryAttributes = path => attributes;
                await Assert.ThrowsExceptionAsync<IOException>(() => f.Operations.ExecuteAsync("rollback"));
                AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Project.Capture().SameAs(before)); Assert.AreEqual(initial, File.ReadAllText(f.Repository.RecoveryFile));
            }
        }

        [TestMethod]
        public void PrepareRecoveryRefusesRealDirectoryBeforeOverwritingBackupOrClearingAfterImport()
        {
            using (var f = new Fixture())
            {
                string initial = SeedRecoveryRefs(f); var before = f.Project.Capture();
                Directory.CreateDirectory(f.Repository.RecoveryFile);
                var error = ObserveSourceFailure(() => f.Repository.PrepareRecovery(before));
                AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(Directory.Exists(f.Repository.RecoveryFile));
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
            }
        }

        [TestMethod]
        public async Task MarkerCreatedDuringImportPreviewRefusesPreparationBeforeProtectedRecoveryRefsChange()
        {
            using (var f = new Fixture())
            {
                string initial = SeedRecoveryRefs(f); var before = f.Project.Capture();
                string checkpoint = f.Repository.Checkpoint(f.Snapshot("2"), "preview target").Id;
                f.Operations.ImportPreview = summary => Directory.CreateDirectory(f.Repository.RecoveryFile);
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("checkpoint_restore", name: checkpoint));
                AssertRecoveryRefs(f, initial); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Project.Capture().SameAs(before));
                Assert.IsTrue(Directory.Exists(f.Repository.RecoveryFile));
            }
        }

        [DataTestMethod, DataRow("access"), DataRow("io")]
        public async Task MetadataFailureAtPreparationAfterReadyCannotClearAfterImportOrStartApply(string kind)
        {
            using (var f = new Fixture())
            {
                string initial = SeedRecoveryRefs(f); var before = f.Project.Capture();
                string checkpoint = f.Repository.Checkpoint(f.Snapshot("2"), "prepare target").Id;
                int observations = 0; var original = MarkerFailure(kind);
                f.Repository.RecoveryAttributes = path => { if (++observations == 1) throw new FileNotFoundException("known absent at Ready"); throw original; };
                Assert.AreSame(original, await ObserveOperationFailure(() => f.Operations.ExecuteAsync("checkpoint_restore", name: checkpoint)));
                Assert.AreEqual(2, observations); AssertRecoveryRefs(f, initial);
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Project.Capture().SameAs(before));
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task LegitimateRollbackWithAbsentMarkerStillAllowsNoOpAndMeasuredCompletedImport(bool different)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); SeedRecoveryRefs(f);
                if (different) f.Import(f.Snapshot("2"), before);
                Assert.IsFalse(f.Repository.RecoveryPending);
                await f.Operations.ExecuteAsync("rollback");
                Assert.IsFalse(f.Repository.RecoveryPending); Assert.IsTrue(f.Project.Capture().SameAs(before));
                Assert.AreEqual(different ? 2 : 0, f.Host.VBComponents.ImportAttempts);
            }
        }

        [DataTestMethod, DataRow("metadata"), DataRow("access"), DataRow("delete"), DataRow("concurrent")]
        public void PreMutationPrimaryAndMarkerCompletionErrorsRemainSeparateAndNeverStartApply(string phase)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); var files = f.Snapshot("2");
                var target = new VbaGitSnapshot(new VbaGitManifest { Components = files.Manifest.Components, References = "different references" }, files.Files);
                Exception cleanup = phase == "access" ? (Exception)new UnauthorizedAccessException("synthetic completion access failure") :
                    new IOException("synthetic completion failure"); int observations = 0, deletions = 0;
                f.Repository.RecoveryAttributes = path => { if (++observations == 2 && (phase == "metadata" || phase == "access")) throw cleanup; return File.GetAttributes(path); };
                f.Repository.DeleteRecoveryMarker = path => {
                    deletions++;
                    if (phase == "delete") throw cleanup;
                    File.Delete(path); Directory.CreateDirectory(path);
                };
                var error = Assert.ThrowsException<AggregateException>(() => f.Import(target, before));
                Assert.AreEqual(2, error.InnerExceptions.Count);
                Assert.AreEqual(UiText.Get("VBA references differ. Align them in Tools > References before importing."), error.InnerExceptions[0].Message);
                if (phase != "concurrent") Assert.AreSame(cleanup, error.InnerExceptions[1]);
                else StringAssert.Contains(error.InnerExceptions[1].Message, "single deletion request");
                Assert.AreEqual(phase == "metadata" || phase == "access" ? 0 : 1, deletions); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Project.Capture().SameAs(before)); Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                Assert.IsNull(f.Repository.Resolve(MacroGitRepository.AfterImport));
                Assert.IsTrue(phase == "concurrent" ? Directory.Exists(f.Repository.RecoveryFile) : File.Exists(f.Repository.RecoveryFile));
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task PreMutationAggregateRetainsTheExactPrimaryAndCompletionErrorInstances(bool deletionFails)
        {
            using (var f = new Fixture())
            {
                string initial = SeedRecoveryRefs(f); var before = f.Project.Capture();
                string checkpoint = f.Repository.Checkpoint(f.Snapshot("2"), "exact primary failure target").Id;
                var primary = new IOException("synthetic exact pre-mutation project observation failure");
                var completion = new UnauthorizedAccessException("synthetic exact completion failure");
                int projectObservations = 0, markerObservations = 0, deletions = 0;
                var project = new VbaGitProject(() => f.Host, f.Host.FileName, value => {
                    if (++projectObservations == 3) throw primary; // Execute, import preflight, then Apply preflight.
                    return f.Host.FileName;
                });
                f.Repository.RecoveryAttributes = path => {
                    if (++markerObservations == 3 && !deletionFails) throw completion;
                    return File.GetAttributes(path);
                };
                f.Repository.DeleteRecoveryMarker = path => { deletions++; throw completion; };
                using (var operations = new MacroGitOperations(project, f.Repository))
                {
                    var error = (AggregateException)await ObserveOperationFailure(() => operations.ExecuteAsync("checkpoint_restore", name: checkpoint));
                    Assert.AreEqual(2, error.InnerExceptions.Count); Assert.AreSame(primary, error.InnerExceptions[0]);
                    Assert.AreSame(completion, error.InnerExceptions[1]);
                }
                Assert.AreEqual(3, projectObservations); Assert.AreEqual(3, markerObservations); Assert.AreEqual(deletionFails ? 1 : 0, deletions);
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Project.Capture().SameAs(before));
                Assert.AreEqual(initial, f.Repository.Resolve(f.Repository.Head)); Assert.AreEqual(initial, f.Repository.Resolve(MacroGitRepository.Baseline));
                Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                Assert.IsNull(f.Repository.Resolve(MacroGitRepository.AfterImport)); Assert.IsTrue(File.Exists(f.Repository.RecoveryFile));
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void StartedUncertainImportNeverCompletesRecoveryAndPreservesImportAndOptionalReadbackErrors(bool readbackFails)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); var target = f.Snapshot("2"); int observations = 0, deletions = 0;
                var captureFailure = new IOException("synthetic post-import observation failure");
                var project = new VbaGitProject(() => f.Host, f.Host.FileName, value => {
                    if (readbackFails && f.Host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2")) throw captureFailure;
                    return f.Host.FileName;
                });
                f.Repository.DeleteRecoveryMarker = path => { deletions++; throw new IOException("Unexpected deletion after uncertain native mutation"); };
                f.Host.VBComponents.ThrowAfterImport = true;
                using (var operations = new MacroGitOperations(project, f.Repository))
                {
                    string checkpoint = f.Repository.Checkpoint(target, "uncertain native target").Id;
                    // Ready observes definite absence, then PrepareRecovery independently observes it again.
                    f.Repository.RecoveryAttributes = path => {
                        if (++observations > 2) throw new IOException("Unexpected completion metadata read after uncertain native mutation");
                        return File.GetAttributes(path);
                    };
                    var error = ObserveOperationFailure(() => operations.ExecuteAsync("checkpoint_restore", name: checkpoint)).GetAwaiter().GetResult();
                    if (readbackFails)
                    {
                        var aggregate = error as AggregateException; Assert.IsNotNull(aggregate); Assert.AreEqual(2, aggregate.InnerExceptions.Count);
                        Assert.AreEqual("Simulated failure after applied import", aggregate.InnerExceptions[0].Message); Assert.AreSame(captureFailure, aggregate.InnerExceptions[1]);
                    }
                    else Assert.AreEqual("Simulated failure after applied import", error.Message);
                    Assert.AreEqual(2, observations, "Only Ready/Prepare may inspect metadata; no completion after the mutation boundary was entered.");
                    Assert.AreEqual(0, deletions); Assert.AreEqual(1, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(File.Exists(f.Repository.RecoveryFile));
                    Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                    if (readbackFails) Assert.IsNull(f.Repository.Resolve(MacroGitRepository.AfterImport));
                    else Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.AfterImport)).SameAs(target));
                }
            }
        }

        [DataTestMethod, DataRow("metadata"), DataRow("delete"), DataRow("postobserve"), DataRow("concurrent")]
        public void SuccessfulApplyWithUnverifiedMarkerCompletionRemainsFailureAndNeverReimports(string phase)
        {
            using (var f = new Fixture())
            {
                var before = f.Project.Capture(); var target = f.Snapshot("2"); var original = new IOException("synthetic successful-apply completion failure");
                int observations = 0, deletions = 0;
                f.Repository.RecoveryAttributes = path => {
                    observations++;
                    if ((phase == "metadata" && observations == 2) || (phase == "postobserve" && observations == 3)) throw original;
                    return File.GetAttributes(path);
                };
                f.Repository.DeleteRecoveryMarker = path => {
                    deletions++; if (phase == "delete") throw original;
                    File.Delete(path); if (phase == "concurrent") Directory.CreateDirectory(path);
                };
                var failure = ObserveSourceFailure(() => f.Import(target, before));
                if (phase != "concurrent") Assert.AreSame(original, failure); else StringAssert.Contains(failure.Message, "single deletion request");
                Assert.AreEqual(phase == "metadata" ? 0 : 1, deletions); Assert.AreEqual(1, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Project.Capture().SameAs(target)); Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.AfterImport)).SameAs(target));
                if (phase == "concurrent") Assert.IsTrue(Directory.Exists(f.Repository.RecoveryFile));
                else Assert.AreEqual(phase != "postobserve", File.Exists(f.Repository.RecoveryFile));
            }
        }

        private static string SeedRecoveryRefs(Fixture fixture)
        {
            string initial = fixture.Seed(); fixture.Repository.SetRef(MacroGitRepository.Backup, initial);
            fixture.Repository.SetRef(MacroGitRepository.AfterImport, initial); return initial;
        }
        private static void AssertRecoveryRefs(Fixture fixture, string initial)
        {
            foreach (string reference in new[] { fixture.Repository.Head, MacroGitRepository.Baseline, MacroGitRepository.Backup, MacroGitRepository.AfterImport })
                Assert.AreEqual(initial, fixture.Repository.Resolve(reference), reference);
        }
        private static Exception MarkerFailure(string kind) => kind == "access" ? (Exception)new UnauthorizedAccessException("synthetic recovery metadata access failure") :
            new IOException("synthetic recovery metadata I/O failure");
        private static async Task<Exception> ObserveOperationFailure(Func<Task> operation)
        { try { await operation(); } catch (Exception error) { return error; } Assert.Fail("An uncertain recovery state must not become acceptance."); return null; }
        private static Exception ObserveSourceFailure(Action operation)
        { try { operation(); } catch (Exception error) { return error; } Assert.Fail("An uncertain recovery state must not become acceptance."); return null; }
    }
}
