using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitRepositoryTests
    {
        [TestMethod]
        public void RealDirectoryRecoveryMarkerIsPendingAndCannotBeCompletedAsAnAbsentFile()
        {
            InRecoveryMarker(repository => {
                Directory.CreateDirectory(repository.RecoveryFile);
                string retained = Path.Combine(repository.RecoveryFile, "retained.txt");
                File.WriteAllText(retained, "synthetic recovery evidence");
                Assert.IsTrue(repository.RecoveryPending, "A directory entry cannot be treated as definite absence.");
                Assert.ThrowsException<IOException>(() => repository.CompleteRecovery());
                Assert.IsTrue(Directory.Exists(repository.RecoveryFile)); Assert.AreEqual("synthetic recovery evidence", File.ReadAllText(retained));
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OnlyDefiniteMissingMarkerExceptionsMeanNoPendingRecoveryAndDoNotDelete(bool parentMissing)
        {
            InRecoveryMarker(repository => {
                int deletions = 0, observations = 0;
                repository.RecoveryAttributes = path => { observations++; Assert.AreEqual(repository.RecoveryFile, path);
                    if (parentMissing) throw new DirectoryNotFoundException("synthetic missing parent");
                    throw new FileNotFoundException("synthetic missing marker"); };
                repository.DeleteRecoveryMarker = path => deletions++;
                Assert.IsFalse(repository.RecoveryPending); repository.RequireValidRecoveryMarker(); repository.CompleteRecovery();
                Assert.AreEqual(3, observations); Assert.AreEqual(0, deletions);
            });
        }

        [DataTestMethod, DataRow("access"), DataRow("io"), DataRow("path")]
        public void MetadataFailuresPropagateUnchangedForPendingRollbackValidationAndCompletion(string kind)
        {
            InRecoveryMarker(repository => {
                var original = RecoveryMetadataError(kind); int deletions = 0;
                repository.RecoveryAttributes = path => { Assert.AreEqual(repository.RecoveryFile, path); throw original; };
                repository.DeleteRecoveryMarker = path => deletions++;
                Assert.AreSame(original, RecoveryFailure(() => { bool pending = repository.RecoveryPending; }));
                Assert.AreSame(original, RecoveryFailure(repository.RequireValidRecoveryMarker));
                Assert.AreSame(original, RecoveryFailure(repository.CompleteRecovery)); Assert.AreEqual(0, deletions);
            });
        }

        [DataTestMethod, DataRow(FileAttributes.Normal), DataRow(FileAttributes.ReadOnly), DataRow(FileAttributes.Directory)]
        [DataRow(FileAttributes.ReparsePoint), DataRow(FileAttributes.Directory | FileAttributes.ReparsePoint)]
        public void EveryExistingEntryIsPendingButInvalidMarkerTypesNeverAuthorizeDelete(FileAttributes attributes)
        {
            InRecoveryMarker(repository => {
                int deletions = 0; repository.RecoveryAttributes = path => attributes;
                repository.DeleteRecoveryMarker = path => deletions++;
                Assert.IsTrue(repository.RecoveryPending);
                if ((attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
                {
                    Assert.ThrowsException<IOException>(repository.RequireValidRecoveryMarker);
                    Assert.ThrowsException<IOException>(repository.CompleteRecovery);
                }
                else repository.RequireValidRecoveryMarker();
                Assert.AreEqual(0, deletions);
            });
        }

        [TestMethod]
        public void MarkerObserversAndDeleteBoundariesBelongToOneRepositoryInstance()
        {
            InRecoveryMarker(repository => {
                var other = new MacroGitRepository(Path.GetDirectoryName(repository.RecoveryFile), "other");
                var original = new IOException("instance-only marker metadata failure");
                repository.RecoveryAttributes = path => { throw original; };
                Assert.AreSame(original, RecoveryFailure(() => { bool pending = repository.RecoveryPending; }));
                Assert.IsFalse(other.RecoveryPending);
                File.WriteAllText(other.RecoveryFile, "synthetic marker");
                int firstDeletes = 0; repository.DeleteRecoveryMarker = path => firstDeletes++;
                other.CompleteRecovery(); Assert.AreEqual(0, firstDeletes); Assert.IsFalse(File.Exists(other.RecoveryFile));
            });
        }

        [TestMethod]
        public void RealRegularMarkerIsDeletedOnceAndItsAbsenceIsIndependentlyObserved()
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "synthetic backup identity"); int observations = 0, deletions = 0;
                repository.RecoveryAttributes = path => { observations++; return File.GetAttributes(path); };
                repository.DeleteRecoveryMarker = path => { deletions++; File.Delete(path); };
                repository.CompleteRecovery();
                Assert.AreEqual(1, deletions); Assert.AreEqual(2, observations); Assert.IsFalse(File.Exists(repository.RecoveryFile));
                repository.CompleteRecovery(); Assert.AreEqual(1, deletions, "Already absent is a read-only no-op."); Assert.AreEqual(3, observations);
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DefiniteFileOrParentAbsenceAfterSingleDeletionVerifiesCompletion(bool parentMissing)
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "owned marker"); int observations = 0, deletions = 0;
                repository.RecoveryAttributes = path => {
                    Assert.AreEqual(repository.RecoveryFile, path);
                    if (++observations == 1) return File.GetAttributes(path);
                    if (parentMissing) throw new DirectoryNotFoundException("known absent parent after deletion");
                    throw new FileNotFoundException("known absent marker after deletion");
                };
                repository.DeleteRecoveryMarker = path => { deletions++; File.Delete(path); };
                repository.CompleteRecovery();
                Assert.AreEqual(2, observations); Assert.AreEqual(1, deletions); Assert.IsFalse(File.Exists(repository.RecoveryFile));
            });
        }

        [DataTestMethod, DataRow("access"), DataRow("io")]
        public void DeleteFailureIsNeverRetriedAndPreservesOriginalMarker(string kind)
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "retained backup identity"); int deletions = 0, observations = 0;
                var original = RecoveryMetadataError(kind);
                repository.RecoveryAttributes = path => { observations++; return File.GetAttributes(path); };
                repository.DeleteRecoveryMarker = path => { deletions++; throw original; };
                Assert.AreSame(original, RecoveryFailure(repository.CompleteRecovery));
                Assert.AreEqual(1, deletions); Assert.AreEqual(1, observations); Assert.AreEqual("retained backup identity", File.ReadAllText(repository.RecoveryFile));
            });
        }

        [TestMethod]
        public void DeleteAppliedThenLostOutcomeDoesNotObserveAgainRecreateMarkerOrReplayDeletion()
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "owned marker"); int deletions = 0, observations = 0;
                var original = new IOException("single deletion applied but outcome failed");
                repository.RecoveryAttributes = path => { observations++; return File.GetAttributes(path); };
                repository.DeleteRecoveryMarker = path => { deletions++; File.Delete(path); throw original; };
                Assert.AreSame(original, RecoveryFailure(repository.CompleteRecovery));
                Assert.AreEqual(1, deletions); Assert.AreEqual(1, observations); Assert.IsFalse(File.Exists(repository.RecoveryFile));
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ConcurrentEntryAfterSingleDeleteMakesCompletionFailWithoutDeletingTheNewEntry(bool directory)
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "initial marker"); int deletions = 0, observations = 0;
                repository.RecoveryAttributes = path => { observations++; return File.GetAttributes(path); };
                repository.DeleteRecoveryMarker = path => {
                    deletions++; File.Delete(path);
                    if (directory) { Directory.CreateDirectory(path); File.WriteAllText(Path.Combine(path, "new.txt"), "concurrent entry"); }
                    else File.WriteAllText(path, "concurrent entry");
                };
                StringAssert.Contains(Assert.ThrowsException<IOException>(repository.CompleteRecovery).Message, "single deletion request");
                Assert.AreEqual(1, deletions); Assert.AreEqual(2, observations);
                Assert.AreEqual("concurrent entry", File.ReadAllText(directory ? Path.Combine(repository.RecoveryFile, "new.txt") : repository.RecoveryFile));
                Assert.IsTrue(repository.RecoveryPending);
            });
        }

        [DataTestMethod, DataRow("access"), DataRow("io")]
        public void PostDeletionMetadataFailureCannotBecomeVerifiedCompletionOrCauseAnotherDelete(string kind)
        {
            InRecoveryMarker(repository => {
                File.WriteAllText(repository.RecoveryFile, "owned marker"); int deletions = 0, observations = 0;
                var original = RecoveryMetadataError(kind);
                repository.RecoveryAttributes = path => { if (++observations == 2) throw original; return File.GetAttributes(path); };
                repository.DeleteRecoveryMarker = path => { deletions++; File.Delete(path); };
                Assert.AreSame(original, RecoveryFailure(repository.CompleteRecovery)); Assert.AreEqual(2, observations); Assert.AreEqual(1, deletions);
            });
        }

        private static Exception RecoveryMetadataError(string kind) => kind == "access" ? (Exception)new UnauthorizedAccessException("synthetic marker access failure") :
            kind == "path" ? (Exception)new PathTooLongException("synthetic marker path failure") : new IOException("synthetic marker I/O failure");
        private static Exception RecoveryFailure(Action action)
        {
            try { action(); } catch (Exception error) { return error; }
            Assert.Fail("An unreadable or invalid recovery state must not become success."); return null;
        }
        private static void InRecoveryMarker(Action<MacroGitRepository> scenario)
        {
            var scratch = global::GitScratchDirectory.Create();
            try { scenario(new MacroGitRepository(scratch.Root, "main")); }
            finally { scratch.ValidateCleanupRoot(scratch.Root); Directory.Delete(scratch.Root, true); }
        }
    }
}
