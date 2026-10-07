using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitOperationsTests
    {
        [DataTestMethod, DataRow("backup"), DataRow("after"), DataRow("marker")]
        public void RevisionIncludesExactRecoveryRefsAndMarkerContentsEvenWhenRecoveryRemainsPending(string changed)
        {
            using (var f = new Fixture())
            {
                var live = f.Project.Capture();
                f.Repository.PrepareRecovery(live); f.Repository.RecordImportedState(live);
                string reviewed = f.Operations.Revision(live);
                ChangeRecoveryAuthority(f, changed);
                Assert.IsTrue(f.Repository.RecoveryPending);
                Assert.AreNotEqual(reviewed, f.Operations.Revision(live));
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
            }
        }

        [TestMethod]
        public void RevisionDistinguishesAnAbsentMarkerFromAnEmptyRegularMarker()
        {
            using (var f = new Fixture())
            {
                var live = f.Project.Capture(); string reviewed = f.Operations.Revision(live);
                File.WriteAllBytes(f.Repository.RecoveryFile, new byte[0]);
                Assert.AreNotEqual(reviewed, f.Operations.Revision(live));
            }
        }

        [DataTestMethod]
        [DataRow("reviewed", "backup"), DataRow("reviewed", "after"), DataRow("reviewed", "marker")]
        [DataRow("preview", "backup"), DataRow("preview", "after"), DataRow("preview", "marker")]
        [DataRow("owner", "backup"), DataRow("owner", "after"), DataRow("owner", "marker")]
        public async Task RollbackRejectsChangedAuthorityBeforeNativeEntryAndPreservesRecovery(string phase, string changed)
        {
            using (var f = new Fixture())
            {
                var backup = f.Project.Capture();
                f.Host.VBComponents.Item("Module1").CodeModule.Text =
                    f.Host.VBComponents.Item("Module1").CodeModule.Text.Replace("Value = 1", "Value = 2");
                var live = f.Project.Capture();
                f.Repository.PrepareRecovery(backup); f.Repository.RecordImportedState(live);
                string reviewed = f.Operations.Revision(live); int boundaries = 0;
                Action mutate = () => { boundaries++; ChangeRecoveryAuthority(f, changed); };
                if (phase == "reviewed") mutate();
                else if (phase == "preview") f.Operations.ImportPreview = _ => mutate();
                else f.Operations.ImportOwnerPreflight = mutate;
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => f.Operations.ExecuteAsync("rollback", reviewed));
                StringAssert.Contains(error.Message, UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                Assert.AreEqual(1, boundaries); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Project.Capture().SameAs(live)); Assert.IsTrue(f.Repository.RecoveryPending);
                Assert.IsTrue(f.Repository.Read(f.Repository.Resolve(MacroGitRepository.Backup)) != null);
            }
        }

        [TestMethod]
        public async Task UnchangedReviewedRollbackRestoresTheFrozenSnapshotAndCompletesRecovery()
        {
            using (var f = new Fixture())
            {
                var backup = f.Project.Capture();
                f.Host.VBComponents.Item("Module1").CodeModule.Text =
                    f.Host.VBComponents.Item("Module1").CodeModule.Text.Replace("Value = 1", "Value = 2");
                var live = f.Project.Capture();
                f.Repository.PrepareRecovery(backup); f.Repository.RecordImportedState(live);
                await f.Operations.ExecuteAsync("rollback", f.Operations.Revision(live));
                Assert.IsTrue(f.Project.Capture().SameAs(backup)); Assert.IsFalse(f.Repository.RecoveryPending);
                Assert.AreEqual(1, f.Host.VBComponents.ImportAttempts);
            }
        }

        [TestMethod]
        public void OversizedOrLockedMarkerCannotBecomeARevisionOrAReadableRecoveryIdentity()
        {
            using (var f = new Fixture())
            {
                var live = f.Project.Capture();
                File.WriteAllBytes(f.Repository.RecoveryFile, new byte[257]);
                Assert.ThrowsException<IOException>(() => f.Operations.Revision(live));
                File.WriteAllText(f.Repository.RecoveryFile, "retained marker", Encoding.ASCII);
                using (var held = new FileStream(f.Repository.RecoveryFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    Assert.ThrowsException<IOException>(() => f.Operations.Revision(live));
                Assert.AreEqual("retained marker", File.ReadAllText(f.Repository.RecoveryFile));
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
            }
        }

        private static void ChangeRecoveryAuthority(Fixture fixture, string changed)
        {
            string alternate = fixture.Commit(fixture.Snapshot("3"));
            if (changed == "marker") File.WriteAllText(fixture.Repository.RecoveryFile, alternate, Encoding.ASCII);
            else fixture.Repository.SetRef(changed == "backup" ? MacroGitRepository.Backup : MacroGitRepository.AfterImport, alternate);
        }
    }
}
