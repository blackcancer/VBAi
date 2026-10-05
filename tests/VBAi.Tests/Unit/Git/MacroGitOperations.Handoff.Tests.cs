using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitOperationsTests
    {
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ImportOwnerRefusalAtTheMutationBoundaryPreservesOriginalSourcesAndError(bool rollback)
        {
            using (var f = new Fixture())
            {
                f.Seed(); var before = f.Project.Capture(); var target = f.Snapshot("2");
                object component = f.Host.VBComponents.Item("Module1");
                if (rollback) f.Repository.PrepareRecovery(target);
                var original = new InvalidOperationException("exact owner is disabled"); int checks = 0;
                f.Operations.ImportOwnerPreflight = () => {
                    checks++; Assert.IsTrue(f.Repository.RecoveryPending);
                    Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); throw original;
                };
                Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => f.Import(target, before, rollback)));
                Assert.AreEqual(1, checks); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.AreSame(component, f.Host.VBComponents.Item("Module1")); Assert.IsTrue(f.Project.Capture().SameAs(before));
                Assert.AreEqual(rollback, f.Repository.RecoveryPending); Assert.IsNotNull(f.Repository.Resolve(MacroGitRepository.Backup));
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ImportOwnerCheckRunsOnceImmediatelyBeforeDeliveryAndNeverRetriesUncertainImport(bool uncertain)
        {
            using (var f = new Fixture())
            {
                f.Seed(); var before = f.Project.Capture(); var target = f.Snapshot("2"); int checks = 0;
                f.Host.VBComponents.ThrowAfterImport = uncertain;
                f.Operations.ImportOwnerPreflight = () => {
                    checks++; Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Repository.RecoveryPending);
                };
                if (uncertain) StringAssert.Contains(Assert.ThrowsException<Exception>(() => f.Import(target, before)).Message, "Simulated failure after applied import");
                else f.Import(target, before);
                Assert.AreEqual(1, checks); Assert.AreEqual(1, f.Host.VBComponents.ImportAttempts);
                Assert.AreEqual(uncertain, f.Repository.RecoveryPending);
                if (!uncertain) Assert.IsTrue(f.Project.Capture().SameAs(target));
            }
        }

        [TestMethod]
        public void ExactNoOpDoesNotInventAnImportOwnerCheckOrNativeDelivery()
        {
            using (var f = new Fixture())
            {
                f.Seed(); var before = f.Project.Capture(); int checks = 0;
                f.Operations.ImportOwnerPreflight = () => { checks++; throw new InvalidOperationException(); };
                f.Import(before, before); Assert.AreEqual(0, checks); Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
            }
        }
    }
}
