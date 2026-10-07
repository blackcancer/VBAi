using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitOperationsTests
    {
        [DataTestMethod, DataRow("backup"), DataRow("after"), DataRow("marker")]
        public void NormalImportOwnerAuthorityRaceDoesNotClearTheChangedRecoveryOrEnterNativeMutation(string changed)
        {
            using (var f = new Fixture())
            {
                f.Seed(); var before = f.Project.Capture(); var target = f.Snapshot("2");
                f.Operations.ImportOwnerPreflight = () => ChangeRecoveryAuthority(f, changed);
                var error = Assert.ThrowsException<AggregateException>(() => f.Import(target, before));
                Assert.AreEqual(2, error.InnerExceptions.Count);
                foreach (var inner in error.InnerExceptions)
                    StringAssert.Contains(inner.Message, UiText.Get("The Git/VBA state changed. Read git_status again before making changes."));
                Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts);
                Assert.IsTrue(f.Repository.RecoveryPending); Assert.IsTrue(f.Project.Capture().SameAs(before));
            }
        }

        [DataTestMethod]
        [DataRow("queued", "backup"), DataRow("queued", "after"), DataRow("queued", "marker")]
        [DataRow("owner", "backup"), DataRow("owner", "after"), DataRow("owner", "marker")]
        public void ChangedNewRecoveryDuringOwnerContinuationRefusesBeforeRecaptureAndPreservesAuthority(string phase, string changed)
        {
            using (var f = new Fixture())
            {
                f.Seed(); var before = f.Project.Capture(); var target = f.Snapshot("2");
                f.Repository.RecordImportedState(target);
                f.Repository.PrepareRecovery(before);
                Assert.IsNull(f.Repository.Resolve(MacroGitRepository.AfterImport), "The frozen new import must not adopt a prior measured state.");
                var flags = BindingFlags.Instance | BindingFlags.NonPublic;
                object frozen = typeof(MacroGitOperations).GetMethod("ReadRecoveryState", flags).Invoke(f.Operations, null);
                var guard = (Action)typeof(MacroGitOperations).GetMethod("RecoveryReadbackGuard", flags).Invoke(f.Operations, new[] { frozen });
                int ownerChecks = 0, captures = 0;
                f.Operations.ImportOwnerReadback = () =>
                {
                    ownerChecks++;
                    if (phase == "owner") ChangeRecoveryAuthority(f, changed);
                };
                var previous = SynchronizationContext.Current; var queue = new ReadbackQueue();
                try
                {
                    SynchronizationContext.SetSynchronizationContext(queue);
                    var task = ImportedFormReadback.VerifyAsync(target, before, true, guard, () => { captures++; return target; }, QueueReadbackPulse);
                    Assert.IsFalse(task.IsCompleted); Assert.AreEqual(1, queue.Count);
                    if (phase == "queued") ChangeRecoveryAuthority(f, changed);
                    queue.RunOne();
                    Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                    Assert.AreEqual(phase == "owner" ? 1 : 0, ownerChecks);
                    Assert.AreEqual(0, captures); Assert.AreEqual(0, queue.Count);
                    Assert.AreEqual(0, f.Host.VBComponents.ImportAttempts); Assert.IsTrue(f.Repository.RecoveryPending);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }

        private static async Task QueueReadbackPulse() { await Task.Yield(); }

        private sealed class ReadbackQueue : SynchronizationContext
        {
            private Action pending;
            internal int Count => pending == null ? 0 : 1;
            public override void Post(SendOrPostCallback callback, object state)
            { Assert.IsNull(pending, "The bounded readback must queue only once."); pending = () => callback(state); }
            internal void RunOne() { var work = pending; pending = null; work(); }
        }
    }
}
