using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitUiProtocolTests
    {
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void KnownTerminalSuccessOrFailureAllowsNormalCloseButNeverReplay(bool success)
        {
            var records = new List<object>(); var ledger = new EmbeddedGitUiProtocol(records.Add); int emits = 0;
            ledger.EmitOnce("connect", () => emits++);
            Assert.IsFalse(ledger.CanClose);
            ledger.Terminal("connect", true, success);
            Assert.IsTrue(ledger.CanClose);
            Assert.ThrowsException<InvalidOperationException>(() => ledger.EmitOnce("connect", () => emits++));
            Assert.AreEqual(1, emits); Assert.AreEqual(3, records.Count);
        }

        [DataTestMethod, DataRow("delivery"), DataRow("deadline"), DataRow("observation")]
        public void UncertainDeliveryRefusesCloseFurtherActionAndTerminalPromotion(string failure)
        {
            var primary = new IOException("synthetic failure"); var ledger = new EmbeddedGitUiProtocol(_ => { }); int emits = 0;
            if (failure == "delivery")
                Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => ledger.EmitOnce("connect", () => { emits++; throw primary; })));
            else { ledger.EmitOnce("connect", () => emits++); ledger.MarkUncertain(failure); }
            Assert.IsFalse(ledger.CanClose);
            Assert.ThrowsException<InvalidOperationException>(() => ledger.Terminal("connect", true, true));
            Assert.ThrowsException<InvalidOperationException>(() => ledger.EmitOnce("close", () => emits++));
            Assert.AreEqual(1, emits);
        }

        [TestMethod]
        public void IntentEvidenceFailureEmitsNothingAndRemainsKnown()
        {
            int calls = 0; var primary = new IOException("synthetic disk error");
            var ledger = new EmbeddedGitUiProtocol(_ => throw primary);
            Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => ledger.EmitOnce("connect", () => calls++)));
            Assert.AreEqual(0, calls); Assert.IsTrue(ledger.CanClose);
        }

        [TestMethod]
        public void ReturnedInvokeOrWrongActionCannotStandForOperationCompletion()
        {
            var ledger = new EmbeddedGitUiProtocol(_ => { }); ledger.EmitOnce("connect", () => { });
            Assert.ThrowsException<InvalidOperationException>(() => ledger.Terminal("connect", false, true));
            Assert.ThrowsException<InvalidOperationException>(() => ledger.Terminal("checkpoint", true, true));
            Assert.IsFalse(ledger.CanClose);
        }

        [DataTestMethod, DataRow(0, 2, 3), DataRow(4, 2, 3), DataRow(1, 0, 3), DataRow(1, 4, 3), DataRow(1, 2, 0), DataRow(1, 2, 4)]
        public void ExactOwnerRefusesMissingOrChangedPidTidHandle(int pid, int tid, int hwnd)
        {
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitUiProtocol.RequireOwner(1, 2, 3, pid, (uint)tid, hwnd));
            EmbeddedGitUiProtocol.RequireOwner(1, 2, 3, 1, 2, 3);
        }

        [DataTestMethod, DataRow(true, false, "same", "same", false), DataRow(false, true, "a", "b", false),
            DataRow(true, true, "same", "same", true), DataRow(true, false, "a", "b", true)]
        public void TerminalNeedsIdleAndIndependentOperationProgress(bool idle, bool busy, string before, string after, bool expected)
            => Assert.AreEqual(expected, EmbeddedGitUiProtocol.IsTerminal(idle, busy, before, after));

        [DataTestMethod, DataRow(false, false, false, false), DataRow(false, true, true, false),
            DataRow(true, false, false, true), DataRow(true, true, true, true), DataRow(false, true, false, true)]
        public void KnownPreEmissionAndClosedModalFailuresPermitCleanupWhileNativeUncertaintyRetains(
            bool nativePending, bool menuEmitted, bool modalClosed, bool expected)
            => Assert.AreEqual(expected, EmbeddedGitUiProtocol.MustRetainOwner(nativePending, menuEmitted, modalClosed));

        [TestMethod]
        public void TerminalEvidenceFailureKeepsDeliveryPendingAndPreservesTheOriginalError()
        {
            int records = 0; var primary = new IOException("synthetic terminal evidence failure");
            var ledger = new EmbeddedGitUiProtocol(_ => { if (++records == 3) throw primary; });
            ledger.EmitOnce("connect", () => { });
            Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => ledger.Terminal("connect", true, false)));
            Assert.IsFalse(ledger.CanClose);
            Assert.ThrowsException<InvalidOperationException>(() => ledger.EmitOnce("close", () => { }));
        }

        [TestMethod]
        public void FastCompareAfterCheckpointHasDistinctKnownTerminalProofWithoutRelaxingTheBusyGuard()
        {
            const string before = "Operation complete: checkpoint_create";
            const string after = "First link: commit then push to publish, or pull to import the repository with a backup first.";
            Assert.IsTrue(EmbeddedGitUiProtocol.IsTerminal(true, false, before, after));
            Assert.AreEqual(1, EmbeddedGitAutomation.ClassifyTerminal("compare", after));
            Assert.AreEqual(0, EmbeddedGitAutomation.ClassifyTerminal("compare", before));
            Assert.IsFalse(EmbeddedGitUiProtocol.IsTerminal(true, false, after, after));
            Assert.AreEqual(-1, EmbeddedGitAutomation.ClassifyTerminal("connect", "Known failure · Check the connection, account and Git state, then retry."));
        }

        [TestMethod]
        public void CoordinatorFailurePreservesAdditionalCleanupErrorWithoutDuplicatingTheSamePrimary()
        {
            var primary = new TimeoutException("synthetic readiness failure");
            var cleanup = new IOException("synthetic normal cleanup failure");
            var owner = new AggregateException(primary, cleanup);
            var combined = (AggregateException)EmbeddedGitUiProtocol.PreserveFailures(primary, owner, cleanup);
            Assert.AreEqual(2, combined.InnerExceptions.Count);
            Assert.AreSame(primary, combined.InnerExceptions[0]); Assert.AreSame(cleanup, combined.InnerExceptions[1]);
            Assert.AreSame(primary, EmbeddedGitUiProtocol.PreserveFailures(primary, primary));
            Assert.AreSame(combined, EmbeddedGitUiProtocol.PreserveFailures(combined, owner, cleanup));
        }
    }
}
