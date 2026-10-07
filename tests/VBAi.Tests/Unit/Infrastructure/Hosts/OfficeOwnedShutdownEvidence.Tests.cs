using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks one-shot terminal Quit and independently observed exit/handle disposition without any Office process.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeOwnedShutdownEvidenceTests
    {
        private static OfficeOwnedShutdownEvidence Create() => new OfficeOwnedShutdownEvidence(42,
            "2026-10-01T00:00:00.0000000Z", "C:\\OwnedFixture\\SyntheticHost.exe", "0x0000000000000123", "owned-candidate");

        [TestMethod]
        public void ReturnedQuitWithNonexitRetainsOriginalIdentityAndRejectsNativeReplay()
        {
            var state = Create(); int quits = 0, disposals = 0, codeReads = 0;
            state.Prepare(() => { }); state.QuitOnce(() => quits++, () => { });
            Assert.IsFalse(state.ObserveExit(() => false, () => { codeReads++; return 0; }, () => disposals++, () => { }));
            Assert.AreEqual("RETURNED", state.Record["QuitOutcome"]);
            Assert.AreEqual("RETAINED_EXIT_NOT_OBSERVED", state.Record["State"]);
            Assert.AreEqual(true, state.Record["OwnershipRetained"]); Assert.AreEqual(true, state.Record["ProcessHandleRetained"]);
            Assert.AreEqual(false, state.Record["ProcessExitObserved"]); Assert.AreEqual(false, state.Record["ExitCodeObserved"]);
            Assert.AreEqual("0x0000000000000123", state.Record["OriginalProcessHandle"]);
            Assert.ThrowsException<InvalidOperationException>(() => state.Prepare(() => { }));
            Assert.ThrowsException<InvalidOperationException>(() => state.QuitOnce(() => quits++, () => { }));
            Assert.ThrowsException<InvalidOperationException>(() => state.ObserveExit(() => true, () => 0, () => disposals++, () => { }));
            Assert.AreEqual("RETAINED_EXIT_NOT_OBSERVED", state.Record["State"], "A later observation must not overwrite the original failed gate.");
            Assert.AreEqual(1, quits); Assert.AreEqual(0, codeReads); Assert.AreEqual(0, disposals);
        }

        [DataTestMethod, DataRow(0), DataRow(-1073741819)]
        public void ObservedExitKeepsExactCodeAndDisposesOnlyAfterDurableObservation(int code)
        {
            var state = Create(); int quits = 0, disposals = 0, codeReads = 0;
            state.Prepare(() => { }); state.QuitOnce(() => quits++, () => { });
            Assert.IsTrue(state.ObserveExit(() => true, () => { codeReads++; return code; }, () =>
            {
                Assert.AreEqual(true, state.Record["ProcessExitObserved"]); Assert.AreEqual(true, state.Record["ExitCodeObserved"]);
                Assert.AreEqual(code, state.Record["ExitCode"]); disposals++;
            }, () => { }));
            Assert.AreEqual(code, state.Record["ExitCode"]); Assert.AreEqual(false, state.Record["ProcessHandleRetained"]);
            Assert.AreEqual(false, state.Record["OwnershipRetained"]); Assert.AreEqual(false, state.Record["ForcedTermination"]);
            Assert.AreEqual(1, quits); Assert.AreEqual(1, codeReads); Assert.AreEqual(1, disposals);
        }

        [TestMethod]
        public void FailedQuitPreservesNativeErrorAndPreventsExitAcceptanceOrAnotherQuit()
        {
            var state = Create(); int quits = 0, waits = 0, disposals = 0;
            var original = new InvalidOperationException("original failed Quit"); state.Prepare(() => { });
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => state.QuitOnce(() => { quits++; throw original; }, () => { })));
            Assert.AreEqual("CALL_FAILED_EFFECT_UNKNOWN", state.Record["QuitOutcome"]);
            Assert.AreEqual(true, state.Record["ProcessHandleRetained"]);
            Assert.ThrowsException<InvalidOperationException>(() => state.ObserveExit(() => { waits++; return true; }, () => 0, () => disposals++, () => { }));
            Assert.ThrowsException<InvalidOperationException>(() => state.QuitOnce(() => quits++, () => { }));
            Assert.AreEqual(1, quits); Assert.AreEqual(0, waits); Assert.AreEqual(0, disposals);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void UnknownWaitOrUnavailableExitCodeKeepsOriginalHandleAndError(bool failCode)
        {
            var state = Create(); int disposals = 0;
            var original = new InvalidOperationException("original unknown exit observation");
            state.Prepare(() => { }); state.QuitOnce(() => { }, () => { });
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => state.ObserveExit(
                () => { if (!failCode) throw original; return true; }, () => { throw original; }, () => disposals++, () => { })));
            Assert.AreEqual(failCode, state.Record["ProcessExitObserved"]);
            Assert.AreEqual(false, state.Record["ExitCodeObserved"]); Assert.IsNull(state.Record["ExitCode"]);
            Assert.AreEqual(true, state.Record["OwnershipRetained"]); Assert.AreEqual(true, state.Record["ProcessHandleRetained"]);
            Assert.AreEqual(0, disposals);
        }

        [TestMethod]
        public void EvidenceFailureBeforeQuitEmitsNothingAndDualFailurePreservesFirstError()
        {
            var state = Create(); int quits = 0; var write = new IOException("preparation write failed");
            Assert.AreSame(write, Assert.ThrowsException<IOException>(() => state.Prepare(() => { throw write; })));
            Assert.AreEqual(0, state.Record["QuitEntries"]);
            Assert.ThrowsException<InvalidOperationException>(() => state.QuitOnce(() => quits++, () => { }));
            var other = Create(); var original = new InvalidOperationException("original Quit"); other.Prepare(() => { }); int writes = 0;
            var aggregate = Assert.ThrowsException<AggregateException>(() => other.QuitOnce(() => { quits++; throw original; },
                () => { if (++writes == 2) throw write; }));
            Assert.AreSame(original, aggregate.InnerExceptions[0]); Assert.AreSame(write, aggregate.InnerExceptions[1]);
            Assert.AreEqual(1, quits); Assert.AreEqual(true, other.Record["ProcessHandleRetained"]);
        }

        [TestMethod]
        public void ExitEvidenceWriteFailureDoesNotDisposeAnUnrecordedIdentityHandle()
        {
            var state = Create(); int disposals = 0; var original = new IOException("exit evidence failed");
            state.Prepare(() => { }); state.QuitOnce(() => { }, () => { });
            Assert.ThrowsException<AggregateException>(() => state.ObserveExit(() => true, () => 0, () => disposals++, () => { throw original; }));
            Assert.AreEqual(0, disposals); Assert.AreEqual(true, state.Record["ProcessHandleRetained"]);
            Assert.AreEqual(true, state.Record["ProcessExitObserved"]); Assert.AreEqual(0, state.Record["ExitCode"]);
        }
    }
}
