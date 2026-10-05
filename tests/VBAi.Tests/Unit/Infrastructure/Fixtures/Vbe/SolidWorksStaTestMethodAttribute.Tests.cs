using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises real save scheduling against managed probes only; no SOLIDWORKS process or COM activation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SolidWorksStaTestMethodAttributeTests
    {
        [DataTestMethod]
        [DataRow(false, false)] [DataRow(false, true)]
        [DataRow(true, false)] [DataRow(true, true)]
        public void PendingRealAdapterStaysOnItsParkedStaAndTheNextHostSaveIsIndependent(bool directAdapter, bool throws)
        {
            Task<object> pending = null;
            int oldSaves = 0, lateCallbacks = 0;
            var original = new AssertFailedException("Original synthetic scenario failure; no late acceptance");
            var retained = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                var probe = new VbeSolidWorksPersistenceTests.Probe();
                probe.OnSave = () => { oldSaves++; probe.Project.Saved = false; };
                pending = directAdapter ? probe.SaveAdapterAsync(probe.Request()) : probe.SaveHostAsync(probe.Request());
                Assert.IsFalse(pending.IsCompleted);
                var otherSession = new VbeSolidWorksPersistenceTests.Probe();
                Assert.AreEqual(VbeSolidWorksPersistenceTests.PendingSaveMessage,
                    Assert.ThrowsException<InvalidOperationException>(() =>
                        VbeSolidWorksPersistenceTests.Complete(otherSession.SaveHostAsync(otherSession.Request()))).Message);
                Assert.AreEqual(0, otherSession.Saves);
                SynchronizationContext.Current.Post(_ => Interlocked.Increment(ref lateCallbacks), null);
                if (throws) throw original;
                return 19;
            });
            Assert.IsTrue(retained.Retained); Assert.IsFalse(retained.DispatcherDisposed); Assert.IsFalse(retained.ThreadExitObserved);
            if (throws) Assert.AreSame(original, retained.Error); else { Assert.IsNull(retained.Error); Assert.AreEqual(19, retained.Value); }
            StringAssert.Contains(retained.Diagnostic, "TraceBound=128");
            StringAssert.Contains(retained.Diagnostic, "OwnerSta");
            StringAssert.Contains(retained.Diagnostic, "LateCompletionAccepted=false");
            StringAssert.Contains(retained.Diagnostic, "SaveReplay=0");
            var next = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                var probe = new VbeSolidWorksPersistenceTests.Probe();
                dynamic result = VbeSolidWorksPersistenceTests.Complete(directAdapter ?
                    probe.SaveHostAsync(probe.Request()) : probe.SaveAdapterAsync(probe.Request()));
                Assert.IsTrue((bool)result.Verified); Assert.AreEqual(1, probe.Saves);
                return Thread.CurrentThread.ManagedThreadId;
            });
            Assert.IsNull(next.Error, next.Error?.ToString()); Assert.IsFalse(next.Retained); Assert.IsTrue(next.ThreadExitObserved);
            Assert.AreNotEqual(retained.OwnerThread, next.OwnerThread);
            Assert.IsFalse(pending.IsCompleted, "Another STA must neither drain nor finish the retained original save.");
            Assert.AreEqual(1, Volatile.Read(ref oldSaves)); Assert.AreEqual(0, Volatile.Read(ref lateCallbacks));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void TerminalSaveAndUncertainReadbackKeepTheirOriginalOutcomeAndPreAdmissionTrace(bool changeIdentity)
        {
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                var probe = new VbeSolidWorksPersistenceTests.Probe();
                int owner = Thread.CurrentThread.ManagedThreadId;
                probe.OnSave = () => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); if (changeIdentity) probe.Identity = false; };
                var task = probe.SaveHostAsync(probe.Request());
                dynamic result = VbeSolidWorksPersistenceTests.Complete(task);
                Assert.AreEqual(!changeIdentity, (bool)result.Verified);
                Assert.AreEqual(changeIdentity, (bool)result.Uncertain); Assert.AreEqual(1, probe.Saves);
                string trace = SolidWorksStaTestMethodAttribute.Describe(task);
                foreach (string phase in new[] { "OwnerSta", "ContinuationEnqueued", "ContinuationEntered", "ContinuationReturned" })
                    StringAssert.Contains(trace, phase);
                StringAssert.Contains(trace, "TraceBound=128");
                return (object)result;
            });
            Assert.IsNull(run.Error, run.Error?.ToString()); Assert.IsFalse(run.Retained);
            Assert.IsTrue(run.DispatcherDisposed); Assert.IsTrue(run.ThreadExitObserved);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void AttributeExecutionPreservesDataRowsAndOriginalFailureButNeverPassesPendingWork(bool originallyFailed)
        {
            var failure = new AssertFailedException("Original five-second outcome");
            var first = new TestResult { DisplayName = "save (identity)", Outcome = originallyFailed ? UnitTestOutcome.Failed : UnitTestOutcome.Passed,
                TestFailureException = originallyFailed ? failure : null, LogOutput = "original details" };
            var second = new TestResult { DisplayName = "save (another row)", Outcome = UnitTestOutcome.Failed, TestFailureException = failure };
            var rows = new[] { first, second };
            var result = SolidWorksStaTestMethodAttribute.ExecuteScenario(() => {
                var probe = new VbeSolidWorksPersistenceTests.Probe();
                probe.OnSave = () => probe.Project.Saved = false;
                var task = probe.SaveHostAsync(probe.Request());
                Assert.IsFalse(task.IsCompleted); Assert.AreEqual(1, probe.Saves);
                return rows;
            });
            Assert.AreSame(rows, result); Assert.AreSame(first, result[0]); Assert.AreSame(second, result[1]);
            Assert.AreEqual(UnitTestOutcome.Failed, first.Outcome); Assert.AreSame(failure, second.TestFailureException);
            if (originallyFailed) Assert.AreSame(failure, first.TestFailureException);
            else Assert.IsInstanceOfType(first.TestFailureException, typeof(AssertFailedException));
            Assert.AreEqual("save (identity)", first.DisplayName); Assert.AreEqual("save (another row)", second.DisplayName);
            StringAssert.Contains(first.LogOutput, "original details"); StringAssert.Contains(first.LogOutput, "LateCompletionAccepted=false");
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void TerminalAttributeExecutionPreservesRowsOrTheExactThrownException(bool throws)
        {
            var original = new InvalidOperationException("Original invocation error");
            var rows = new[] { new TestResult { Outcome = UnitTestOutcome.Passed, DisplayName = "unchanged row" } };
            Func<TestResult[]> invoke = () => {
                var probe = new VbeSolidWorksPersistenceTests.Probe();
                VbeSolidWorksPersistenceTests.Complete(probe.SaveHostAsync(probe.Request()));
                if (throws) throw original;
                return rows;
            };
            if (throws) Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => SolidWorksStaTestMethodAttribute.ExecuteScenario(invoke)));
            else Assert.AreSame(rows, SolidWorksStaTestMethodAttribute.ExecuteScenario(invoke));
        }

        [TestMethod]
        public void ShimRefusesSaveAdmissionOutsideAnOwnedTestScopeBeforeCallingTheAdapter()
        {
            int attempts = 0;
            Assert.ThrowsException<InvalidOperationException>(() => SolidWorksStaTestMethodAttribute.StartSave(() => {
                attempts++; return Task.FromResult<object>(null);
            }));
            Assert.AreEqual(0, attempts);
        }
    }
}
