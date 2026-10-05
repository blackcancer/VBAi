using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureShutdownTests
    {
        [TestMethod]
        public void SuccessfulOriginalObservationReleasesOnceAndSecondDisposeDoesNothing()
        {
            WithFixture((fixture, process, calls) => {
                fixture.WaitForOwnedExcelExit = (actual, timeout) => { Assert.AreSame(process, actual); Assert.AreEqual(10000, timeout); calls.Add("wait"); return true; };
                fixture.ReadOwnedExcelExitCode = actual => { Assert.AreSame(process, actual); calls.Add("code"); return 0; };
                fixture.Dispose();
                CollectionAssert.AreEqual(new[] { "close", "quit", "wait", "code", "release" }, calls);
                Assert.IsNull(Get(fixture, "ownedProcess"));
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["ProcessHandleRetained"]);
                var receipt = fixture.ShutdownDiagnostics;
                fixture.Dispose(); Assert.AreSame(receipt, fixture.ShutdownDiagnostics);
                Assert.AreEqual(5, calls.Count);
            });
        }

        [DataTestMethod]
        [DataRow("timeout")]
        [DataRow("wait-error")]
        [DataRow("code-error")]
        [DataRow("abnormal")]
        [DataRow("release-error")]
        [DataRow("release-after-close")]
        public void FailedOriginalObservationRetainsSameDescriptorAndForbidsCleanupReplay(string fault)
        {
            WithFixture((fixture, process, calls) => {
                fixture.WaitForOwnedExcelExit = (actual, timeout) => {
                    Assert.AreSame(process, actual); Assert.AreEqual(10000, timeout); calls.Add("wait");
                    if (fault == "wait-error") throw new InvalidOperationException("wait primary");
                    return fault != "timeout";
                };
                fixture.ReadOwnedExcelExitCode = actual => { calls.Add("code"); if (fault == "code-error") throw new InvalidOperationException("code primary"); return fault == "abnormal" ? 7 : 0; };
                bool releaseFault = fault.StartsWith("release-", StringComparison.Ordinal);
                if (releaseFault) fixture.ReleaseOwnedExcelProcess = actual => {
                    calls.Add("release"); if (fault == "release-after-close") actual.Dispose();
                    throw new InvalidOperationException("release primary");
                };
                Exception failure = Capture(fixture.Dispose); Assert.IsNotNull(failure);
                Assert.AreSame(process, Get(fixture, "ownedProcess"));
                if (releaseFault) {
                    Assert.IsNull(fixture.ShutdownDiagnostics["ProcessHandleRetained"]);
                    Assert.AreEqual("UNKNOWN_AFTER_ENTRY", fixture.ShutdownDiagnostics["ProcessHandleReleaseOutcome"]);
                } else Assert.AreEqual(true, fixture.ShutdownDiagnostics["ProcessHandleRetained"]);
                Assert.AreEqual("FAILED_NO_REPLAY", fixture.ShutdownDiagnostics["CleanupState"]);
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["LaterObservationCanQualify"]);
                Assert.AreEqual(true, fixture.ShutdownDiagnostics["OriginalCleanupVerdictFinal"]);
                if (fault == "timeout") { Assert.AreEqual(false, fixture.ShutdownDiagnostics["Exited"]); Assert.IsFalse(calls.Contains("code")); Assert.IsFalse(calls.Contains("release")); }
                var receipt = fixture.ShutdownDiagnostics; int count = calls.Count;
                Assert.ThrowsException<InvalidOperationException>(fixture.Dispose);
                Assert.AreEqual(count, calls.Count); Assert.AreSame(receipt, fixture.ShutdownDiagnostics);
                Assert.AreSame(process, Get(fixture, "ownedProcess"));
                if (!releaseFault) Assert.AreEqual(process.Handle, ((Process)Get(fixture, "ownedProcess")).Handle);
            });
        }

        [DataTestMethod]
        [DataRow("close")]
        [DataRow("quit")]
        [DataRow("both")]
        public void NativeErrorsRemainFailuresAfterNormalObservationAndCannotRearmCleanup(string fault)
        {
            WithFixture((fixture, process, calls) => {
                ((FakeWorkbook)Get(fixture, "workbook")).Fail = fault != "quit";
                ((FakeApplication)Get(fixture, "application")).Fail = fault != "close";
                fixture.WaitForOwnedExcelExit = (actual, timeout) => { calls.Add("wait"); return true; };
                fixture.ReadOwnedExcelExitCode = actual => 0;
                var error = Assert.ThrowsException<AggregateException>(fixture.Dispose);
                Assert.AreEqual(fault == "both" ? 2 : 1, error.InnerExceptions.Count);
                int count = calls.Count; Assert.ThrowsException<InvalidOperationException>(fixture.Dispose);
                Assert.AreEqual(count, calls.Count);
            });
        }

        [TestMethod]
        public void WrongThreadRefusesBeforeAnyCleanupAndOriginalOwnerCanStillEnterOnce()
        {
            WithFixture((fixture, process, calls) => {
                Exception foreign = null;
                var thread = new Thread(() => foreign = Capture(fixture.Dispose)); thread.Start(); thread.Join();
                Assert.IsInstanceOfType(foreign, typeof(InvalidOperationException)); Assert.AreEqual(0, calls.Count);
                fixture.WaitForOwnedExcelExit = (actual, timeout) => true; fixture.ReadOwnedExcelExitCode = actual => 0;
                fixture.Dispose(); CollectionAssert.AreEqual(new[] { "close", "quit", "release" }, calls);
            });
        }

        [TestMethod]
        public void ReentrantDisposeIsRefusedBeforeSecondNativeEntry()
        {
            WithFixture((fixture, process, calls) => {
                ((FakeWorkbook)Get(fixture, "workbook")).DuringClose = () => Assert.ThrowsException<InvalidOperationException>(fixture.Dispose);
                fixture.WaitForOwnedExcelExit = (actual, timeout) => true; fixture.ReadOwnedExcelExitCode = actual => 0;
                fixture.Dispose(); CollectionAssert.AreEqual(new[] { "close", "quit", "release" }, calls);
            });
        }

        [TestMethod]
        public void SuspendedCleanupRetainsOriginalWithoutCloseQuitWaitReleaseOrReplay()
        {
            WithFixture((fixture, process, calls) => {
                fixture.PreserveForDiagnosticRecovery = true;
                Assert.ThrowsException<InvalidOperationException>(fixture.Dispose);
                Assert.AreEqual(0, calls.Count); Assert.AreSame(process, Get(fixture, "ownedProcess"));
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["ExitWaitAttempted"]);
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["ExitWaitReturned"]);
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["ExitCodeObserved"]);
                fixture.PreserveForDiagnosticRecovery = false;
                Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(0, calls.Count);
            });
        }

        [TestMethod]
        public void RecordsObservedCallerThreadAndApartmentBeforeAndAfterOriginalWait()
        {
            WithFixture((fixture, process, calls) => {
                fixture.WaitForOwnedExcelExit = (actual, timeout) => true; fixture.ReadOwnedExcelExitCode = actual => 0;
                fixture.Dispose();
                foreach (string phase in new[] { "CleanupStarted", "BeforeExitWait", "AfterExitWait" }) {
                    Assert.AreEqual(Thread.CurrentThread.ManagedThreadId, fixture.ShutdownDiagnostics[phase + "ManagedThreadId"]);
                    Assert.AreEqual(Thread.CurrentThread.GetApartmentState().ToString(), fixture.ShutdownDiagnostics[phase + "Apartment"]);
                    Assert.IsTrue((uint)fixture.ShutdownDiagnostics[phase + "NativeThreadId"] > 0);
                }
            });
        }

        [TestMethod]
        public void EvidencePublicationFailureRetainsOriginalAndNeverClaimsSuccessfulCleanup()
        {
            WithFixture((fixture, process, calls) => {
                fixture.WriteShutdownReceipt = diagnostics => { throw new IOException("receipt primary"); };
                fixture.WaitForOwnedExcelExit = (actual, timeout) => true; fixture.ReadOwnedExcelExitCode = actual => 0;
                Assert.IsNotNull(Capture(fixture.Dispose));
                Assert.AreSame(process, Get(fixture, "ownedProcess")); Assert.IsFalse(calls.Contains("release"));
                int count = calls.Count; Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(count, calls.Count);
            });
        }

        [DataTestMethod]
        [DataRow("scenario")]
        [DataRow("shutdown")]
        [DataRow("both")]
        public void PreparedScenarioPreservesIndependentErrorsAndDrainsOnlyOnce(string fault)
        {
            WithFixture((fixture, process, calls) => {
                var scenarioFailure = new InvalidOperationException("scenario primary");
                fixture.CollectScenarioReferences = () => calls.Add("drain");
                fixture.WaitForOwnedExcelExit = (actual, timeout) => { calls.Add("wait"); return fault == "scenario"; };
                fixture.ReadOwnedExcelExitCode = actual => 0;
                Exception error = Capture(() => ExcelVbeFixture.RunPreparedScenario(fixture, current => {
                    Assert.AreSame(fixture, current); calls.Add("scenario"); if (fault != "shutdown") throw scenarioFailure;
                }));
                Assert.IsNotNull(error);
                if (fault == "scenario") Assert.AreSame(scenarioFailure, error);
                if (fault == "both") {
                    var aggregate = (AggregateException)error; Assert.AreEqual(2, aggregate.InnerExceptions.Count);
                    Assert.AreSame(scenarioFailure, aggregate.InnerExceptions[0]);
                }
                CollectionAssert.AreEqual(fault == "scenario" ?
                    new[] { "scenario", "drain", "close", "quit", "wait", "release" } :
                    new[] { "scenario", "drain", "close", "quit", "wait" }, calls);
                Assert.AreEqual(Thread.CurrentThread.GetApartmentState().ToString(), fixture.ShutdownDiagnostics["ScenarioReturnedApartment"]);
                int count = calls.Count;
                if (fault == "scenario") fixture.Dispose(); else Assert.ThrowsException<InvalidOperationException>(fixture.Dispose);
                Assert.AreEqual(count, calls.Count);
            });
        }

        [DataTestMethod]
        [DataRow("missing")]
        [DataRow("foreign-pid")]
        public void MissingOrForeignOriginalOwnershipRefusesBeforeCloseQuitOrWait(string fault)
        {
            WithFixture((fixture, process, calls) => {
                if (fault == "missing") Set(fixture, "ownedProcess", null); else Set(fixture, "ProcessId", process.Id + 1);
                Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(0, calls.Count);
                Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(0, calls.Count);
            });
        }

        [TestMethod]
        public void PreparedScenarioRejectsInvalidDependenciesBeforeAnyCleanup()
        {
            WithFixture((fixture, process, calls) => {
                Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.RunPreparedScenario(null, current => { }));
                Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.RunPreparedScenario(fixture, null));
                Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.Run(null));
                Assert.AreEqual(0, calls.Count);
            });
        }

        [DataTestMethod]
        [DataRow("close-wait")]
        [DataRow("quit-code")]
        [DataRow("both-timeout")]
        public void ObservationFailuresPreserveEarlierCloseAndQuitErrors(string fault)
        {
            WithFixture((fixture, process, calls) => {
                ((FakeWorkbook)Get(fixture, "workbook")).Fail = fault != "quit-code";
                ((FakeApplication)Get(fixture, "application")).Fail = fault != "close-wait";
                fixture.WaitForOwnedExcelExit = (actual, timeout) => {
                    if (fault == "close-wait") throw new InvalidOperationException("wait primary");
                    return fault != "both-timeout";
                };
                fixture.ReadOwnedExcelExitCode = actual => { throw new InvalidOperationException("code primary"); };
                var error = Assert.ThrowsException<AggregateException>(fixture.Dispose);
                Assert.AreEqual(fault == "both-timeout" ? 3 : 2, error.Flatten().InnerExceptions.Count);
                StringAssert.Contains(error.ToString(), fault == "quit-code" ? "quit primary" : "close primary");
                if (fault == "both-timeout") StringAssert.Contains(error.ToString(), "quit primary");
                Assert.AreSame(process, Get(fixture, "ownedProcess"));
                Assert.AreEqual(true, fixture.ShutdownDiagnostics["ExitWaitAttempted"]);
                Assert.AreEqual(fault != "close-wait", fixture.ShutdownDiagnostics["ExitWaitReturned"]);
                Assert.AreEqual(false, fixture.ShutdownDiagnostics["ExitCodeObserved"]);
                int count=calls.Count; Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(count,calls.Count);
            });
        }

        [TestMethod]
        public void DuplicateObservationAggregateCannotHideAnEarlierNativeFailure()
        {
            WithFixture((fixture, process, calls) => {
                ((FakeWorkbook)Get(fixture, "workbook")).Fail = true;
                var waitFailure = new InvalidOperationException("wait unique");
                fixture.WaitForOwnedExcelExit = (actual, timeout) => { throw new AggregateException(waitFailure, waitFailure); };
                var failure = Assert.ThrowsException<AggregateException>(fixture.Dispose);
                Assert.AreEqual(2, failure.InnerExceptions.Count);
                StringAssert.Contains(failure.ToString(), "close primary");
                Assert.AreSame(waitFailure, failure.InnerExceptions[1]);
                Assert.AreSame(process, Get(fixture, "ownedProcess"));
            });
        }

        private static Exception Capture(Action action) { try { action(); return null; } catch (Exception error) { return error; } }
        private static object Get(object instance, string name) { return typeof(ExcelVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(instance); }
        private static void Set(object instance, string name, object value) {
            var field = typeof(ExcelVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
            if (field != null) field.SetValue(instance, value);
            else typeof(ExcelVbeFixture).GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(instance, value, null);
        }
        private static void WithFixture(Action<ExcelVbeFixture, Process, List<string>> action)
        {
            var fixture = (ExcelVbeFixture)Activator.CreateInstance(typeof(ExcelVbeFixture), true);
            var calls = new List<string>();
            using (var process = Process.GetCurrentProcess()) {
                // This descriptor is read only; no real application or process is closed by this mirror.
                Set(fixture, "ownedProcess", process); Set(fixture, "owned", true); Set(fixture, "ProcessId", process.Id);
                Set(fixture, "application", new FakeApplication(calls)); Set(fixture, "workbook", new FakeWorkbook(calls));
                Set(fixture, "retainEvidence", true);
                fixture.ReleaseOwnedExcelProcess = actual => { Assert.AreSame(process, actual); calls.Add("release"); };
                try { action(fixture, process, calls); }
                finally {
                    var retained = (IList)typeof(ExcelVbeFixture).GetField("retainedBootstraps", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                    lock (retained) retained.Remove(fixture);
                }
            }
        }
        public sealed class FakeWorkbook {
            private readonly List<string> calls; public bool Fail; public Action DuringClose;
            public FakeWorkbook(List<string> log) { calls = log; }
            public void Close(bool save) { calls.Add("close"); Assert.IsFalse(save); DuringClose?.Invoke(); if (Fail) throw new InvalidOperationException("close primary"); }
        }
        public sealed class FakeApplication {
            private readonly List<string> calls; public bool Fail;
            public FakeApplication(List<string> log) { calls = log; }
            public void Quit() { calls.Add("quit"); if (Fail) throw new InvalidOperationException("quit primary"); }
        }
    }
}
