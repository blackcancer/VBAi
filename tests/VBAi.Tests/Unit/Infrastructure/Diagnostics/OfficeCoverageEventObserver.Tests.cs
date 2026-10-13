using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the lifecycle oracle without COM activation or opening native Office applications.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeCoverageEventObserverTests
    {
        private static readonly string Original = Path.Combine(Path.GetTempPath(), "OracleOriginal.docm");
        private static readonly string Copy = Path.Combine(Path.GetTempPath(), "OracleCopy.docm");
        [TestMethod]
        public void ExactlyOneOwnedOpenAndCloseIsAcceptedWithoutChangingTheTrace()
        {
            var trace = Completed(); trace.RequireCompletedCopy(Copy);
            Assert.AreEqual(2, trace.Events.Count); Assert.AreEqual(0, trace.Failures.Count);
            trace.RequireCompletedCopy(Copy); Assert.AreEqual(2, trace.Events.Count);
        }
        [DataTestMethod, DataRow("Original"), DataRow("Missing"), DataRow("Extra"), DataRow("Reversed"),
            DataRow("Thread"), DataRow("Cancelled"), DataRow("OtherPath"), DataRow("NoPath"), DataRow("Error")]
        public void MissingMisorderedOrForeignLifecycleEvidenceCannotQualifyACopy(string fault)
        {
            var trace = Completed();
            switch (fault)
            {
                case "Original": trace.Events[0].Path=Original; break;
                case "Missing": trace.Events.RemoveAt(1); break;
                case "Extra": trace.Record("Open", Copy, false, 7); break;
                case "Reversed": trace.Events.Reverse(); break;
                case "Thread": trace.Events[1].Thread=8; break;
                case "Cancelled": trace.Events[1].Cancel=true; break;
                case "OtherPath": trace.Events[1].Path=Copy + ".other"; break;
                case "NoPath": trace.Events[0].Path=null; break;
                case "Error": trace.Failures.Add("Unclassified callback"); break;
                default: Assert.Fail("Unknown scenario"); break;
            }
            Assert.ThrowsException<AssertFailedException>(() => trace.RequireCompletedCopy(Copy));
        }
        [TestMethod]
        public void OriginalAsCopyAndEventOverflowAreRefusedWithBoundedEvidence()
        {
            Assert.ThrowsException<AssertFailedException>(() => Completed().RequireCompletedCopy(Original));
            var trace = Completed();
            for (int i=0; i<100; i++) trace.Record("Open", Copy, false, 7);
            Assert.AreEqual(32, trace.Events.Count); Assert.AreEqual(1, trace.Failures.Count);
            Assert.ThrowsException<AssertFailedException>(() => trace.RequireCompletedCopy(Copy));
        }
        [TestMethod]
        public void PumpedResponseSendsOneNetworkOperationAndRunsCallbacksOnTheCallingThread()
        {
            int owner=Thread.CurrentThread.ManagedThreadId, sends=0, pumps=0;
            using (var callback=new ManualResetEventSlim())
            {
                string result=OfficeCoverageEventObserver.ReadPumping(() =>
                {
                    Assert.AreNotEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    Interlocked.Increment(ref sends);
                    if (!callback.Wait(10000)) throw new TimeoutException("The owner did not pump its callback.");
                    return "terminal-response";
                }, () => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); pumps++; callback.Set(); }, Thread.Sleep);
                Assert.AreEqual("terminal-response", result); Assert.AreEqual(1, sends); Assert.IsTrue(pumps>0);
            }
        }
        [TestMethod]
        public void TransportFailureIsReturnedUnwrappedWithoutAnotherRequest()
        {
            int sends=0; var expected=new IOException("original transport failure");
            var actual=Assert.ThrowsException<IOException>(() => OfficeCoverageEventObserver.ReadPumping<string>(() =>
                { Interlocked.Increment(ref sends); throw expected; }, () => { }, Thread.Sleep));
            Assert.AreSame(expected, actual); Assert.AreEqual(1, sends);
        }
        [DataTestMethod, DataRow("Word"), DataRow("PowerPoint")]
        public void CompleteNativeAttachmentUsesPublishedIdsAndDisposesEachOwnedBoundaryOnce(string host)
        {
            var probe = new LeaseProbe(host);
            var original = probe.Fixture.Dispatch;
            var observer = new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary);
            CollectionAssert.AreEqual(host == "Word" ? new[] { 4, 6 } : new[] { 2006, 2025 }, probe.Added.ToArray());
            Assert.AreNotSame(original, probe.Fixture.Dispatch);
            observer.Dispose(); observer.Dispose();
            Assert.AreSame(original, probe.Fixture.Dispatch);
            CollectionAssert.AreEqual(probe.Added.ToArray(), probe.Removed.ToArray());
            Assert.AreEqual(1, probe.Released);
            Assert.IsFalse(probe.Fixture.NativeExecutionUnsettled);
        }

        [DataTestMethod, DataRow("dispatcher"), DataRow("acquire"), DataRow("window"), DataRow("zero"), DataRow("owner"),
            DataRow("path"), DataRow("addOpen"), DataRow("addClose")]
        public void ConstructorRefusalBalancesApplicationAndAttemptedSubscriptions(string fault)
        {
            var probe = new LeaseProbe("PowerPoint") { Fault = fault };
            if (fault == "path") Set(probe.Fixture, "DocumentPath", "\0");
            var original = probe.Fixture.Dispatch;
            Exception failure = Capture(() => new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary));
            Assert.IsNotNull(failure);
            Assert.AreSame(original, probe.Fixture.Dispatch);
            CollectionAssert.AreEquivalent(probe.Added.ToArray(), probe.Removed.ToArray());
            Assert.AreEqual(fault == "acquire" || fault == "dispatcher" ? 0 : 1, probe.Released);
            Assert.AreEqual(fault == "dispatcher" ? 0 : 1, probe.DispatcherDisposed);
            Assert.IsFalse(probe.Fixture.NativeExecutionUnsettled);
        }

        [DataTestMethod, DataRow("remove"), DataRow("release"), DataRow("dispatcher")]
        public void ConstructorRecoveryFailureRetainsTheHostAndPreservesBothErrors(string fault)
        {
            var probe = new LeaseProbe("Word") { Fault = "addClose", CleanupFault = fault };
            var failure = Assert.ThrowsException<AggregateException>(() => new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary));
            Assert.AreEqual(2, failure.InnerExceptions.Count);
            Assert.IsTrue(probe.Fixture.NativeExecutionUnsettled);
            Assert.AreEqual(fault == "remove" ? 0 : 1, probe.Released);
        }

        [TestMethod]
        public void UnsettledExecutionRetainsTheSinkAndWrongThreadCannotConsumeDisposal()
        {
            var probe = new LeaseProbe("Word");
            var observer = new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary);
            Exception refusal = null;
            var thread = new Thread(() => { try { observer.Dispose(); } catch (Exception error) { refusal = error; } });
            thread.Start(); thread.Join();
            Assert.IsInstanceOfType(refusal, typeof(AssertFailedException));
            Assert.AreEqual(0, probe.Removed.Count + probe.Released);
            probe.Fixture.NativeExecutionUnsettled = true;
            var dispatch = probe.Fixture.Dispatch;
            observer.Dispose(); observer.Dispose();
            Assert.AreSame(dispatch, probe.Fixture.Dispatch);
            Assert.AreEqual(0, probe.Removed.Count + probe.Released);
        }

        [DataTestMethod, DataRow("dispatch"), DataRow("remove"), DataRow("release"), DataRow("dispatcher")]
        public void DisposalFailureRetainsReferencesAndCannotRepeatAnUnsubscribeOrRelease(string fault)
        {
            var probe = new LeaseProbe("PowerPoint");
            var observer = new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary);
            if (fault == "dispatch") probe.Fixture.Dispatch = (pid, request) => null;
            else probe.CleanupFault = fault;
            Assert.IsNotNull(Capture(observer.Dispose));
            Assert.IsTrue(probe.Fixture.NativeExecutionUnsettled);
            int entries = probe.Removed.Count + probe.Released;
            observer.Dispose();
            Assert.AreEqual(entries, probe.Removed.Count + probe.Released);
            Assert.AreEqual(fault == "release" || fault == "dispatcher" ? 1 : 0, probe.Released);
        }

        [DataTestMethod, DataRow("Word", false), DataRow("PowerPoint", false), DataRow("Word", true)]
        public void RpcCallbacksReadBorrowedPathsOnTheOwnerBeforeReturningAndKeepCancelUnchanged(string host, bool cancelled)
        {
            OfficeCoverageEventDispatcherTests.OnSta(() =>
            {
                var probe = new LeaseProbe(host);
                probe.Boundary.CreateDispatcher = () => new OfficeCoverageEventDispatcher();
                using (var observer = new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary))
                {
                    int owner = Thread.CurrentThread.ManagedThreadId;
                    var document = new BorrowedDocument { Owner = owner, Path = Copy };
                    OfficeCoverageEventObserver.ReadPumping(() =>
                    {
                        probe.Callbacks[0].DynamicInvoke(document);
                        Assert.AreEqual(1, document.Reads, "The borrowed argument must be read before Open returns.");
                        var args = new object[] { document, cancelled };
                        probe.Callbacks[1].DynamicInvoke(args);
                        Assert.AreEqual(cancelled, args[1]);
                        Assert.AreEqual(2, document.Reads, "BeforeClose must not defer access to the closed argument.");
                        return true;
                    }, System.Windows.Forms.Application.DoEvents, Thread.Sleep);
                    Assert.AreEqual(2, observer.Trace.Events.Count);
                    Assert.IsTrue(observer.Trace.Events.All(item => item.Thread == owner && item.DeliveryThread != owner));
                    if (cancelled) Assert.ThrowsException<AssertFailedException>(() => observer.Trace.RequireCompletedCopy(Copy));
                    else observer.Trace.RequireCompletedCopy(Copy);
                }
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void CallbackFailureIsRecordedWithoutThrowingIntoOfficeOrChangingCancel(bool dispatchFailure)
        {
            var probe = new LeaseProbe("Word");
            using (var observer = new OfficeCoverageEventObserver(probe.Fixture, boundary: probe.Boundary))
            {
                if (dispatchFailure) probe.Fault = "callbackDispatch";
                var args = new object[] { new BorrowedDocument { Owner = -1, Path = Copy }, true };
                probe.Callbacks[1].DynamicInvoke(args);
                Assert.AreEqual(true, args[1]);
                Assert.AreEqual(1, observer.Trace.Failures.Count);
                Assert.AreEqual(0, observer.Trace.Events.Count);
            }
        }

        /// <summary>A borrowed argument which rejects getters on any thread except the fixture's owner.</summary>
        public sealed class BorrowedDocument
        {
            public int Owner, Reads;
            public string Path;
            public string FullName
            {
                get
                {
                    Assert.AreEqual(Owner, Thread.CurrentThread.ManagedThreadId);
                    Reads++;
                    return Path;
                }
            }
        }

        private static Exception Capture(Action action)
        {
            try { action(); return null; } catch (Exception error) { return error; }
        }

        private static void Set(OfficeVbeFixture fixture, string property, object value) =>
            typeof(OfficeVbeFixture).GetField("<" + property + ">k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, value);

        private sealed class LeaseProbe
        {
            internal readonly OfficeVbeFixture Fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            internal readonly OfficeCoverageEventObserver.EventBoundary Boundary = new OfficeCoverageEventObserver.EventBoundary();
            internal readonly List<int> Added = new List<int>(), Removed = new List<int>();
            internal readonly List<Delegate> Callbacks = new List<Delegate>();
            internal string Fault, CleanupFault;
            internal int Released, DispatcherDisposed;
            internal LeaseProbe(string host)
            {
                Set(Fixture, "Kind", host); Set(Fixture, "ProcessId", 123); Set(Fixture, "DocumentPath", Original);
                var application = new object();
                Boundary.CreateDispatcher = () =>
                {
                    if (Fault == "dispatcher") throw new IOException("dispatcher construction");
                    return new InlineDispatcher(this);
                };
                Boundary.Acquire = progId => { Assert.AreEqual(host + ".Application", progId); if (Fault == "acquire") throw new IOException("acquire"); return application; };
                Boundary.Window = (kind, value) => { Assert.AreSame(application, value); if (Fault == "window") throw new IOException("window"); return Fault == "zero" ? IntPtr.Zero : new IntPtr(99); };
                Boundary.Owner = window => Fault == "owner" ? 777u : 123u;
                Boundary.Add = (value, events, id, callback) =>
                {
                    Assert.AreSame(application, value); Added.Add(id); Callbacks.Add(callback);
                    if (Fault == (Added.Count == 1 ? "addOpen" : "addClose")) throw new IOException("add");
                };
                Boundary.Remove = (value, events, id, callback) =>
                {
                    Assert.AreSame(application, value); Removed.Add(id);
                    if (CleanupFault == "remove") { CleanupFault = null; throw new IOException("remove"); }
                };
                Boundary.Release = value => { Assert.AreSame(application, value); Released++; if (CleanupFault == "release") throw new IOException("release"); };
            }
            private sealed class InlineDispatcher : IOfficeCoverageEventDispatcher
            {
                private readonly LeaseProbe probe;
                internal InlineDispatcher(LeaseProbe probe) { this.probe = probe; }
                public void Invoke(Action callback)
                {
                    if (probe.Fault == "callbackDispatch") throw new IOException("callback dispatch");
                    callback();
                }
                public void Dispose()
                {
                    probe.DispatcherDisposed++;
                    if (probe.CleanupFault == "dispatcher") throw new IOException("dispatcher disposal");
                }
            }
        }

        private static OfficeCoverageEventTrace Completed()
        {
            var trace = new OfficeCoverageEventTrace(Original, 7);
            trace.Record("Open", Copy, false, 7); trace.Record("BeforeClose", Copy, false, 7);
            return trace;
        }
    }
}
