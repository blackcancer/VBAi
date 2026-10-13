using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using VBAi;

namespace VBAi.Tests.Unit
{
    /// <summary>Native-attempt ownership scenarios complement the dispatch and focus mirrors.</summary>
    public sealed partial class VbaNativeTestExecutionHostTests
    {
        private sealed class PreparedResources : VbaNativeTestExecutionHost.IPreparedResources
        {
            internal int Disposals, Retentions;
            internal bool FailDispose;
            public void Dispose() { Disposals++; if (FailDispose) throw new InvalidOperationException("Owned release failed."); }
            public void Retain() { Retentions++; }
        }

        [TestMethod]
        public void EveryCompletedNativeAttemptReleasesItsOwnPreparationExactlyOnce()
        {
            using (var f = new Fixture())
            {
                var attempts = new List<PreparedResources>();
                for (int i = 0; i < 3; i++)
                {
                    var prepared = new PreparedResources(); attempts.Add(prepared); f.Native.Prepared = prepared;
                    Assert.AreEqual(VbaTestOutcome.Passed, f.Start().GetAwaiter().GetResult().Outcome);
                    Assert.AreEqual(1, prepared.Disposals); Assert.AreEqual(0, prepared.Retentions);
                }
                f.Host.Dispose();
                Assert.IsTrue(attempts.All(p => p.Disposals == 1 && p.Retentions == 0));
                Assert.AreEqual(3, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("prepare", false)]
        [DataRow("authority", false)]
        [DataRow("revalidate", false)]
        [DataRow("guard", false)]
        [DataRow("execute", true)]
        [DataRow("observe", true)]
        public void RefusedAttemptReleasesUndispatchedPreparationAndRetainsDispatchedPreparation(string stage, bool dispatched)
        {
            using (var f = new Fixture())
            {
                var prepared = new PreparedResources(); f.Native.Prepared = prepared;
                Action fail = () => { throw new InvalidOperationException("Ownership refusal at " + stage); };
                if (stage == "prepare") f.Native.OnPrepare = fail;
                if (stage == "authority") f.ValidateAction = () => { if (f.Guards == 3) fail(); };
                if (stage == "revalidate") f.Native.OnRevalidate = fail;
                if (stage == "guard") f.GuardAction = () => { if (f.Guards == 4) fail(); };
                if (stage == "execute") f.Native.OnExecute = fail;
                if (stage == "observe") f.Native.OnReadMode = fail;
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                Assert.AreEqual(dispatched, error.Uncertain);
                Assert.AreEqual(stage == "prepare" || dispatched ? 0 : 1, prepared.Disposals);
                Assert.AreEqual(dispatched ? 1 : 0, prepared.Retentions);
                f.Host.Dispose();
                Assert.AreEqual(stage == "prepare" || dispatched ? 0 : 1, prepared.Disposals);
                Assert.AreEqual(dispatched ? 1 : 0, prepared.Retentions);
                Assert.AreEqual(dispatched ? 1 : 0, f.Native.Invocations);
            }
        }

        [DataTestMethod]
        [DataRow("timeout")]
        [DataRow("disconnect")]
        [DataRow("break")]
        public void UnsettledNativeAttemptRetainsPreparationWithoutAnotherExecute(string failure)
        {
            using (var f = new Fixture())
            {
                var prepared = new PreparedResources(); f.Native.Prepared = prepared;
                f.Native.Mode = 0;
                if (failure == "timeout") f.Host.VerificationTimeout = TimeSpan.Zero;
                var task = f.Start();
                if (failure == "disconnect") f.Host.Dispose();
                if (failure == "break") { f.Native.Mode = 1; f.Tick(); }
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => task.GetAwaiter().GetResult()).Uncertain);
                Assert.AreEqual(0, prepared.Disposals); Assert.AreEqual(1, prepared.Retentions);
                Assert.AreEqual(1, f.Native.Invocations);
                f.Host.Dispose();
                Assert.AreEqual(1, prepared.Retentions);
                Assert.IsNull(f.Tick);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void PreparationReleaseFailureSettlesTheTaskAndIsNeverRetried(bool nativeCompleted)
        {
            using (var f = new Fixture())
            {
                var prepared = new PreparedResources { FailDispose = true }; f.Native.Prepared = prepared;
                if (!nativeCompleted) f.Native.OnRevalidate = () => { throw new InvalidOperationException("Before dispatch."); };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Start().GetAwaiter().GetResult());
                StringAssert.Contains(error.Message, "Owned release failed");
                if (!nativeCompleted)
                {
                    StringAssert.Contains(error.Message, "Before dispatch");
                    var failures = ((AggregateException)error.InnerException).Flatten().InnerExceptions;
                    Assert.AreEqual(2, failures.Count);
                    Assert.IsTrue(failures.Any(failure => failure.Message.Contains("Owned release failed")));
                    Assert.IsTrue(failures.Any(failure => failure.Message.Contains("Before dispatch")));
                }
                Assert.AreEqual(nativeCompleted, error.Uncertain);
                Assert.AreEqual(1, prepared.Disposals); Assert.AreEqual(0, prepared.Retentions);
                f.Host.Dispose();
                Assert.AreEqual(1, prepared.Disposals);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void PreparationResourcesRemainOwnedUntilDelayedPositiveCompletion()
        {
            using (var f = new Fixture())
            {
                var prepared = new PreparedResources(); f.Native.Prepared = prepared; f.Native.Mode = 0;
                var task = f.Start();
                Assert.IsFalse(task.IsCompleted);
                Assert.AreEqual(0, prepared.Disposals); Assert.AreEqual(0, prepared.Retentions);
                f.Native.Mode = 2; f.Tick();
                Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(1, prepared.Disposals); Assert.AreEqual(0, prepared.Retentions);
            }
        }

        [TestMethod]
        public void NativePreparationReleasesTheTransferredPaneAndModuleOnlyAfterSettling()
        {
            RunOnSta(() =>
            {
                var released = new List<object>(); var f = new NativeFixture(released.Add);
                var prepared = (VbaNativeTestExecutionHost.IPreparedResources)f.Prepare();
                Assert.IsFalse(released.Contains(f.Editor.CommandBars.Control));
                int moduleBefore = released.Count(p => ReferenceEquals(p, f.Module));
                int paneBefore = released.Count(p => ReferenceEquals(p, f.Module.CodePane));
                prepared.Dispose();
                Assert.AreEqual(moduleBefore + 1, released.Count(p => ReferenceEquals(p, f.Module)));
                Assert.AreEqual(paneBefore + 1, released.Count(p => ReferenceEquals(p, f.Module.CodePane)));
                Assert.AreEqual(1, released.Count(p => ReferenceEquals(p, f.Editor.CommandBars.Control)));
                Assert.AreEqual(1, released.Count(p => ReferenceEquals(p, f.Editor.CommandBars)));
                int total = released.Count; prepared.Dispose(); Assert.AreEqual(total, released.Count);
            });
        }

        [TestMethod]
        public void UncertainNativePreparationRemainsRootedWithoutReleasingItsHeldInterfaces()
        {
            RunOnSta(() =>
            {
                var released = new List<object>(); var f = new NativeFixture(released.Add);
                var prepared = (VbaNativeTestExecutionHost.IPreparedResources)f.Prepare();
                int count = released.Count; prepared.Retain(); prepared.Dispose(); prepared.Retain();
                Assert.AreEqual(count, released.Count);
                Assert.IsFalse(released.Contains(f.Editor.CommandBars.Control));
            });
        }

        [DataTestMethod]
        [DataRow("released")]
        [DataRow("retained")]
        [DataRow("release-failed")]
        public void SettledNativePreparationCannotDispatchAgain(string settlement)
        {
            RunOnSta(() =>
            {
                var released = new List<object>(); var f = new NativeFixture(value =>
                {
                    released.Add(value);
                    if (settlement == "release-failed" && value is RunControl) throw new InvalidOperationException("Release failed.");
                });
                var prepared = (VbaNativeTestExecutionHost.IPreparedResources)f.Prepare(); int executed = 0;
                f.Editor.CommandBars.Control.OnExecute = () => executed++;
                if (settlement == "retained") prepared.Retain();
                else if (settlement == "release-failed") Assert.ThrowsException<InvalidOperationException>(prepared.Dispose);
                else prepared.Dispose();
                Assert.ThrowsException<ObjectDisposedException>(() => f.Probe.Revalidate(f.Editor, f.Project, prepared));
                Assert.ThrowsException<ObjectDisposedException>(() => f.Probe.Execute(prepared));
                int count = released.Count; prepared.Dispose(); prepared.Retain();
                Assert.AreEqual(count, released.Count); Assert.AreEqual(0, executed);
            });
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NativePreparationCannotReleaseOrRetainInterfacesOnAnotherThread(bool retain)
        {
            RunOnSta(() =>
            {
                var released = new List<object>(); var f = new NativeFixture(released.Add);
                var prepared = (VbaNativeTestExecutionHost.IPreparedResources)f.Prepare(); int before = released.Count;
                Exception failure = null;
                var thread = new System.Threading.Thread(() =>
                {
                    try { if (retain) prepared.Retain(); else prepared.Dispose(); }
                    catch (Exception error) { failure = error; }
                });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType<InvalidOperationException>(failure);
                Assert.AreEqual(before, released.Count);
                prepared.Dispose();
                Assert.AreEqual(1, released.Count(p => ReferenceEquals(p, f.Editor.CommandBars.Control)));
            });
        }

        [DataTestMethod]
        [DataRow("missing")]
        [DataRow("duplicate")]
        [DataRow("component-type")]
        [DataRow("source")]
        [DataRow("wrapper")]
        [DataRow("hidden-main")]
        [DataRow("run-control")]
        public void NativePreparationFailureReleasesAlreadyAcquiredInterfacesWithoutExecuting(string fault)
        {
            RunOnSta(() =>
            {
                var released = new List<object>(); var f = new NativeFixture(released.Add); int executed = 0;
                f.Editor.CommandBars.Control.OnExecute = () => executed++;
                string source = f.Module.Lines.Source;
                if (fault == "missing") f.Project.VBComponents.Clear();
                if (fault == "duplicate") f.Project.VBComponents.Add(f.Module.Parent);
                if (fault == "component-type") f.Module.Parent.Type = 2;
                if (fault == "source") source = "Changed source";
                if (fault == "wrapper") f.Module.Lines.Source = source = "Private Sub Hidden()";
                if (fault == "hidden-main") f.Editor.MainWindow.AllowVisibility = false;
                if (fault == "run-control") f.Editor.CommandBars.Control.OnAction = "PersonalMacro";
                Assert.ThrowsException<InvalidOperationException>(() => f.Probe.Prepare(f.Editor, f.Project, source));
                Assert.AreEqual(0, executed);
                Assert.IsTrue(released.Contains(f.Editor.MainWindow));
                Assert.IsTrue(released.Contains(f.Project.VBComponents));
                if (fault != "missing") Assert.IsTrue(released.Contains(f.Module.Parent));
                if (fault == "run-control") Assert.AreEqual(1, released.Count(p => ReferenceEquals(p, f.Editor.CommandBars.Control)));
            });
        }
    }
}
