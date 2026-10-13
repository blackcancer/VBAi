using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbaNativeTestExecutionHostTests
    {
        [TestMethod]
        public void NativeTimingPairsPreserveOneDispatchAndCorrelatedCompletion()
        {
            using (var fixture = new Fixture())
            {
                var events = new List<NativeExecutionTimings.StageEvent>();
                fixture.Host.Timings = new NativeExecutionTimings(events.Add);
                Assert.AreEqual(VbaTestOutcome.Passed, fixture.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(1, fixture.Native.Invocations);
                foreach (var stage in new[] { NativeTestTimingStage.PrepareNativeDispatch, NativeTestTimingStage.ArmCallback,
                    NativeTestTimingStage.ExposeCallback, NativeTestTimingStage.ExecuteNativeDispatch, NativeTestTimingStage.ReadNativeMode,
                    NativeTestTimingStage.CompleteCallback, NativeTestTimingStage.RevokeCallback })
                {
                    var pair = events.Where(item => item.Stage == stage.ToString()).ToArray();
                    Assert.AreEqual(2, pair.Length, stage.ToString());
                    Assert.AreEqual("Started", pair[0].State); Assert.AreEqual("Returned", pair[1].State);
                }
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void BrokenTimingSinkCannotPreventNativeCallbackCompletionOrRevoke()
        {
            using (var fixture = new Fixture())
            {
                var observerFailure = new IOException("owned diagnostic failure");
                fixture.Host.Timings = new NativeExecutionTimings(_ => { throw observerFailure; });
                Assert.AreEqual(VbaTestOutcome.Passed, fixture.Start().GetAwaiter().GetResult().Outcome);
                Assert.AreEqual(1, fixture.Native.Invocations);
                Assert.AreSame(observerFailure, fixture.Host.Timings.DiagnosticFailure);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [TestMethod]
        public void DispatchedNativeFailureRemainsUncertainWithTimingsEnabled()
        {
            using (var fixture = new Fixture())
            {
                var events = new List<NativeExecutionTimings.StageEvent>();
                fixture.Host.Timings = new NativeExecutionTimings(events.Add);
                fixture.Native.OnExecute = () => { throw new InvalidOperationException("native failed after dispatch"); };
                var failure = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Start().GetAwaiter().GetResult());
                Assert.IsTrue(failure.Uncertain); Assert.AreEqual(1, fixture.Native.Invocations);
                Assert.AreEqual("Failed", events.Single(item => item.Stage == NativeTestTimingStage.ExecuteNativeDispatch.ToString() && item.State != "Started").State);
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }
    }
}
