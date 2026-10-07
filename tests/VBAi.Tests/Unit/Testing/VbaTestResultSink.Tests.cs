using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestResultSinkTests
    {
        private static readonly VbaTestDescriptor Test = new VbaTestDescriptor { Id = "test", Module = "Tests", Procedure = "Check", Kind = "Sub" };
        private static void Arm(VbaTestResultSink sink, object project)
            => sink.Arm(project, @"C:\fixture\macro.swp", "2", "revision", "run", Test, "Test", "signature");

        [TestMethod]
        public void VerdictRequiresClaimAndNativeCompletionAndIsDecodedByTheExistingRuntime()
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative();
                var job = (object[])sink.Request(project, @"C:\fixture\macro.swp", "2");
                Assert.AreEqual(7, job.Length); Assert.AreEqual(64, ((string)job[0]).Length);
                CollectionAssert.AreEqual(new object[] { "Tests", "Check", "Test", "revision", "run", "test" }, new object[] { job[1], job[2], job[3], job[4], job[5], job[6] });
                Assert.IsTrue(sink.Publish((string)job[0], "revision", "Failed", "An assertion failed.", 0));
                var result = VbaTestRuntimeSource.Decode(Test, sink.CompleteNative(true));
                Assert.AreEqual(VbaTestOutcome.Failed, result.Outcome); Assert.AreEqual("An assertion failed.", result.Message);
            }
        }

        [DataTestMethod]
        [DataRow("project")]
        [DataRow("path")]
        [DataRow("version")]
        [DataRow("duplicate")]
        public void ARejectedClaimPoisonsTheNativeAttempt(string failure)
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative();
                if (failure == "duplicate") sink.Request(project, @"C:\fixture\macro.swp", "2");
                Assert.ThrowsException<InvalidOperationException>(() => sink.Request(failure == "project" ? new object() : project,
                    failure == "path" ? @"C:\foreign\macro.swp" : @"C:\fixture\macro.swp", failure == "version" ? "1" : "2"));
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
                Assert.ThrowsException<InvalidOperationException>(() => Arm(sink, project));
            }
        }

        [DataTestMethod]
        [DataRow("nonce")]
        [DataRow("revision")]
        [DataRow("status")]
        [DataRow("message")]
        [DataRow("error")]
        [DataRow("duplicate")]
        public void InvalidOrRepeatedVerdictsCannotBecomePassed(string failure)
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative();
                var job = (object[])sink.Request(project, @"C:\fixture\macro.swp", "2");
                if (failure == "duplicate") sink.Publish((string)job[0], "revision", "Passed", "", 0);
                Assert.ThrowsException<InvalidOperationException>(() => sink.Publish(failure == "nonce" ? "foreign" : (string)job[0],
                    failure == "revision" ? "foreign" : "revision", failure == "status" ? "passed" : "Passed",
                    failure == "message" ? new string('x', VbaTestRuntimeSource.MaximumMessageLength + 1) : "", failure == "error" ? 5 : 0));
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NativeFailureOrMissingVerdictNeverAuthorizesTheNextCall(bool callback)
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative();
                if (callback)
                {
                    var job = (object[])sink.Request(project, @"C:\fixture\macro.swp", "2");
                    sink.Publish((string)job[0], "revision", "Passed", "", 0);
                }
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(!callback, "Native call failed.")).Uncertain);
                Assert.ThrowsException<InvalidOperationException>(() => Arm(sink, project));
            }
        }

        [TestMethod]
        public void ForeignThreadCallbackPoisonsOnlyAnAlreadyDispatchedAttempt()
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative(); Exception caught = null;
                var thread = new Thread(() => { try { sink.Request(project, @"C:\fixture\macro.swp", "2"); } catch (Exception error) { caught = error; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(caught, typeof(InvalidOperationException));
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
            }
        }

        [TestMethod]
        public void PreDispatchRefusalsAndCancellationLeaveNoExecutableJob()
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project);
                Assert.ThrowsException<InvalidOperationException>(() => sink.Request(project, @"C:\fixture\macro.swp", "2"));
                Assert.ThrowsException<InvalidOperationException>(() => Arm(sink, project));
                sink.CancelUndispatched(); Arm(sink, project); sink.CancelUndispatched();
            }
        }

        [TestMethod]
        public void DisconnectionDuringAStartedCallRemainsUncertain()
        {
            var sink = new VbaTestResultSink(ReferenceEquals); Arm(sink, new object()); sink.BeginNative(); sink.Dispose();
            Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
        }
        [TestMethod]
        public void ArmRejectsEveryInvalidIdentityBeforeCreatingAJobAndAllowsEveryFixturePhase()
        {
            using (var sink = new VbaTestResultSink())
            {
                for (int index = 0; index < 12; index++)
                {
                    var descriptor = new VbaTestDescriptor { Id = index == 6 ? null : "test", Module = index == 7 ? "" : "Tests", Procedure = index == 8 ? new string('x', 256) : "Check" };
                    Assert.ThrowsException<ArgumentException>(() => sink.Arm(index == 0 ? null : new object(), index == 2 ? null : "path", index == 3 ? new string('x', 33) : "3",
                        index == 4 ? " " : "revision", index == 5 ? new string('x', 129) : "run", index == 1 ? null : descriptor,
                        index == 9 ? "unknown" : "Test", index == 10 ? " " : index == 11 ? new string('x', 129) : null));
                    Assert.IsFalse(sink.HasVerdict); Assert.IsFalse(sink.HasFault);
                }
                foreach (var phase in new[] { "Test", "ModuleInitialize", "ModuleCleanup", "TestInitialize", "TestCleanup" })
                { sink.Arm(new object(), "path", "3", "revision", "run", Test, phase); sink.CancelUndispatched(); }
                Assert.ThrowsException<InvalidOperationException>(() => sink.BeginNative());
                Assert.ThrowsException<InvalidOperationException>(() => sink.CompleteNative(true));
            }
        }

        [TestMethod]
        public void LifecycleGuardsPreserveDispatchAndFaultState()
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project);
                Assert.IsFalse(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
                Arm(sink, project); sink.BeginNative();
                Assert.ThrowsException<InvalidOperationException>(() => sink.BeginNative());
                Assert.ThrowsException<InvalidOperationException>(() => sink.CancelUndispatched());
                Assert.IsFalse(sink.HasFault); Assert.IsFalse(sink.HasVerdict);
                sink.Request(project, @"C:\fixture\macro.swp", "2");
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
            }
            var disposed = new VbaTestResultSink(); disposed.Dispose();
            Assert.ThrowsException<InvalidOperationException>(() => Arm(disposed, project));
            Assert.ThrowsException<InvalidOperationException>(() => disposed.BeginNative());
            Assert.ThrowsException<InvalidOperationException>(() => disposed.BindRuntime());
        }

        [DataTestMethod]
        [DataRow("nullPath")]
        [DataRow("longPath")]
        [DataRow("nullVersion")]
        [DataRow("longVersion")]
        [DataRow("beforeClaim")]
        [DataRow("nullMessage")]
        [DataRow("inconclusiveError")]
        [DataRow("localSignature")]
        [DataRow("localDuplicate")]
        [DataRow("localBinding")]
        public void RejectedCallbackDetailsRemainStickyAndCannotBeClaimedAgain(string failure)
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative();
                if (failure.EndsWith("Path") || failure.EndsWith("Version"))
                    Assert.ThrowsException<InvalidOperationException>(() => sink.Request(project, failure == "nullPath" ? null : failure == "longPath" ? new string('x', 32769) : @"C:\fixture\macro.swp", failure == "nullVersion" ? null : failure == "longVersion" ? new string('x', 33) : "2"));
                else if (failure.StartsWith("local"))
                {
                    string binding = sink.BindRuntime();
                    if (failure == "localDuplicate") sink.RequestLocal(binding, "2", "signature");
                    Assert.ThrowsException<InvalidOperationException>(() => sink.RequestLocal(failure == "localBinding" ? "other" : binding, "2", failure == "localSignature" ? null : "signature"));
                }
                else
                {
                    var job = failure == "beforeClaim" ? null : (object[])sink.Request(project, @"C:\fixture\macro.swp", "2");
                    Assert.ThrowsException<InvalidOperationException>(() => sink.Publish(job == null ? "unknown" : (string)job[0], "revision", failure == "inconclusiveError" ? "Inconclusive" : "Passed", failure == "nullMessage" ? null : "", 5));
                }
                Assert.IsTrue(sink.HasFault);
                Assert.ThrowsException<InvalidOperationException>(() => sink.BindRuntime());
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
            }
        }

        [DataTestMethod]
        [DataRow("Failed", 5)]
        [DataRow("Error", -5)]
        [DataRow("Inconclusive", 0)]
        public void ValidNonPassedVerdictsRetainTheirExactPayload(string status, int number)
        {
            var project = new object();
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink, project); sink.BeginNative(); var job = (object[])sink.Request(project, @"C:\fixture\macro.swp", "2");
                Assert.IsTrue(sink.Publish((string)job[0], "revision", status, "reason", number)); Assert.IsTrue(sink.HasVerdict);
                CollectionAssert.AreEqual(new object[] { status, "reason", number.ToString(System.Globalization.CultureInfo.InvariantCulture) }, (object[])sink.CompleteNative(true));
                Exception caught = null;
                var thread = new Thread(() => { try { sink.CancelUndispatched(); } catch (Exception error) { caught = error; } });
                thread.Start(); thread.Join(); Assert.IsInstanceOfType(caught, typeof(InvalidOperationException));
            }
        }
        [TestMethod]
        public void RepeatedUndispatchedCancellationIsSafeAndLeavesNoClaimableJob()
        {
            using (var sink = new VbaTestResultSink())
            { sink.CancelUndispatched(); sink.CancelUndispatched(); Assert.ThrowsException<InvalidOperationException>(() => sink.BindRuntime()); }
        }
    }
}
