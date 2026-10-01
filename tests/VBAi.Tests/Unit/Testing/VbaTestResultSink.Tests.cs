using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
    }
}
