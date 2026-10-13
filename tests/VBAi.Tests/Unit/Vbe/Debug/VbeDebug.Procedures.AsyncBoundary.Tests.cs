using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugTests
    {
        [STATestMethod]
        [DataRow("object"), DataRow("null"), DataRow("text")]
        public void NativeProcedureWorkerPreservesOneCallResultAndMtaIsolation(string scenario)
        {
            int caller = Thread.CurrentThread.ManagedThreadId, worker = 0, calls = 0;
            ApartmentState apartment = ApartmentState.Unknown;
            object expected = scenario == "null" ? null : scenario == "text" ? (object)"native result" : new object();
            var task = VbeDebug.RunProcedureWorker("Call Project.Module.Procedure(7)", command =>
            {
                calls++; worker = Thread.CurrentThread.ManagedThreadId;
                apartment = Thread.CurrentThread.GetApartmentState();
                Assert.AreEqual("Call Project.Module.Procedure(7)", command);
                return expected;
            });
            Assert.AreSame(expected, task.GetAwaiter().GetResult());
            Assert.AreEqual(1, calls);
            Assert.AreNotEqual(caller, worker);
            Assert.AreEqual(ApartmentState.MTA, apartment);
        }

        [STATestMethod]
        [DataRow("COM"), DataRow("surface"), DataRow("element"), DataRow("argument")]
        [DataRow("cancellation"), DataRow("task cancellation")]
        [DataRow("live token"), DataRow("requested token")]
        public void NativeProcedureWorkerPreservesOriginalFailureWithoutASecondInvocation(string scenario)
        {
            using (var cancellation = new CancellationTokenSource())
            {
                if (scenario == "requested token") cancellation.Cancel();
                var failures = new Dictionary<string, Exception>
                {
                    ["COM"] = new COMException("native refusal", unchecked((int)0x80004005)),
                    ["surface"] = new InvalidOperationException("unsupported Immediate surface"),
                    ["element"] = new ElementNotAvailableException("owned element closed"),
                    ["argument"] = new ArgumentException("invalid command"),
                    ["cancellation"] = new OperationCanceledException("cancelled native boundary"),
                    ["task cancellation"] = new TaskCanceledException("cancelled native task"),
                    ["live token"] = new OperationCanceledException("native token", cancellation.Token),
                    ["requested token"] = new OperationCanceledException("requested native token", cancellation.Token)
                };
                Exception expected = failures[scenario];
                int calls = 0;
                var task = VbeDebug.RunProcedureWorker("one call", command => { calls++; throw expected; });
                Exception actual = null;
                try { task.GetAwaiter().GetResult(); }
                catch (Exception error) { actual = error; }
                Assert.AreSame(expected, actual);
                if (expected is OperationCanceledException cancelled)
                    Assert.AreEqual(cancelled.CancellationToken, ((OperationCanceledException)actual).CancellationToken);
                Assert.AreEqual(1, calls);
                // Cancellation classification remains that of the original Task.Run boundary.
                var original = Task.Run<object>(() => { throw expected; });
                try { original.GetAwaiter().GetResult(); }
                catch (Exception) { }
                Assert.AreEqual(original.Status, task.Status);
            }
        }

        [TestMethod]
        public void NativeProcedureWorkerRejectsAMissingBoundaryBeforeScheduling()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbeDebug.RunProcedureWorker("call", null));
        }

        private sealed class ProcedureSubmissionContext : SynchronizationContext
        {
            private readonly Queue<Action> pending = new Queue<Action>();
            internal bool WrongSendThread;
            public override void Post(SendOrPostCallback callback, object state) => pending.Enqueue(() => callback(state));
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (!WrongSendThread) { callback(state); return; }
                ExceptionDispatchInfo failure = null;
                var thread = new Thread(() =>
                {
                    try { callback(state); }
                    catch (Exception error) { failure = ExceptionDispatchInfo.Capture(error); }
                });
                thread.Start(); thread.Join(); failure?.Throw();
            }
            internal void Drain() { while (pending.Count != 0) pending.Dequeue()(); }
        }

        [TestMethod]
        public void NativeProcedureSubmissionValidatesAndEntersOnceOnTheOwner()
        {
            var previous = SynchronizationContext.Current;
            var context = new ProcedureSubmissionContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = Create(2); var request = ProcedureRequest(fixture);
                fixture.Service.ShowProcedureImmediate = () => { };
                int calls = 0, entries = 0, owner = Thread.CurrentThread.ManagedThreadId;
                fixture.Service.ExecuteGuardedProcedureCall = (command, submit) =>
                {
                    calls++; submit(() => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); entries++; });
                    return Task.FromResult<object>("native delivery");
                };
                dynamic queued = fixture.Service.RunProcedure(request);
                context.Drain();
                dynamic result = fixture.Service.ProcedureRunStatus(new Request { Project = request.Project, Query = queued.Query });
                Assert.AreEqual("Delivered", (string)result.State);
                Assert.AreEqual("native delivery", (string)result.Output);
                Assert.AreEqual(1, calls); Assert.AreEqual(1, entries);
                Assert.IsFalse((bool)result.RuntimeSuccessVerified);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        [DataRow("source"), DataRow("mode"), DataRow("project"), DataRow("wrong owner")]
        public void NativeProcedureSubmissionRefusesChangesAfterPreparationBeforeEnter(string scenario)
        {
            var previous = SynchronizationContext.Current;
            var context = new ProcedureSubmissionContext { WrongSendThread = scenario == "wrong owner" };
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = Create(2); var request = ProcedureRequest(fixture);
                fixture.Service.ShowProcedureImmediate = () => { };
                int calls = 0, entries = 0;
                fixture.Service.ExecuteGuardedProcedureCall = (command, submit) =>
                {
                    calls++;
                    if (scenario == "source") fixture.Module.Code += "\r\n' source changed during preparation";
                    if (scenario == "mode") fixture.Project.Mode = 1;
                    if (scenario == "project") fixture.Project.Name = "DifferentProject";
                    submit(() => entries++);
                    return Task.FromResult<object>(null);
                };
                dynamic queued = fixture.Service.RunProcedure(request);
                context.Drain();
                dynamic result = fixture.Service.ProcedureRunStatus(new Request { Project = request.Project, Query = queued.Query });
                Assert.AreEqual("Failed", (string)result.State);
                Assert.IsFalse((bool)result.Pending);
                Assert.IsFalse(string.IsNullOrEmpty((string)result.Error));
                Assert.AreEqual(1, calls); Assert.AreEqual(0, entries);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }
}
