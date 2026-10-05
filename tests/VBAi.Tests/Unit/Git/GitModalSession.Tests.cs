using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class GitModalSessionTests
    {
        private static GitModalSession.Request Request(string action = "checkpoint_restore") =>
            new GitModalSession.Request(action, "revision", "message", "choice", "Module1", new[] { "Module1" }, true, "expected-state");

        [DataTestMethod]
        [DataRow("checkpoint_restore", true)][DataRow("branch_switch", true)][DataRow("module_restore", true)]
        [DataRow("pull", true)][DataRow("merge_complete", true)][DataRow("rollback", true)]
        [DataRow("checkpoint_create", false)][DataRow("branch_create", false)][DataRow("branch_track", false)]
        [DataRow("commit", false)][DataRow("commit_selected", false)][DataRow("fetch", false)][DataRow("push", false)]
        [DataRow("merge_begin", false)][DataRow("merge_resolve", false)][DataRow("merge_abort", false)]
        [DataRow("remote_branches", false)][DataRow("pr_prepare", false)]
        public void OnlyTheSixImportActionsRequireModalReturn(string action, bool expected)
        { Assert.AreEqual(expected, GitModalSession.RequiresHandoff(action)); }

        [DataTestMethod, DataRow("checkpoint_restore"), DataRow("branch_switch"), DataRow("module_restore")]
        [DataRow("pull"), DataRow("merge_complete"), DataRow("rollback")]
        public async Task NoImportCanStartInsideTheActualShowCallEvenWhenLeaveCallbackReturnsInline(string action)
        {
            var events = new List<string>(); var request = Request(action); GitModalSession session = null;
            Task work = null; int shows = 0;
            session = new GitModalSession(() => {
                shows++;
                if (shows == 1)
                {
                    var admission = session.Queue(request, () => events.Add("leave-returned"));
                    Assert.IsFalse(admission.IsCompleted); Assert.AreEqual("AwaitingModalReturn", request.Phase);
                    work = Execute(admission, () => events.Add("execute"), request);
                    Assert.IsFalse(work.IsCompleted); events.Add("show-returning");
                }
                else { Assert.AreEqual("Succeeded", request.Phase); events.Add("presented"); }
            }, () => { Assert.IsTrue(events.Count == 2 || events.Count == 3); Assert.AreEqual("show-returning", events[1]); events.Add("validated"); });
            await session.RunAsync(); await work;
            CollectionAssert.AreEqual(new[] { "leave-returned", "show-returning", "validated", "validated", "execute", "presented" }, events);
            Assert.AreEqual(2, shows);
        }
        private static async Task Execute(Task admission, Action execute, GitModalSession.Request request)
        {
            Exception failure = null;
            try { await admission; execute(); } catch (Exception error) { failure = error; }
            finally { request.Complete(failure); }
        }

        [DataTestMethod, DataRow("show"), DataRow("leave"), DataRow("owner"), DataRow("execute")]
        public async Task AdmissionAndExecutionFailuresRetainTheOriginalErrorAndNeverReplay(string stage)
        {
            var original = new InvalidOperationException(stage); var request = Request(); GitModalSession session = null;
            Task work = null; int shows = 0, executions = 0;
            session = new GitModalSession(() => {
                if (++shows != 1) return;
                var admission = session.Queue(request, () => { if (stage == "leave") throw original; });
                work = Execute(admission, () => { executions++; if (stage == "execute") throw original; }, request);
                if (stage == "show") throw original;
            }, () => { if (stage == "owner") throw original; });
            if (stage == "show") Assert.AreSame(original, await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => session.RunAsync()));
            else await session.RunAsync();
            await work; Assert.AreSame(original, request.Error); Assert.AreEqual("Failed", request.Phase);
            Assert.AreEqual(stage == "execute" ? 1 : 0, executions);
            Assert.AreEqual(stage == "owner" || stage == "execute" ? 2 : 1, shows);
        }

        [TestMethod]
        public void RequestCopiesEveryArgumentAndCannotBeQueuedReleasedOrCompletedTwice()
        {
            var modules = new[] { "Module1" };
            var request = new GitModalSession.Request("module_restore", "commit", "text", "ours", "Module1", modules, true, "state");
            modules[0] = "foreign"; var copy = request.Modules; copy[0] = "foreign";
            Assert.AreEqual("Module1", request.Modules[0]); Assert.AreEqual("commit", request.Name);
            Assert.AreEqual("text", request.Text); Assert.AreEqual("ours", request.Choice); Assert.AreEqual("Module1", request.Path);
            Assert.AreEqual("state", request.Revision); Assert.IsTrue(request.References);
            request.Queue(); Assert.ThrowsException<InvalidOperationException>(request.Queue);
            request.Release(() => { }); Assert.ThrowsException<InvalidOperationException>(() => request.Release(() => Assert.Fail()));
            request.Complete(null); Assert.ThrowsException<InvalidOperationException>(() => request.Complete(null));
        }

        [TestMethod]
        public async Task PendingCompletionKeepsTheSessionAliveAndForbidsAnotherRunOrRequest()
        {
            var request = Request(); GitModalSession session = null; int shows = 0;
            session = new GitModalSession(() => {
                if (++shows != 1) return;
                session.Queue(request, () => { });
                Assert.ThrowsException<InvalidOperationException>(() => session.Queue(Request(), () => Assert.Fail()));
            }, () => { });
            var running = session.RunAsync();
            Assert.IsFalse(running.IsCompleted); Assert.AreEqual("Executing", request.Phase); Assert.AreEqual(1, shows);
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => session.RunAsync());
            Assert.ThrowsException<InvalidOperationException>(() => session.Queue(Request(), () => Assert.Fail()));
            request.Complete(null); await running; Assert.AreEqual(2, shows);
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => session.RunAsync());
        }

        [TestMethod]
        public async Task FailedLeaveStillRetainsTheSessionUntilTheRefusedOperationFinishesCleanup()
        {
            var request = Request(); var original = new InvalidOperationException("leave failed"); GitModalSession session = null;
            int shows = 0;
            session = new GitModalSession(() => { shows++; session.Queue(request, () => { throw original; }); }, () => Assert.Fail("No grant"));
            var running = session.RunAsync();
            Assert.IsFalse(running.IsCompleted); Assert.AreEqual("Refused", request.Phase);
            Assert.AreSame(original, await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => request.Admission));
            request.Complete(original); await running; Assert.AreEqual(1, shows);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task IndependentLeaveAndShowFailuresRemainAvailableAfterPendingCleanup(bool publicationFails)
        {
            var request = Request(); var leave = new InvalidOperationException("leave"); var show = new System.IO.IOException("show");
            GitModalSession session = null;
            session = new GitModalSession(() => { session.Queue(request, () => { throw leave; }); throw show; }, () => Assert.Fail());
            var run = session.RunAsync(); Assert.IsFalse(run.IsCompleted);
            Assert.AreSame(leave, await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => request.Admission));
            var publication = new InvalidOperationException("publication");
            request.Complete(leave, () => { if (publicationFails) throw publication; });
            var error = await Assert.ThrowsExceptionAsync<AggregateException>(() => run);
            CollectionAssert.AreEqual(publicationFails ? new Exception[] { show, leave, publication } : new Exception[] { show, leave }, error.Flatten().InnerExceptions);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task TerminalPublicationFailureCannotProduceSuccessOrMaskAnOperationFailure(bool operationFailed)
        {
            var request = Request(); var original = new InvalidOperationException("operation"); var publication = new System.IO.IOException("publication");
            request.Queue(); request.Release(() => { });
            request.Complete(operationFailed ? original : null, () => { throw publication; });
            Exception observed = null;
            try { await request.Completion; } catch (Exception error) { observed = error; }
            Assert.AreEqual("Failed", request.Phase);
            if (operationFailed) CollectionAssert.AreEqual(new Exception[] { original, publication }, ((AggregateException)observed).InnerExceptions);
            else Assert.AreSame(publication, observed);
        }

        [TestMethod]
        public async Task ReadOnlyRevalidationMustFinishBeforeTheSecondOwnerCheckAndAdmission()
        {
            var gate = new TaskCompletionSource<bool>(); int owners = 0, executed = 0, shows = 0;
            var request = new GitModalSession.Request("pull", null, null, null, null, null, false, "revision", () => gate.Task);
            GitModalSession session = null; Task work = null;
            session = new GitModalSession(() => { if (++shows == 1) work = Execute(session.Queue(request, () => { }), () => executed++, request); }, () => owners++);
            var running = session.RunAsync(); Assert.AreEqual(1, owners); Assert.AreEqual(0, executed); Assert.IsFalse(request.Admission.IsCompleted);
            gate.SetResult(true); await running; await work; Assert.AreEqual(2, owners); Assert.AreEqual(1, executed);
        }

        [DataTestMethod, DataRow("revalidate"), DataRow("second-owner")]
        public async Task PostReturnScopeOrOwnerDriftRefusesBeforeGrantAndRetainsTheOriginalFailure(string stage)
        {
            int owners = 0, shows = 0, executed = 0; var error = new InvalidOperationException(stage);
            var request = new GitModalSession.Request("pull", null, null, null, null, null, false, "revision",
                () => stage == "revalidate" ? Task.FromException(error) : Task.CompletedTask);
            GitModalSession session = null; Task work = null;
            session = new GitModalSession(() => { if (++shows == 1) work = Execute(session.Queue(request, () => { }), () => executed++, request); },
                () => { if (++owners == 2) throw error; });
            await session.RunAsync(); await work; Assert.AreEqual(0, executed); Assert.AreSame(error, request.Error);
            Assert.AreEqual(stage == "revalidate" ? 1 : 2, owners);
        }

        [DataTestMethod, DataRow("pid"), DataRow("tid"), DataRow("current"), DataRow("thread"), DataRow("handle")]
        [DataRow("missing"), DataRow("disabled"), DataRow("zero-pid"), DataRow("zero-tid"), DataRow("valid")]
        public void OriginalNativeOwnerOperandsMustAllRemainValid(string changed)
        {
            Action verify = () => GitModalSession.RequireOwner(changed == "zero-pid" ? 0u : 10u,
                changed == "zero-tid" ? 0u : 20u, changed == "thread" ? 21u : 20u,
                changed == "handle" ? IntPtr.Zero : new IntPtr(30), changed == "pid" ? 11u : 10u,
                changed == "tid" ? 21u : 20u, changed == "current" ? 21u : 20u,
                changed != "missing", changed != "disabled");
            if (changed == "valid") verify(); else Assert.ThrowsException<InvalidOperationException>(verify);
        }

        [DataTestMethod, DataRow(null, "state"), DataRow("fetch", "state"), DataRow("pull", null)]
        public void InvalidRequestsCannotEnterTheHandoff(string action, string revision)
        { Assert.ThrowsException<ArgumentException>(() => new GitModalSession.Request(action, null, null, null, null, null, false, revision)); }

        [DataTestMethod, DataRow("exact"), DataRow("foreign"), DataRow("null")]
        public async Task MutationOwnerCheckRequiresTheExactStillExecutingRequest(string kind)
        {
            var request = Request(); var candidate = kind == "exact" ? request : kind == "foreign" ? Request() : null;
            int checks = 0, shows = 0; GitModalSession session = null;
            session = new GitModalSession(() => {
                Assert.ThrowsException<InvalidOperationException>(() => session.RequireImportOwner(candidate));
                if (++shows == 1) session.Queue(request, () => { });
            }, () => checks++);
            Assert.ThrowsException<InvalidOperationException>(() => session.RequireImportOwner(candidate));
            var run = session.RunAsync();
            if (kind == "exact") session.RequireImportOwner(candidate);
            else Assert.ThrowsException<InvalidOperationException>(() => session.RequireImportOwner(candidate));
            Assert.AreEqual(kind == "exact" ? 3 : 2, checks);
            request.Complete(null); await run;
            Assert.ThrowsException<InvalidOperationException>(() => session.RequireImportOwner(candidate));
        }
    }
}
