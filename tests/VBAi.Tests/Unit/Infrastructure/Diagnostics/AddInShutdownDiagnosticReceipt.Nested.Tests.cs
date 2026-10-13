using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class AddInShutdownDiagnosticReceiptNestedTests
    {
        private static readonly string[] Runtime = {
            "contextMonitor.Stop", "referenceEvents.Dispose", "projectEvents.Dispose", "componentEvents.Dispose", "SaveCurrentSession",
            "streamRender.Stop", "pendingFollow.Abort", "saveTimer.Stop", "historySearch.Stop", "projectRetry.Stop", "scopeIdentities.Dispose",
            "persistenceWorker.Dispose", "sessionStore.Dispose", "httpClient.Dispose", "codex.Dispose", "changesMenu.Dispose",
            "menuImages.Dispose", "entryViews.Dispose", "composer.Dispose"
        };
        private static Dictionary<string, object> Row(object value) => (Dictionary<string, object>)value;
        private static void Number(List<object> events)
        {
            for (int i = 0; i < events.Count; i++) { Row(events[i])["Sequence"] = i + 1; Row(events[i])["ElapsedTicks"] = (long)i; }
        }
        private static AddInShutdownDiagnosticReceipt.Invocation[] Validate(List<object> events)
        {
            Number(events);
            return AddInShutdownDiagnosticReceipt.Validate(events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity());
        }
        private static List<object> Chain(bool runtime = true, bool boundaries = true, bool reentrant = false,
            int caughtRuntime = -1, bool observed = false, string propagated = null, Exception closeError = null)
        {
            var observation = new ShutdownDiagnosticTestData.Observation();
            var root = observation.Begin();
            root.Enter("VbeNativeTheme.Disconnect"); root.Returned("VbeNativeTheme.Disconnect");
            foreach (string field in ShutdownDiagnosticTestData.Cleared)
            {
                if (field == "nativeTestWindow" && closeError != null)
                {
                    root.Enter("nativeTestWindow.Close"); root.CaughtFault("nativeTestWindow.Close", closeError);
                }
                if (field == "crashReporter") { root.Enter("StopUpdateCheck"); root.Returned("StopUpdateCheck"); }
                if (field == "chat")
                {
                    root.Enter("chat.Dispose"); var chat = root.Child("Dispose");
                    if (runtime)
                        for (int index = 0; index < Runtime.Length; index++)
                        {
                            string stage = "chat.Runtime." + Runtime[index]; chat.Enter(stage);
                            if (reentrant && index == 13)
                            {
                                var nested = chat.Child("Dispose");
                                nested.Enter("chat.Components.Dispose"); nested.Returned("chat.Components.Dispose");
                                nested.Enter("chat.Base.Dispose"); nested.Returned("chat.Base.Dispose"); nested.Complete();
                            }
                            if (index == caughtRuntime) chat.CaughtFault(stage, new IOException("original runtime cleanup error"));
                            else chat.Returned(stage);
                        }
                    chat.Enter("chat.Components.Dispose");
                    if (propagated == "components") { chat.Fault(new IOException("components")); root.Fault(new IOException("components")); return observation.Events; }
                    chat.Returned("chat.Components.Dispose");
                    chat.Enter("chat.Base.Dispose");
                    if (boundaries)
                    {
                        var destroy = chat.Child("Dispose"); destroy.Enter("chat.NativeHandle.Destroy");
                        var notification = destroy.Child("Dispose"); notification.Enter("chat.NativeHandle.DestroyedNotification");
                        notification.Returned("chat.NativeHandle.DestroyedNotification"); notification.Complete();
                        if (observed) destroy.ExceptionObserved(new InvalidComObjectException("caught framework exception"));
                        if (propagated == "boundary")
                        {
                            destroy.Fault(new IOException("native boundary")); chat.Fault(new IOException("native boundary"));
                            root.Fault(new IOException("native boundary")); return observation.Events;
                        }
                        destroy.Returned("chat.NativeHandle.Destroy"); destroy.Complete();
                        var parent = chat.Child("Dispose"); parent.Enter("chat.Parent.ChangedNotification");
                        parent.Returned("chat.Parent.ChangedNotification"); parent.Complete();
                    }
                    if (propagated == "base") { chat.Fault(new IOException("base")); root.Fault(new IOException("base")); return observation.Events; }
                    chat.Returned("chat.Base.Dispose"); chat.Complete(); root.Returned("chat.Dispose");
                }
                root.ManagedReferenceCleared(field);
            }
            root.Complete(); return observation.Events;
        }

        [DataTestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void CompletedChatRetainsExactNestedRolesAndDoesNotRequireAnUncreatedHandle(bool runtime, bool boundaries)
        {
            var result = Validate(Chain(runtime, boundaries));
            Assert.AreEqual(boundaries ? 5 : 2, result.Length);
            Assert.IsTrue(result.All(invocation => invocation.Outcome == "Returned"));
            Assert.AreEqual(runtime ? 21 : 2, result.Single(invocation => invocation.Role == "Chat").Stages.Count);
            Assert.AreEqual(16, result.Single(invocation => invocation.Role == "AddIn").ClearedReferences.Count);
        }

        [TestMethod]
        public void ReentrantDisposalPreservesTheOuterRuntimeAndDoesNotInventRepeatedRuntimeCleanup()
        {
            var result = Validate(Chain(reentrant: true));
            Assert.AreEqual(21, result.Single(invocation => invocation.Role == "Chat").Stages.Count);
            Assert.AreEqual(2, result.Single(invocation => invocation.Role == "ChatReentrant").Stages.Count);
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6), DataRow(7), DataRow(8), DataRow(9)]
        [DataRow(10), DataRow(11), DataRow(12), DataRow(13), DataRow(14), DataRow(15), DataRow(16), DataRow(17), DataRow(18)]
        public void EveryCaughtRuntimeFailureRetainsLaterResourcesAndItsOriginalFault(int stage)
        {
            var result = Validate(Chain(caughtRuntime: stage));
            var chat = result.Single(invocation => invocation.Role == "Chat");
            Assert.AreEqual("Fault", chat.Stages["chat.Runtime." + Runtime[stage]]);
            Assert.AreEqual("Returned", chat.Stages["chat.Base.Dispose"]);
            Assert.AreEqual("Returned", chat.Outcome);
        }

        [DataTestMethod, DataRow("components"), DataRow("base"), DataRow("boundary")]
        public void PropagatedFailureRetainsItsExactUnfinishedNativeScopeWithoutClaimingSuccessfulCleanup(string stage)
        {
            var result = Validate(Chain(propagated: stage));
            Assert.AreEqual("Fault", result.Single(invocation => invocation.Role == "AddIn").Outcome);
            Assert.AreEqual("Fault", result.Single(invocation => invocation.Role == "Chat").Outcome);
            Assert.IsFalse(result.Any(invocation => invocation.Role == "AddIn" && invocation.ClearedReferences.Contains("chat")));
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OnlyExistingNativeTestCloseCatchPoliciesAdmitBothComExceptionTypes(bool detached)
        {
            Exception error = detached ? (Exception)new InvalidComObjectException("detached native test window") : new COMException("native test window close");
            Assert.AreEqual("Fault", Validate(Chain(closeError: error)).Single(invocation => invocation.Role == "AddIn").Stages["nativeTestWindow.Close"]);
        }

        [TestMethod]
        public void FirstChanceObservationDoesNotTerminateItsStageOrCreateAnInvocationTerminal()
        {
            var events = Chain(observed: true); var observed = events.Select(Row).Single(row => (string)row["Phase"] == "ExceptionObserved");
            var result = Validate(events);
            Assert.IsNull(observed["FaultCaught"]); Assert.AreEqual(false, observed["ErrorIsComException"]);
            Assert.IsTrue(result.All(invocation => invocation.Outcome == "Returned"));
            Assert.AreEqual("Returned", result.Single(invocation => invocation.Stages.ContainsKey("chat.NativeHandle.Destroy")).Stages["chat.NativeHandle.Destroy"]);
        }

        [DataTestMethod]
        [DataRow("runtime gap"), DataRow("runtime order"), DataRow("runtime duplicate"), DataRow("invented runtime")]
        [DataRow("component omitted"), DataRow("base omitted"), DataRow("runtime reentry"), DataRow("missing chat child")]
        [DataRow("unknown boundary"), DataRow("empty boundary"), DataRow("foreign boundary parent"), DataRow("parent under destroy")]
        [DataRow("destroyed under base"), DataRow("invented component catch"), DataRow("invented boundary catch")]
        [DataRow("duplicate boundary"), DataRow("boundary order")]
        [DataRow("chat clear"), DataRow("chat operation outside chat"), DataRow("nested host callback")]
        [DataRow("observer error absent"), DataRow("observer caught flag"), DataRow("observer com flag")]
        [DataRow("observer wrong stage"), DataRow("observer false terminal"), DataRow("observer foreign thread")]
        public void NestedChainRejectsEveryMissingForeignOrInventedBoundary(string kind)
        {
            var events = Chain(reentrant: true, observed: true);
            var rows = events.Select(Row).ToList();
            var chat = rows.First(row => row["Stage"] as string == "chat.Components.Dispose");
            string chatId = (string)rows.First(row => row["Stage"] as string == "chat.Runtime.contextMonitor.Stop")["InvocationId"];
            var observer = rows.Single(row => row["Phase"] as string == "ExceptionObserved");
            var boundary = rows.First(row => row["Stage"] as string == "chat.NativeHandle.Destroy");
            string destroyId = (string)boundary["InvocationId"];
            if (kind == "runtime gap" || kind == "component omitted" || kind == "base omitted")
            {
                string stage = kind == "runtime gap" ? "chat.Runtime.SaveCurrentSession" : kind == "component omitted" ? "chat.Components.Dispose" : "chat.Base.Dispose";
                events.RemoveAll(value => Row(value)["Stage"] as string == stage && Row(value)["InvocationId"] as string == chatId);
            }
            if (kind == "runtime order")
                foreach (var row in rows)
                {
                    if (row["Stage"] as string == "chat.Runtime.contextMonitor.Stop") row["Stage"] = "chat.Runtime.referenceEvents.Dispose";
                    else if (row["Stage"] as string == "chat.Runtime.referenceEvents.Dispose") row["Stage"] = "chat.Runtime.contextMonitor.Stop";
                }
            if (kind == "runtime duplicate")
            {
                int position = events.FindIndex(value => Row(value)["Stage"] as string == "chat.Runtime.projectEvents.Dispose");
                var pair = rows.Where(row => row["Stage"] as string == "chat.Runtime.referenceEvents.Dispose").Select(ShutdownDiagnosticTestData.Clone).ToArray();
                events.InsertRange(position, pair);
            }
            if (kind == "invented runtime") rows.First(row => row["Stage"] as string == "chat.Runtime.contextMonitor.Stop")["Stage"] = "chat.Runtime.Unknown.Release";
            if (kind == "runtime reentry") chat["Stage"] = "chat.Runtime.contextMonitor.Stop";
            if (kind == "missing chat child")
                events.RemoveAll(value => { var row = Row(value); return row["InvocationId"] as string != (string)rows.First()["InvocationId"] && row["ParentInvocationId"] != null; });
            if (kind == "unknown boundary") boundary["Stage"] = "chat.NativeHandle.Unknown";
            if (kind == "empty boundary") events.RemoveAll(value =>
                Row(value)["InvocationId"] as string == destroyId && Row(value)["Stage"] != null || Row(value)["ParentInvocationId"] as string == destroyId);
            if (kind == "duplicate boundary" || kind == "boundary order")
            {
                string parentId = (string)rows.First(row => row["Stage"] as string == "chat.Parent.ChangedNotification")["InvocationId"];
                var changed = events.Where(value => Row(value)["InvocationId"] as string == parentId).ToArray();
                if (kind == "duplicate boundary")
                {
                    string duplicateId = Guid.NewGuid().ToString("N");
                    var duplicate = changed.Select(ShutdownDiagnosticTestData.Clone).ToArray();
                    foreach (var item in duplicate) Row(item)["InvocationId"] = duplicateId;
                    events.InsertRange(events.IndexOf(changed.Last()) + 1, duplicate);
                }
                else
                {
                    events.RemoveAll(value => Row(value)["InvocationId"] as string == parentId);
                    events.InsertRange(events.FindIndex(value => Row(value)["InvocationId"] as string == destroyId), changed);
                }
            }
            if (kind == "foreign boundary parent") rows.First(row => row["InvocationId"] as string == destroyId)["ParentInvocationId"] = Guid.NewGuid().ToString("N");
            if (kind == "parent under destroy")
                foreach (var row in rows.Where(row => row["Stage"] as string == "chat.NativeHandle.DestroyedNotification")) row["Stage"] = "chat.Parent.ChangedNotification";
            if (kind == "destroyed under base") boundary["Stage"] = "chat.NativeHandle.DestroyedNotification";
            if (kind == "invented component catch" || kind == "invented boundary catch")
            {
                string stage = kind == "invented component catch" ? "chat.Components.Dispose" : "chat.NativeHandle.Destroy";
                var row = rows.First(item => item["Stage"] as string == stage && item["Phase"] as string == "Returned");
                row["Phase"] = "Fault"; row["Error"] = "invented swallowed exception"; row["FaultCaught"] = true; row["ErrorIsComException"] = false;
            }
            if (kind == "chat clear") { chat["Phase"] = "ManagedReferenceCleared"; chat["Stage"] = "chat"; }
            if (kind == "chat operation outside chat") rows.First(row => row["Stage"] as string == "chat.Dispose")["Stage"] = "chat.Base.Dispose";
            if (kind == "nested host callback") rows.First(row => row["InvocationId"] as string == chatId)["EntryPoint"] = "OnBeginShutdown";
            if (kind == "observer error absent") observer["Error"] = null;
            if (kind == "observer caught flag") observer["FaultCaught"] = true;
            if (kind == "observer com flag") observer["ErrorIsComException"] = null;
            if (kind == "observer wrong stage") observer["Stage"] = "chat.Base.Dispose";
            if (kind == "observer false terminal") observer["Stage"] = null;
            if (kind == "observer foreign thread") Row(observer["Thread"])["NativeThreadId"] = 999;
            Assert.ThrowsException<InvalidOperationException>(() => Validate(events), kind);
        }

        [DataTestMethod, DataRow("callback"), DataRow("chat"), DataRow("boundary")]
        public void NativeCallbacksReenterTheirExactActiveParentWithoutAdoptingItsCleanup(string role)
        {
            var observation = new ShutdownDiagnosticTestData.Observation();
            var root = observation.Begin(role == "callback" ? "OnBeginShutdown" : "Dispose");
            if (role == "callback")
            {
                root.Enter("CleanupTemporaryToolbarCommands");
                var callback = root.Child("OnDisconnection"); ShutdownDiagnosticTestData.FinishDispose(callback.Child("Dispose")); callback.Complete();
                root.Returned("CleanupTemporaryToolbarCommands");
                ShutdownDiagnosticTestData.FinishDispose(root.Child("Dispose")); root.Complete();
            }
            else
            {
                root.Enter("VbeNativeTheme.Disconnect"); root.Returned("VbeNativeTheme.Disconnect");
                foreach (string field in ShutdownDiagnosticTestData.Cleared)
                {
                    if (field == "crashReporter") { root.Enter("StopUpdateCheck"); root.Returned("StopUpdateCheck"); }
                    if (field == "chat")
                    {
                        root.Enter("chat.Dispose"); var chat = root.Child("Dispose");
                        chat.Enter("chat.Components.Dispose"); chat.Returned("chat.Components.Dispose");
                        chat.Enter("chat.Base.Dispose");
                        var parent = role == "chat" ? chat : chat.Child("Dispose");
                        if (role == "boundary") parent.Enter("chat.NativeHandle.Destroy");
                        var callback = parent.Child("OnDisconnection"); ShutdownDiagnosticTestData.FinishDispose(callback.Child("Dispose")); callback.Complete();
                        if (role == "boundary") { parent.Returned("chat.NativeHandle.Destroy"); parent.Complete(); }
                        chat.Returned("chat.Base.Dispose"); chat.Complete(); root.Returned("chat.Dispose");
                    }
                    root.ManagedReferenceCleared(field);
                }
                root.Complete();
            }
            var result = Validate(observation.Events);
            var reentered = result.Single(invocation => invocation.EntryPoint == "OnDisconnection");
            Assert.IsNotNull(reentered.Parent);
            Assert.IsTrue(result.All(invocation => invocation.Outcome == "Returned"));
            Assert.AreEqual(16, result.Single(invocation => invocation.Parent == reentered.Id).ClearedReferences.Count);
        }

        [TestMethod]
        public void ReentrantCallbackCannotAppearBetweenOperations()
        {
            var observation = new ShutdownDiagnosticTestData.Observation(); var root = observation.Begin();
            var callback = root.Child("OnDisconnection"); ShutdownDiagnosticTestData.FinishDispose(callback.Child("Dispose")); callback.Complete();
            ShutdownDiagnosticTestData.FinishDispose(root);
            Assert.ThrowsException<InvalidOperationException>(() => Validate(observation.Events));
        }
    }
}
