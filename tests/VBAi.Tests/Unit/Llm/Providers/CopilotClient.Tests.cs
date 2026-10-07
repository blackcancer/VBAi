using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie le client Copilot avec sessions, transports et callbacks simulés.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class CopilotClientCoverageTests
    {
        /// <summary>Définition de l’outil de lecture de module utilisée par les fixtures.</summary>
        /// <returns>Tableau contenant l’outil disponible dans les tours simulés.</returns>
        private static object[] Tools() { return new object[] { new { function = new { name = "read_module", description = "fixture tool", parameters = new { type = "object", properties = new { } } } } }; }
        /// <summary>Crée l’historique système et utilisateur commun aux scénarios Copilot.</summary>
        /// <returns>Messages de la conversation de test.</returns>
        private static List<object> History() { return new List<object> { new Dictionary<string, object> { { "role", "system" }, { "content", "Read before editing." } }, new Dictionary<string, object> { { "role", "user" }, { "content", "request" } } }; }
        /// <summary>Crée le client après neutralisation du contexte de synchronisation WinForms courant.</summary>
        /// <returns>Client Copilot associé au transport configuré dans le scope.</returns>
        private static CopilotClient Client()
        {
            var context = SynchronizationContext.Current; try { SynchronizationContext.SetSynchronizationContext(null); return new CopilotClient(); } finally { SynchronizationContext.SetSynchronizationContext(context); }
        }
        /// <summary>Refuse les démarrages, requêtes après disposal et entêtes de protocole incomplets.</summary>
        /// <returns>Tâche terminée après les vérifications asynchrones.</returns>
        [TestMethod]
        public async Task NativeStartFailuresDisposedRequestsAndProtocolHeadersFailWithoutExternalCli()
        {
            using (var scope = new LlmBoundaryScope())
            {
                Assert.AreEqual("copilot.exe", CopilotClient.Executable);
                Environment.SetEnvironmentVariable("VBAi_COPILOT_CLI", Path.Combine(scope.Root, "missing.exe")); using (var client = Client()) { var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync()); StringAssert.Contains(error.Message, "VBAi_COPILOT_CLI"); Assert.IsNotNull(error.InnerException); }
                using (var client = Client()) { client.StartProcess = p => { Assert.AreEqual(ProviderSessionStorage.CopilotHome, p.StartInfo.EnvironmentVariables["COPILOT_HOME"]); return false; }; var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync()); Assert.IsInstanceOfType<InvalidOperationException>(error.InnerException); client.Dispose(); client.Dispose(); await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => client.ListModelsAsync()); await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => (Task<IDictionary<string, object>>)LlmBoundaryScope.Call(client, "RequestAsync", "after.dispose", new { })); var send = Assert.ThrowsException<TargetInvocationException>(() => LlmBoundaryScope.Call(client, "Send", new { })); Assert.IsInstanceOfType<ObjectDisposedException>(send.InnerException); }
                foreach (var mode in new[] { "header-eof", "header-long", "size-zero", "size-large", "size-negative", "size-invalid", "body-eof", "bad-json", "version-invalid", "version-low", "version-high", "rpc-error" })
                {
                    scope.UseCopilot(mode); using (var client = Client()) { Exception error = null; try { await client.ListModelsAsync(); } catch (Exception ex) { error = ex; } Assert.IsNotNull(error, mode); Assert.IsFalse(error.Message.Contains("secret must not leak"), mode); }
                }
            }
        }
        /// <summary>Vérifie sessions encadrées, modèles, outils, permissions et streaming dans leur portée autorisée.</summary>
        /// <returns>Tâche terminée après le tour simulé.</returns>
        [TestMethod]
        public async Task NativeFramedSessionsModelsToolsPermissionsAndStreamingStayScoped()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var mode in new[] { "normal", "normal-no-delta", "legacy", "unknown-tool", "empty-answer", "early-tool" })
                {
                    scope.UseCopilot(mode); using (var client = Client())
                    {
                        var models = await client.ListModelsAsync(); CollectionAssert.AreEqual(new[] { "model", "fallback" }, models.Select(x => x.Id).ToArray()); Assert.AreEqual("fallback", models[1].Label); Assert.AreEqual(2, (await client.ListModelsAsync()).Length);
                        int calls = 0; var chunks = new List<string>(); if (mode != "normal-no-delta") client.TextDelta = chunks.Add; var history = History(); var result = await client.CompleteAsync("model", history, Tools(), (name, arguments) => { calls++; Assert.AreEqual("read_module", name); Assert.AreEqual("{}", arguments); return Task.FromResult("fixture tool result"); });
                        Assert.AreEqual(mode == "empty-answer" ? "" : "answer été", result["content"]); Assert.AreEqual(mode == "unknown-tool" || mode == "empty-answer" ? 0 : 1, calls); Assert.AreEqual(mode == "empty-answer" ? 2 : 4, history.Count); if (mode != "empty-answer" && mode != "normal-no-delta") CollectionAssert.AreEqual(new[] { "delta été" }, chunks); else Assert.AreEqual(0, chunks.Count);
                        LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { method = "session.event", @params = new { sessionId = LlmBoundaryScope.Get<string>(client, "sessionId"), @event = new { type = "session.idle", data = new { } } } }));
                    }
                }
                scope.UseCopilot("models-null"); using (var client = Client()) Assert.AreEqual(0, (await client.ListModelsAsync()).Length); scope.UseCopilot(); CopilotClient.StartLogin(); var marker = Path.Combine(scope.Root, "login.marker"); var deadline = DateTime.UtcNow.AddSeconds(5); while (!File.Exists(marker + ".ready") && DateTime.UtcNow < deadline) await Task.Delay(10); Assert.IsTrue(File.Exists(marker + ".ready"), "The fixture must finish publishing its login marker before readback."); Assert.AreEqual("fixture login only", File.ReadAllText(marker)); StringAssert.Contains(await CopilotClient.ReadStatusAsync(), "2");
            }
        }
        /// <summary>Termine les requêtes en attente après erreurs RPC, outils défaillants ou délais dépassés.</summary>
        /// <returns>Tâche terminée après les scénarios d’erreur.</returns>
        [TestMethod]
        public async Task NativeRpcSessionAndToolFailuresAndDeadlinesCompletePendingRequests()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var mode in new[] { "session-mismatch", "session-error", "permission-error", "tool-reply-error" }) { scope.UseCopilot(mode); using (var client = Client()) { await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync("model", History(), Tools(), (n, a) => Task.FromResult("result"))); } }
                scope.UseCopilot(); using (var client = Client()) { await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync("model", History(), Tools(), null)); await Assert.ThrowsExceptionAsync<IOException>(() => client.CompleteAsync("model", History(), Tools(), (n, a) => Task.FromException<string>(new IOException("fixture tool failure")))); }
                scope.UseCopilot("request-wait"); using (var client = Client()) { client.Delay = t => Task.CompletedTask; await Assert.ThrowsExceptionAsync<TimeoutException>(() => client.ListModelsAsync()); }
                scope.UseCopilot("completion-wait"); using (var client = Client()) { client.Delay = t => t == TimeSpan.FromMinutes(5) ? Task.CompletedTask : Task.Delay(t); await Assert.ThrowsExceptionAsync<TimeoutException>(() => client.CompleteAsync("model", History(), Tools(), (n, a) => Task.FromResult("unused"))); }
                scope.UseCopilot("exit-models"); using (var client = Client()) { await client.ListModelsAsync(); Assert.IsTrue(LlmBoundaryScope.Get<Process>(client, "process").WaitForExit(5000)); client.Dispose(); }
                scope.UseCopilot(); using (var client = Client()) { await client.ListModelsAsync(); var process = LlmBoundaryScope.Get<Process>(client, "process"); process.Kill(); Assert.IsTrue(process.WaitForExit(5000)); process.Close(); client.Dispose(); }
            }
        }
        /// <summary>Revalide la disposal et l’état du tour avant l’exécution d’un outil par callback différé.</summary>
        /// <returns>Tâche terminée après traitement des callbacks.</returns>
        [TestMethod]
        public async Task QueuedUiCallbacksRecheckDisposalAndCompletedTurnsBeforeInvokingTools()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var mode in new[] { "queued-tool", "queued-tool-no-delta", "queued-wait" })
                {
                    scope.UseCopilot(mode == "queued-tool-no-delta" ? "queued-tool" : mode); var queue = new LlmQueuedContext(); var original = SynchronizationContext.Current; CopilotClient client; try { SynchronizationContext.SetSynchronizationContext(queue); client = new CopilotClient(); } finally { SynchronizationContext.SetSynchronizationContext(original); }
                    using (client)
                    {
                        int calls = 0; var chunks = new List<string>(); if (mode != "queued-tool-no-delta") client.TextDelta = chunks.Add; var turn = client.CompleteAsync("model", History(), Tools(), (n, a) => { calls++; return Task.FromResult("unused"); });
                        if (mode.StartsWith("queued-tool")) { await turn; queue.Drain(); Assert.AreEqual(0, calls); CollectionAssert.AreEqual(mode == "queued-tool" ? new[] { "delta été" } : new string[0], chunks); }
                        else { var deadline = DateTime.UtcNow.AddSeconds(5); while (queue.Count < 2 && DateTime.UtcNow < deadline) await Task.Delay(5); Assert.IsTrue(queue.Count >= 2); client.Dispose(); queue.Drain(); await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await turn); Assert.AreEqual(0, calls); }
                    }
                }
                scope.UseCopilot(); using (var client = Client()) { client.TextDelta = fragment => client.Dispose(); await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.CompleteAsync("model", History(), Tools(), (n, a) => Task.FromResult("result"))); }
                using (var client = Client())
                {
                    LlmBoundaryScope.Call(client, "Dispatch", new object[] { null }); Assert.AreEqual("denied-by-rules", LlmBoundaryScope.Call(client, "PermissionDecision", LlmBoundaryScope.Object(new { kind = "custom-tool", toolName = "read_module" }))); client.Dispose(); Assert.AreEqual("denied-by-rules", LlmBoundaryScope.Call(client, "PermissionDecision", LlmBoundaryScope.Object(new { kind = "custom-tool", toolName = "read_module" })));
                }
            }
        }

        /// <summary>RPC errors close only the captured turn before a following idle frame can report success.</summary>
        [TestMethod]
        public void RpcErrorsSettleCapturedTurnBeforeIdleWithoutFaultingReplacementOwners()
        {
            foreach (string state in new[] { "active", "new-session", "new-turn", "disposed" })
                using (var client = Client())
                {
                    var original = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                    var source = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
                    LlmBoundaryScope.Set(client, "sessionId", "original"); LlmBoundaryScope.Set(client, "completion", original);
                    var pending = LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "pending");
                    var owners = LlmBoundaryScope.Get<Dictionary<int, Tuple<string, TaskCompletionSource<string>>>>(client, "pendingOwners");
                    pending.Add(1, source); owners.Add(1, Tuple.Create("original", original));
                    var current = original; string session = "original";
                    if (state.StartsWith("new-", StringComparison.Ordinal))
                    {
                        current = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                        session = state == "new-session" ? "newer" : "original";
                        LlmBoundaryScope.Set(client, "sessionId", session); LlmBoundaryScope.Set(client, "completion", current);
                    }
                    if (state == "disposed") client.Dispose();
                    LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { id = 1, error = new { } }));
                    Assert.IsTrue(source.Task.IsFaulted, state);
                    if (state == "active") Assert.ThrowsException<InvalidOperationException>(() => original.Task.GetAwaiter().GetResult());
                    if (state.StartsWith("new-", StringComparison.Ordinal))
                    {
                        Assert.IsFalse(original.Task.IsCompleted, state);
                        Assert.IsFalse(current.Task.IsCompleted, state);
                        Assert.AreEqual(false, LlmBoundaryScope.Call(client, "TryFaultTurn", "original", original, new IOException("delayed old permission/tool failure")), state);
                    }
                    if (state == "disposed") Assert.AreEqual(0, owners.Count);
                    else
                    {
                        LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { method = "session.event", @params = new { sessionId = session, @event = new { type = "session.idle", data = new { } } } }));
                        Assert.AreEqual(state == "active", current.Task.IsFaulted, state);
                        if (state != "active") Assert.AreEqual("", current.Task.GetAwaiter().GetResult());
                    }
                    Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
                }
        }

        /// <summary>A failed scoped request removes its metadata even when no transport could be written.</summary>
        [TestMethod]
        public async Task ScopedRequestFailureRemovesRpcOwnerMetadata()
        {
            using (var client = Client())
            {
                var owner = new TaskCompletionSource<string>();
                LlmBoundaryScope.Set(client, "sessionId", "owned"); LlmBoundaryScope.Set(client, "completion", owner);
                var request = (Task<IDictionary<string, object>>)LlmBoundaryScope.Call(client, "RequestCoreAsync", "owned.synthetic.request", new { }, "owned", owner);
                // A cold client has no transport; this boundary must not leave a registered RPC owner behind.
                await Assert.ThrowsExceptionAsync<NullReferenceException>(() => request);
                Assert.AreEqual(0, LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "pending").Count);
                Assert.AreEqual(0, LlmBoundaryScope.Get<Dictionary<int, Tuple<string, TaskCompletionSource<string>>>>(client, "pendingOwners").Count);
                Assert.IsFalse(owner.Task.IsCompleted);
                Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
            }
        }

        /// <summary>Delayed permission/tool failure forwarding is rejected after owner replacement or disposal.</summary>
        [TestMethod]
        public void QueuedScopedFailureCannotPoisonANewerTurnOrDisposedClient()
        {
            foreach (bool dispose in new[] { false, true })
                using (var client = Client())
                {
                    var queue = new LlmQueuedContext();
                    var original = new TaskCompletionSource<string>(); var newer = new TaskCompletionSource<string>();
                    LlmBoundaryScope.Set(client, "sessionId", "original"); LlmBoundaryScope.Set(client, "completion", original);
                    bool accepted = true;
                    queue.Post(_ => accepted = (bool)LlmBoundaryScope.Call(client, "TryFaultTurn", "original", original, new IOException("old queued failure")), null);
                    LlmBoundaryScope.Set(client, "sessionId", "newer"); LlmBoundaryScope.Set(client, "completion", newer);
                    if (dispose) client.Dispose();
                    queue.Drain();
                    Assert.IsFalse(accepted); Assert.IsFalse(original.Task.IsCompleted);
                    if (dispose) Assert.ThrowsException<OperationCanceledException>(() => newer.Task.GetAwaiter().GetResult());
                    else Assert.IsFalse(newer.Task.IsCompleted);
                    Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
                }
        }
        /// <summary>A duplicate error cannot undo an already successful RPC while its finally continuation is queued.</summary>
        [TestMethod]
        public void DuplicateErrorAfterRpcSuccessCannotFaultItsStillActiveOwner()
        {
            using (var client = Client())
            {
                var owner = new TaskCompletionSource<string>();
                var rpc = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
                LlmBoundaryScope.Set(client, "sessionId", "owned"); LlmBoundaryScope.Set(client, "completion", owner);
                LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "pending").Add(1, rpc);
                LlmBoundaryScope.Get<Dictionary<int, Tuple<string, TaskCompletionSource<string>>>>(client, "pendingOwners").Add(1, Tuple.Create("owned", owner));
                var events = new List<CodexAgentActivity>(); client.ActivityUpdate = events.Add;
                rpc.SetResult(new Dictionary<string, object>());
                LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { id = 1, error = new { } }));
                Assert.AreEqual(TaskStatus.RanToCompletion, rpc.Task.Status); Assert.IsFalse(owner.Task.IsCompleted);
                Assert.AreEqual(0, events.Count);
            }
        }

        /// <summary>Failed receipts are reserved before completion wakes cleanup and survive disposal of the same owner.</summary>
        [TestMethod]
        public void FailedActivityReservationSurvivesQueuedDisposeButNotOwnerReplacement()
        {
            foreach (bool replace in new[] { false, true })
            {
                var queue = new LlmQueuedContext(); var previous = SynchronizationContext.Current;
                CopilotClient client;
                try { SynchronizationContext.SetSynchronizationContext(queue); client = new CopilotClient(); }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
                using (client)
                {
                    var owner = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                    LlmBoundaryScope.Set(client, "sessionId", "owned"); LlmBoundaryScope.Set(client, "completion", owner);
                    var events = new List<CodexAgentActivity>(); client.ActivityUpdate = events.Add;
                    LlmBoundaryScope.Call(client, "PublishActivity", new CodexAgentActivity { Id = "owned:reasoning", Kind = "reasoning", Status = "inProgress", Detail = "owned" });
                    queue.Drain(); Assert.AreEqual(1, events.Count);
                    var error = new IOException("owned failed RPC");
                    Assert.AreEqual(true, LlmBoundaryScope.Call(client, "TryFaultTurn", "owned", owner, error));
                    Assert.IsTrue(owner.Task.IsFaulted);
                    Assert.AreEqual("failed", LlmBoundaryScope.Get<Dictionary<string, CodexAgentActivity>>(client, "activitySnapshots")["owned:reasoning"].Status);
                    if (replace) { LlmBoundaryScope.Set(client, "sessionId", "newer"); LlmBoundaryScope.Set(client, "completion", new TaskCompletionSource<string>()); }
                    else client.Dispose();
                    LlmBoundaryScope.Call(client, "EndTurnActivities", "owned", owner, error);
                    queue.Drain();
                    Assert.AreEqual(replace ? 1 : 2, events.Count);
                    if (!replace) Assert.AreEqual("failed", events.Last().Status);
                    Assert.AreEqual(0, LlmBoundaryScope.Get<Dictionary<TaskCompletionSource<string>, CodexAgentActivity[]>>(client, "failedTurnActivities").Count);
                }
            }
        }

        /// <summary>Synchronous progress callbacks must not admit a tool after disposing or replacing its turn.</summary>
        [TestMethod]
        public void ToolProgressSubscriberCannotRebindThenAdmitAnOldTool()
        {
            foreach (string state in new[] { "disposed", "new-turn", "completed" })
                using (var client = Client())
                {
                    var owner = new TaskCompletionSource<string>();
                    LlmBoundaryScope.Set(client, "sessionId", "owned"); LlmBoundaryScope.Set(client, "completion", owner);
                    LlmBoundaryScope.Set(client, "allowedTools", new HashSet<string> { "read_module" });
                    var history = new List<object>(); LlmBoundaryScope.Set(client, "history", history);
                    int calls = 0;
                    LlmBoundaryScope.Set(client, "invoke", (Func<string, string, Task<string>>)((n, a) => { calls++; return Task.FromResult("unused"); }));
                    client.ActivityUpdate = activity =>
                    {
                        if (activity.Status != "inProgress") return;
                        if (state == "disposed") client.Dispose();
                        if (state == "new-turn") LlmBoundaryScope.Set(client, "completion", new TaskCompletionSource<string>());
                        if (state == "completed") owner.TrySetResult("settled by subscriber");
                    };
                    LlmBoundaryScope.Call(client, "RunTool", LlmBoundaryScope.Object(new { requestId = "synthetic", toolName = "read_module" }), null);
                    Assert.AreEqual(0, calls, state); Assert.AreEqual(0, history.Count, state);
                    Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
                }
        }
        /// <summary>A failed tool receipt subscriber cannot escape the owned failure or retain reserved activities.</summary>
        [TestMethod]
        public void ToolFailurePublicationCannotPoisonReplacementOwnerOrPreventReceiptCleanup()
        {
            using (var client = Client())
            {
                var original = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                var newer = new TaskCompletionSource<string>();
                var catalog = new TaskCompletionSource<IDictionary<string, object>>();
                LlmBoundaryScope.Set(client, "sessionId", "original"); LlmBoundaryScope.Set(client, "completion", original);
                LlmBoundaryScope.Set(client, "allowedTools", new HashSet<string> { "read_module" });
                LlmBoundaryScope.Set(client, "history", new List<object>());
                LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "pending").Add(2, catalog);
                int calls = 0, failedReceipts = 0;
                LlmBoundaryScope.Set(client, "invoke", (Func<string, string, Task<string>>)((name, args) => { calls++; throw new InvalidOperationException("Primary tool failure"); }));
                client.ActivityUpdate = activity =>
                {
                    if (activity.Status != "failed") return;
                    failedReceipts++;
                    LlmBoundaryScope.Set(client, "sessionId", "newer"); LlmBoundaryScope.Set(client, "completion", newer);
                    throw new IOException("Failed receipt subscriber");
                };
                LlmBoundaryScope.Call(client, "RunTool", LlmBoundaryScope.Object(new { requestId = "owned", toolName = "read_module" }), null);
                Assert.AreEqual(1, calls); Assert.AreEqual(1, failedReceipts);
                var error = Assert.ThrowsException<InvalidOperationException>(() => original.Task.GetAwaiter().GetResult());
                Assert.AreEqual("Primary tool failure", error.Message);
                Assert.IsFalse(newer.Task.IsCompleted); Assert.IsFalse(catalog.Task.IsCompleted);
                Assert.AreEqual(0, LlmBoundaryScope.Get<Dictionary<TaskCompletionSource<string>, CodexAgentActivity[]>>(client, "failedTurnActivities").Count);
                Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
            }
        }
        private sealed class RejectingActivityContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback callback, object state) { throw new InvalidOperationException("Owned disposed context"); }
        }

        /// <summary>Failed activity publication cannot escape into reader failure and poison another turn or catalog RPC.</summary>
        [TestMethod]
        public void FailedActivityPublicationErrorsPreservePrimaryRpcAndReplacementOwner()
        {
            foreach (string state in new[] { "post-error", "inline-rebind-error", "queued-rebind-error" })
            {
                var queue = new LlmQueuedContext(); var previous = SynchronizationContext.Current;
                CopilotClient client;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(state == "post-error" ? (SynchronizationContext)new RejectingActivityContext() : state == "queued-rebind-error" ? queue : null);
                    client = new CopilotClient();
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
                using (client)
                {
                    var original = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                    var newer = new TaskCompletionSource<string>();
                    var rpc = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
                    var catalog = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
                    LlmBoundaryScope.Set(client, "sessionId", "original"); LlmBoundaryScope.Set(client, "completion", original);
                    var pending = LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "pending");
                    pending.Add(1, rpc); pending.Add(2, catalog);
                    LlmBoundaryScope.Get<Dictionary<int, Tuple<string, TaskCompletionSource<string>>>>(client, "pendingOwners").Add(1, Tuple.Create("original", original));
                    var activity = new CodexAgentActivity { Id = "original:reasoning", Kind = "reasoning", Status = "inProgress", Detail = "owned" };
                    LlmBoundaryScope.Get<Dictionary<string, CodexAgentActivity>>(client, "pendingActivities").Add(activity.Id, activity);
                    LlmBoundaryScope.Get<Dictionary<string, CodexAgentActivity>>(client, "activitySnapshots").Add(activity.Id, activity);
                    client.ActivityUpdate = receipt =>
                    {
                        LlmBoundaryScope.Set(client, "sessionId", "newer"); LlmBoundaryScope.Set(client, "completion", newer);
                        throw new IOException("Owned subscriber failure after rebind");
                    };
                    LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { id = 1, error = new { } }));
                    if (state == "queued-rebind-error") queue.Drain();
                    if (state == "post-error") { LlmBoundaryScope.Set(client, "sessionId", "newer"); LlmBoundaryScope.Set(client, "completion", newer); }
                    Assert.ThrowsException<InvalidOperationException>(() => rpc.Task.GetAwaiter().GetResult(), state);
                    Assert.ThrowsException<InvalidOperationException>(() => original.Task.GetAwaiter().GetResult(), state);
                    Assert.IsFalse(newer.Task.IsCompleted, state); Assert.IsFalse(catalog.Task.IsCompleted, state);
                    Assert.AreEqual("failed", activity.Status, state);
                    Assert.IsNull(LlmBoundaryScope.Get<Process>(client, "process"));
                }
            }
        }
        private static void ActivityEvent(CopilotClient client, string type, object data, string agent = "main")
        {
            LlmBoundaryScope.Call(client, "Dispatch", LlmBoundaryScope.Object(new { method = "session.event", @params = new { sessionId = "fixture", @event = new { type, data, agentId = agent } } }));
        }

        private static CopilotClient ActivityClient(List<CodexAgentActivity> events)
        {
            var client = Client();
            LlmBoundaryScope.Set(client, "sessionId", "fixture");
            LlmBoundaryScope.Set(client, "completion", new TaskCompletionSource<string>());
            client.ActivityUpdate = events.Add;
            return client;
        }

        [TestMethod]
        public void PublicReasoningUsesAuthoritativeSnapshotsAndTurnAgentIdentityWithoutOpaqueContent()
        {
            var events = new List<CodexAgentActivity>();
            using (var client = ActivityClient(events))
            {
                ActivityEvent(client, "assistant.turn_start", new { turnId = "turn1" });
                ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "block", deltaContent = "Plan" });
                ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "block", deltaContent = " first" });
                ActivityEvent(client, "assistant.reasoning", new { reasoningId = "block", content = "Plan final", encryptedContent = "secret", reasoningOpaque = "secret" });
                ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "block", deltaContent = " late" });
                ActivityEvent(client, "assistant.message", new { messageId = "message1", content = "answer", reasoningText = "Plan final", reasoningOpaque = "secret" });
                Assert.AreEqual(3, events.Count);
                Assert.IsTrue(events[0].Append); Assert.IsTrue(events[1].Append); Assert.IsFalse(events[2].Append);
                Assert.AreEqual("Plan final", events[2].Detail); Assert.AreEqual("completed", events[2].Status);
                Assert.AreEqual(events[0].Id, events[2].Id);
                ActivityEvent(client, "assistant.turn_start", new { turnId = "turn2" });
                ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "block", deltaContent = "Next" });
                Assert.AreNotEqual(events[0].Id, events[3].Id, "A provider may reuse a local reasoning block identifier across native turns.");
                ActivityEvent(client, "assistant.reasoning", new { reasoningId = "block", content = "Child" }, "child");
                Assert.AreNotEqual(events[3].Id, events[4].Id);
                ActivityEvent(client, "assistant.message", new { messageId = "opaque", content = "answer", reasoningOpaque = "secret", encryptedContent = "secret" });
                ActivityEvent(client, "assistant.reasoning", new { reasoningId = "invalid", content = 123 });
                Assert.AreEqual(5, events.Count);
                Assert.IsFalse(events.Any(e => (e.Detail ?? "").Contains("secret")));
                ActivityEvent(client, "session.idle", new { });
                Assert.AreEqual("interrupted", events.Last().Status);
                Assert.IsTrue(events.Last().Append); Assert.AreEqual("", events.Last().Detail);
            }
        }

        [TestMethod]
        public void MessageFallbackAndTerminalFailureCancellationSettlePartialReasoningWithoutDuplicatingText()
        {
            foreach (var ending in new[] { "session.error", "session.idle", "dispose" })
            {
                var events = new List<CodexAgentActivity>();
                using (var client = ActivityClient(events))
                {
                    ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "partial", deltaContent = "Published" });
                    if (ending == "dispose") client.Dispose(); else ActivityEvent(client, ending, new { });
                    Assert.AreEqual(2, events.Count, ending);
                    Assert.AreEqual(ending == "session.error" ? "failed" : "interrupted", events[1].Status);
                    Assert.AreEqual(events[0].Id, events[1].Id); Assert.IsTrue(events[1].Append); Assert.AreEqual("", events[1].Detail);
                    ActivityEvent(client, "assistant.reasoning_delta", new { reasoningId = "late", deltaContent = "late" });
                    Assert.AreEqual(2, events.Count);
                    var task = LlmBoundaryScope.Get<TaskCompletionSource<string>>(client, "completion").Task;
                    if (task.IsFaulted) Assert.IsNotNull(task.Exception);
                }
            }
            var fallback = new List<CodexAgentActivity>();
            using (var client = ActivityClient(fallback))
            {
                ActivityEvent(client, "assistant.message", new { messageId = "message", content = "answer", reasoningText = "Published fallback" });
                ActivityEvent(client, "assistant.message", new { messageId = "message", content = "answer", reasoningText = "Published fallback" });
                Assert.AreEqual(1, fallback.Count); Assert.AreEqual("completed", fallback[0].Status); Assert.IsFalse(fallback[0].Append);
                ActivityEvent(client, "session.idle", new { });
            }
        }

        [TestMethod]
        public async Task ActualToolReceiptWinsOverTransportDeliveryAndDuplicateRequestsExecuteOnce()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var receipt in new[] { "{\"Ok\":true}", "{\"Ok\":false,\"Error\":\"Refused\"}", "{\"Ok\":false,\"Data\":{\"Uncertain\":true,\"Reason\":\"Inspect state\"}}" })
                {
                    scope.UseCopilot();
                    using (var client = Client())
                    {
                        var events = new List<CodexAgentActivity>(); client.ActivityUpdate = events.Add;
                        int calls = 0;
                        var history = History();
                        var result = await client.CompleteAsync("model", history, Tools(), (name, args) => { calls++; return Task.FromResult(receipt); });
                        Assert.AreEqual(1, calls); Assert.AreEqual(2, events.Count);
                        Assert.AreEqual("inProgress", events[0].Status); Assert.AreEqual(ProviderActivityProjection.ToolOutcome(receipt), events[1].Status);
                        Assert.AreEqual(events[0].Id, events[1].Id); Assert.AreEqual("answer été", result["content"]);
                        Assert.AreEqual(4, history.Count);
                    }
                }
                scope.UseCopilot("tool-reply-error");
                using (var client = Client())
                {
                    var events = new List<CodexAgentActivity>(); client.ActivityUpdate = events.Add;
                    await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync("model", History(), Tools(), (n, a) => Task.FromResult("{\"Ok\":true}")));
                    Assert.AreEqual(2, events.Count); Assert.AreEqual("completed", events.Last().Status, "A delivery failure must not rewrite the actual completed VBAi receipt.");
                }
            }
        }

        [TestMethod]
        public void QueuedToolsAndActivitiesCannotCrossIntoAReplacementNativeSession()
        {
            var queue = new LlmQueuedContext(); var previousContext = SynchronizationContext.Current;
            CopilotClient client;
            try { SynchronizationContext.SetSynchronizationContext(queue); client = new CopilotClient(); }
            finally { SynchronizationContext.SetSynchronizationContext(previousContext); }
            using (client)
            {
                var firstCompletion = new TaskCompletionSource<string>();
                LlmBoundaryScope.Set(client, "sessionId", "first"); LlmBoundaryScope.Set(client, "completion", firstCompletion);
                var firstEvents = new List<CodexAgentActivity>(); client.ActivityUpdate = firstEvents.Add;
                int calls = 0;
                LlmBoundaryScope.Set(client, "allowedTools", new HashSet<string> { "read_module" });
                LlmBoundaryScope.Set(client, "invoke", (Func<string, string, Task<string>>)((name, arguments) => { calls++; return Task.FromResult("{\"Ok\":true}"); }));
                var history = new List<object>(); LlmBoundaryScope.Set(client, "history", history);
                LlmBoundaryScope.Call(client, "PublishActivity", new CodexAgentActivity { Id = "first:reasoning", Kind = "reasoning", Detail = "Published first", Status = "inProgress" });
                LlmBoundaryScope.Call(client, "RunTool", LlmBoundaryScope.Object(new { requestId = "queued", toolName = "read_module" }), "old-rpc");
                firstCompletion.SetResult("first complete");
                var secondCompletion = new TaskCompletionSource<string>();
                LlmBoundaryScope.Set(client, "sessionId", "second"); LlmBoundaryScope.Set(client, "completion", secondCompletion);
                var secondEvents = new List<CodexAgentActivity>(); client.ActivityUpdate = secondEvents.Add;
                queue.Drain();
                Assert.AreEqual(0, calls, "A deferred callback cannot execute an old tool against a new native session.");
                Assert.AreEqual(0, history.Count); Assert.IsFalse(secondCompletion.Task.IsCompleted);
                Assert.AreEqual(0, firstEvents.Count); Assert.AreEqual(0, secondEvents.Count);
                secondCompletion.SetResult("second complete");
            }
        }
    }
}
