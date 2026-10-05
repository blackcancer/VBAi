using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
    }
}
