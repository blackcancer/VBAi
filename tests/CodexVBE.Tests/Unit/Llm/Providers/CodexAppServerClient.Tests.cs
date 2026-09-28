namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE.Tests.Infrastructure;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;

    /// <summary>Vérifie le client app-server Codex avec un transport RPC simulé et sans processus externe.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class CodexAppServerClientTests
    {
        /// <summary>Refuse les fournisseurs, catalogues et transports qui ne respectent pas le contrat requis.</summary>
        /// <returns>Tâche terminée après les vérifications asynchrones.</returns>
        [TestMethod]
        public async Task ConstructorCatalogueAndTransportGuardsRejectIncompleteContracts()
        {
            var settings = new LlmSettings(); var tools = new LlmVbeTools(null, null, settings); var context = new ImmediateContext();
            Assert.ThrowsException<ArgumentNullException>(() => new CodexAppServerClient(null, tools, null, settings, null, new FakeTransport()));
            Assert.ThrowsException<ArgumentNullException>(() => new CodexAppServerClient(context, null, null, settings, null, new FakeTransport()));
            Assert.ThrowsException<ArgumentNullException>(() => new CodexAppServerClient(context, tools, null, null, null, new FakeTransport()));
            Assert.ThrowsException<ArgumentNullException>(() => new CodexAppServerClient(context, tools, null, settings, null, null));
            using (var native = new CodexAppServerClient(context, tools, null, settings)) Assert.IsNull(native.ThreadId);
            var transport = new FakeTransport(); using (var client = Client(transport))
            {
                await client.InterruptAsync();
                foreach (var method in new[] { "RequestAsync", "Send" }) { var error = Assert.ThrowsException<TargetInvocationException>(() => LlmBoundaryScope.Call(client, method, method == "Send" ? new object[] { new { } } : new object[] { "fixture", new { } })); Assert.IsInstanceOfType<InvalidOperationException>(error.InnerException); }
                transport.Intercept = message => { if (Method(message) != "model/list") return false; transport.Emit(new { id = message["id"], result = new { data = new object[] { null, new { }, new { model = " " }, new { model = "one", supportedReasoningEfforts = (object)null }, new { model = "two", supportedReasoningEfforts = new object[] { null, new { }, new { reasoningEffort = " " }, new { reasoningEffort = "high", description = "High" } } } } } }); return true; };
                var models = await client.ListModelsAsync(); Assert.AreEqual(2, models.Length); Assert.AreEqual(0, models[0].Efforts.Length); Assert.AreEqual("high", models[1].Efforts.Single().Id);
                transport.EmitRaw("null"); transport.Emit(new { id = "nonnumeric" }); transport.Emit(new { id = 99999 });
                transport.Stop(); var requestError = Assert.ThrowsException<TargetInvocationException>(() => LlmBoundaryScope.Call(client, "RequestAsync", "fixture", new { })); Assert.IsInstanceOfType<InvalidOperationException>(requestError.InnerException);
                var sendError = Assert.ThrowsException<TargetInvocationException>(() => LlmBoundaryScope.Call(client, "Send", new { })); Assert.IsInstanceOfType<InvalidOperationException>(sendError.InnerException); client.Dispose(); client.Dispose();
            }
            transport = new FakeTransport(); using (var client = Client(transport)) { await client.ListModelsAsync(); transport.BeforeSend = () => throw new IOException("native write failure"); await Assert.ThrowsExceptionAsync<IOException>(() => client.ListModelsAsync()); transport.BeforeSend = null; Assert.AreEqual(0, LlmBoundaryScope.Get<Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>>(client, "requests").Count); }
            transport = new FakeTransport(); using (var client = Client(transport)) { await client.ListModelsAsync(); transport.Intercept = m => { if (Method(m) != "model/list") return false; transport.Emit(new { id = m["id"], error = new { } }); return true; }; var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync()); Assert.AreEqual("Codex request failed.", error.Message); }
        }

        /// <summary>Respecte les deux frontières d’annulation durant l’initialisation et le démarrage d’un tour.</summary>
        /// <returns>Tâche terminée après vérification de l’annulation.</returns>
        [TestMethod]
        public async Task InterruptedInitializationAndTurnStartRespectBothCancellationBoundaries()
        {
            foreach (bool initializing in new[] { true, false })
            {
                var transport = new FakeTransport(); object pendingId = null; transport.Intercept = m => { if (Method(m) != (initializing ? "initialize" : "turn/start")) return false; pendingId = m["id"]; return true; };
                using (var client = Client(transport)) { var turn = client.TurnAsync("request", null, null); Assert.IsNotNull(pendingId); await client.InterruptAsync(); transport.Emit(new { id = pendingId, result = initializing ? (object)new { } : new { turn = new { id = "turn-1" } } }); await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => turn); Assert.AreEqual(!initializing, transport.Methods.Contains("turn/interrupt")); }
            }
            var alternate = new FakeTransport(); alternate.Intercept = m => { if (Method(m) != "turn/start") return false; alternate.Emit(new { method = "turn/started", @params = new { threadId = "thread-1", turn = new { id = "event-turn" } } }); alternate.Emit(new { id = m["id"], result = new { } }); alternate.EmitTurnCompleted("completed", null); return true; }; using (var client = Client(alternate)) Assert.IsFalse(string.IsNullOrWhiteSpace(await client.TurnAsync("request", null, null)));
        }

        /// <summary>Traite les notifications et ignore les mises à jour en attente après libération.</summary>
        /// <returns>Tâche terminée après traitement des notifications.</returns>
        [TestMethod]
        public async Task AllNotificationShapesAndQueuedDisposalKeepUpdatesWithinTheActiveTurn()
        {
            var transport = new FakeTransport(); using (var client = Client(transport))
            {
                var updates = new List<string>(); client.ChatUpdate += (kind, id, text, complete) => updates.Add(kind + ":" + text); var turn = client.TurnAsync("request", null, null); await transport.TurnStarted.Task;
                foreach (var index in new object[] { null, 0, 1 }) transport.Emit(new { method = "item/reasoning/summaryPartAdded", @params = index == null ? (object)new { threadId = "thread-1", itemId = "r" } : new { threadId = "thread-1", itemId = "r", summaryIndex = index } });
                transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread-1", delta = "no item" } });
                foreach (var summary in new object[] { null, "invalid", new object[0], new object[] { "text", new { text = "more" }, null, 5, new { text = " " } } }) transport.Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "reasoning", id = "r", summary } } });
                transport.Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "reasoning", id = "r" } } });
                transport.Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "agentMessage", phase = "commentary", id = "m", text = "intermediate" } } });
                transport.Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "other" } } }); transport.Emit(new { method = "unknown", @params = new { threadId = "thread-1" } }); transport.EmitTurnCompleted("failed", null); var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => turn); Assert.AreEqual("Codex turn: failed", error.Message); Assert.IsTrue(updates.Contains("message:intermediate")); Assert.IsTrue(updates.Contains("summary:text\n\nmore"));
            }
            foreach (bool dispose in new[] { false, true }) { transport = new FakeTransport(); var queue = new LlmQueuedContext(); var settings = new LlmSettings(); using (var client = new CodexAppServerClient(queue, new LlmVbeTools(null, null, settings), null, settings, null, transport)) { var turn = client.TurnAsync("request", null, null); await transport.TurnStarted.Task; transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread-1", itemId = "m", delta = "queued" } }); if (dispose) client.Dispose(); queue.Drain(); if (dispose) await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => turn); else { transport.EmitTurnCompleted("completed", null); queue.Drain(); await turn; } } }
            transport = new FakeTransport(); using (var client = Client(transport)) { var turn = client.TurnAsync("request", null, null); await transport.TurnStarted.Task; transport.Emit(new { method = "item/reasoning/summaryPartAdded", @params = new { threadId = "thread-1", summaryIndex = "invalid" } }); await Assert.ThrowsExceptionAsync<FormatException>(() => turn); }
        }

        /// <summary>Transforme les résultats ou exceptions d’outils sans appeler un véritable hôte VBE.</summary>
        /// <returns>Tâche terminée après les appels d’outil simulés.</returns>
        [TestMethod]
        public async Task ToolBoundaryResponsesExceptionsAndLateCallbacksNeverCallARealHost()
        {
            foreach (var output in new[] { "null", "{\"Ok\":true}", "{\"Ok\":false}", "not-json", "throw" }) foreach (bool observe in new[] { false, true })
            {
                var transport = new FakeTransport(); using (var client = Client(transport)) { int calls = 0; var updates = new List<string>(); if (observe) client.ChatUpdate += (k, i, t, c) => updates.Add(t); client.InvokeTool = (n, a) => { calls++; Assert.AreEqual("fixture", n); Assert.AreEqual("{}", a); return output == "throw" ? Task.FromException<string>(new IOException("tool boundary")) : Task.FromResult(output); }; var turn = client.TurnAsync("request", null, null); await transport.TurnStarted.Task; transport.Emit(new { id = "tool", method = "item/tool/call", @params = new { threadId = "thread-1", tool = "fixture", arguments = new { } } }); var reply = transport.Sent.Last(); Assert.AreEqual("tool", reply["id"]); Assert.AreEqual(output.Contains("true"), FakeTransport.Object(reply["result"])["success"]); Assert.AreEqual(1, calls); if (observe) Assert.IsTrue(updates.Count >= 1); transport.EmitTurnCompleted("completed", null); await turn; }
            }
            var idle = new FakeTransport(); using (var client = Client(idle)) { await client.ListModelsAsync(); idle.Emit(new { id = "idle-tool", method = "item/tool/call", @params = new { threadId = "thread-1" } }); Assert.AreEqual(false, FakeTransport.Object(idle.Sent.Last()["result"])["success"]); }
            foreach (bool dispose in new[] { false, true }) { var transport = new FakeTransport(); var queue = new LlmQueuedContext(); var settings = new LlmSettings(); using (var client = new CodexAppServerClient(queue, new LlmVbeTools(null, null, settings), null, settings, null, transport)) { int calls = 0; client.InvokeTool = (n, a) => { calls++; return Task.FromResult("null"); }; var turn = client.TurnAsync("request", null, null); await transport.TurnStarted.Task; transport.Emit(new { id = "late", method = "item/tool/call", @params = new { threadId = "thread-1", tool = "fixture", arguments = new { } } }); if (dispose) client.Dispose(); else await client.InterruptAsync(); queue.Drain(); Assert.AreEqual(0, calls); if (dispose) await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => turn); else await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => turn); } }
            var pending = new FakeTransport(); using (var client = Client(pending)) { await client.ListModelsAsync(); pending.Intercept = m => Method(m) == "model/list"; var request = client.ListModelsAsync(); client.Dispose(); await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => request); }
        }

        /// <summary>Vérifie le transport de processus avec flux UTF-8, entrée protocolaire et vidage de stderr.</summary>
        /// <returns>Tâche terminée après arrêt du processus fixture.</returns>
        [TestMethod]
        public async Task NativeProcessTransportUsesDisposableUtf8ChildAndDrainsOnlyProtocolOutput()
        {
            using (var fixture = new NativeProtocolFixture()) using (var scope = new LlmBoundaryScope())
            {
                foreach (var installed in new[] { false, true }) foreach (var explicitPath in new[] { false, true })
                {
                    Environment.SetEnvironmentVariable("CODEXVBE_CODEX_CLI", explicitPath ? "fixture-explicit.exe" : null); using (var transport = new CodexProcessTransport())
                    {
                        Assert.IsFalse(transport.IsRunning); Assert.ThrowsException<InvalidOperationException>(() => transport.Send("before")); string chosen = null; transport.InstalledExists = p => { StringAssert.EndsWith(p, Path.Combine("Codex", "bin", "codex.exe")); return installed; }; transport.StartProcess = p => { chosen = p.StartInfo.FileName; return fixture.Start(p, "echo"); };
                        var output = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously); var exited = new TaskCompletionSource<Exception>(TaskCreationOptions.RunContinuationsAsynchronously); transport.LineReceived += line => output.TrySetResult(line); transport.Exited += ex => exited.TrySetResult(ex); transport.Start(); Assert.IsTrue(transport.IsRunning); Assert.AreEqual(explicitPath ? "fixture-explicit.exe" : installed ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OpenAI", "Codex", "bin", "codex.exe") : "codex.exe", chosen); transport.Send("équipe"); Assert.AreEqual("équipe", await output.Task); transport.Send("exit"); Assert.IsInstanceOfType<InvalidOperationException>(await exited.Task); Assert.IsFalse(transport.IsRunning); transport.Dispose(); transport.Dispose();
                    }
                }
                using (var transport = new CodexProcessTransport()) { transport.StartProcess = p => false; Assert.ThrowsException<InvalidOperationException>(() => transport.Start()); transport.Dispose(); }
                using (var transport = new CodexProcessTransport()) { transport.StartProcess = p => fixture.Start(p, "echo"); transport.Start(); transport.Send("unobserved"); var process = LlmBoundaryScope.Get<Process>(transport, "process"); transport.Send("exit"); Assert.IsTrue(process.WaitForExit(5000)); transport.Dispose(); }
                using (var transport = new CodexProcessTransport()) { transport.StartProcess = p => fixture.Start(p, "echo"); transport.Start(); var process = LlmBoundaryScope.Get<Process>(transport, "process"); process.Kill(); Assert.IsTrue(process.WaitForExit(5000)); process.Close(); transport.Dispose(); }
                using (var transport = new CodexProcessTransport()) { transport.StartProcess = p => fixture.Start(p, "echo"); transport.Start(); transport.Dispose(); }
            }
        }
        /// <summary>Initialise une session ChatGPT après validation du compte et des outils en lecture seule.</summary>
        /// <returns>Tâche terminée après l’initialisation.</returns>
        [TestMethod]
        public async Task InitializeChecksChatGptAndStartsReadOnlyThread()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                string ready = null;
                client.ThreadReady += id => ready = id;
                await client.ListModelsAsync();
                Assert.AreEqual("thread-1", client.ThreadId);
                Assert.AreEqual("thread-1", ready);
                CollectionAssert.AreEqual(new[] { "initialize", "initialized", "account/read", "thread/start", "model/list" }, transport.Methods.ToArray());
                var thread = transport.Request("thread/start");
                var arguments = FakeTransport.Object(thread["params"]);
                Assert.AreEqual("read-only", arguments["sandbox"]);
                Assert.AreEqual("untrusted", arguments["approvalPolicy"]);
                Assert.IsTrue(((object[])arguments["dynamicTools"]).Length > 0);
            }

            Assert.IsTrue(transport.Disposed);
        }

        /// <summary>Reprend un fil existant en conservant exactement son identifiant.</summary>
        /// <returns>Tâche terminée après la reprise du fil.</returns>
        [TestMethod]
        public async Task ExistingThreadIsResumedWithItsExactIdentifier()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport, "saved-thread"))
            {
                await client.ListModelsAsync();
                Assert.AreEqual("saved-thread", client.ThreadId);
                Assert.IsTrue(transport.Methods.Contains("thread/resume"));
                Assert.IsFalse(transport.Methods.Contains("thread/start"));
                var arguments = FakeTransport.Object(transport.Request("thread/resume")["params"]);
                Assert.AreEqual("saved-thread", arguments["threadId"]);
                Assert.AreEqual("read-only", arguments["sandbox"]);
            }
        }

        /// <summary>Arrête l’initialisation d’un compte non ChatGPT avant l’ouverture d’un fil.</summary>
        /// <returns>Tâche terminée après le refus d’accès.</returns>
        [TestMethod]
        public async Task NonChatGptAccountStopsBeforeOpeningThread()
        {
            var transport = new FakeTransport
            {
                AccountType = "apiKey"
            };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "ChatGPT");
                Assert.IsFalse(transport.Methods.Contains("thread/start"));
                Assert.IsTrue(transport.Disposed);
            }
        }

        /// <summary>Refuse un fil serveur auquel manque son identité obligatoire.</summary>
        /// <returns>Tâche terminée après la réponse d’échec.</returns>
        [TestMethod]
        public async Task MissingThreadIdentityFailsClosed()
        {
            var transport = new FakeTransport
            {
                ReturnEmptyThread = true
            };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "did not create a thread");
                Assert.IsTrue(transport.Disposed);
            }
        }

        /// <summary>Parcourt les pages du catalogue et lit les métadonnées d’effort du modèle.</summary>
        /// <returns>Tâche terminée après lecture du catalogue.</returns>
        [TestMethod]
        public async Task ModelCataloguePaginatesAndReadsEffortMetadata()
        {
            var transport = new FakeTransport
            {
                PaginateModels = true
            };
            using (var client = Client(transport))
            {
                var models = await client.ListModelsAsync();
                Assert.AreEqual(2, models.Length);
                Assert.AreEqual("gpt-alpha", models[0].Id);
                Assert.IsTrue(models[0].IsDefault);
                Assert.AreEqual("high", models[0].DefaultEffort);
                Assert.AreEqual(2, models[0].Efforts.Length);
                Assert.AreEqual("medium", models[0].Efforts[0].Id);
                Assert.AreEqual("gpt-beta", models[1].Id);
                Assert.AreEqual(2, transport.Methods.Count(x => x == "model/list"));
                var second = transport.Sent.Where(x => Convert.ToString(x["method"]) == "model/list").Last();
                Assert.AreEqual("page-2", FakeTransport.Object(second["params"])["cursor"]);
            }
        }

        /// <summary>Refuse un catalogue dépourvu des données de modèle requises.</summary>
        /// <returns>Tâche terminée après validation du catalogue.</returns>
        [TestMethod]
        public async Task MissingCatalogueDataIsRejected()
        {
            var transport = new FakeTransport
            {
                MissingModelData = true
            };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync());
                Assert.IsFalse(string.IsNullOrWhiteSpace(error.Message));
            }
        }

        /// <summary>Propage l’erreur du serveur tout en gardant utilisable le fil déjà prêt.</summary>
        /// <returns>Tâche terminée après la requête refusée.</returns>
        [TestMethod]
        public async Task RequestErrorPropagatesServerMessageWithoutKillingReadyThread()
        {
            var transport = new FakeTransport
            {
                ModelRequestError = true
            };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync());
                Assert.AreEqual("model unavailable", error.Message);
                Assert.AreEqual("thread-1", client.ThreadId);
                Assert.IsFalse(transport.Disposed);
            }
        }

        /// <summary>Libère le client si le transport échoue au démarrage sans envoyer initialize.</summary>
        /// <returns>Tâche terminée après vérification du démarrage avorté.</returns>
        [TestMethod]
        public async Task TransportStartFailureDisposesClientAndDoesNotSendInitialize()
        {
            var transport = new FakeTransport
            {
                FailStart = true
            };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "start refused");
                Assert.AreEqual(0, transport.Sent.Count);
                Assert.IsTrue(transport.Disposed);
            }
        }

        /// <summary>Publie les mises à jour de texte en flux puis le texte final du tour.</summary>
        /// <returns>Tâche terminée après réception du résultat.</returns>
        [TestMethod]
        public async Task TurnPublishesStreamingUpdatesAndFinalText()
        {
            var transport = new FakeTransport
            {
                CompleteTurn = true
            };
            using (var client = Client(transport))
            {
                var updates = new List<string>();
                client.ChatUpdate += (kind, id, text, complete) => updates.Add(kind + ":" + id + ":" + text + ":" + complete);
                string answer = await client.TurnAsync("Explain this module", "gpt-alpha", "high");
                Assert.AreEqual("Final answer", answer);
                Assert.IsTrue(updates.Any(x => x.StartsWith("message:msg-1:partial:")));
                Assert.IsTrue(updates.Any(x => x.StartsWith("summary:reason-1:thinking:")));
                Assert.IsTrue(updates.Any(x => x.StartsWith("summary:reason-1:\n\n:")));
                Assert.IsTrue(updates.Any(x => x.StartsWith("summary:reason-1:summary complete:True")));
                Assert.IsTrue(updates.Any(x => x.StartsWith("final:msg-1:Final answer:True")));
                var arguments = FakeTransport.Object(transport.Request("turn/start")["params"]);
                Assert.AreEqual("thread-1", arguments["threadId"]);
                Assert.AreEqual("gpt-alpha", arguments["model"]);
                Assert.AreEqual("high", arguments["effort"]);
            }
        }

        /// <summary>Expose une erreur native de tour et utilise le repli lorsque le texte final manque.</summary>
        /// <returns>Tâche terminée après réception de la réponse.</returns>
        [TestMethod]
        public async Task FailedTurnSurfacesNativeErrorAndMissingFinalTextUsesFallback()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var failed = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.EmitTurnCompleted("failed", "quota exceeded");
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => failed);
                Assert.AreEqual("quota exceeded", error.Message);
                var completed = client.TurnAsync("second", null, null);
                transport.EmitTurnCompleted("completed", null);
                var answer = await completed;
                Assert.IsFalse(string.IsNullOrWhiteSpace(answer));
                Assert.AreNotEqual("quota exceeded", answer);
            }
        }

        /// <summary>Empêche les notifications reçues sur un autre thread de modifier la conversation.</summary>
        /// <returns>Tâche terminée après vérification de la notification.</returns>
        [TestMethod]
        public async Task NotificationsFromAnotherThreadDoNotReachChatUpdate()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                int updates = 0;
                client.ChatUpdate += (kind, id, text, complete) => updates++;
                await client.ListModelsAsync();
                transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread-1", itemId = "ignored", delta = "before turn" } });
                var turn = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "other", itemId = "ignored", delta = "foreign" } });
                transport.EmitTurnCompleted("completed", null);
                await turn;
                Assert.AreEqual(0, updates);
            }
        }

        /// <summary>Fait échouer le tour en attente lorsqu’une ligne serveur est mal formée.</summary>
        /// <returns>Tâche terminée après traitement de la ligne.</returns>
        [TestMethod]
        public async Task MalformedServerLineFailsPendingTurn()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var turn = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.EmitRaw("{bad json");
                try
                {
                    await turn;
                    Assert.Fail("Malformed JSON must fail the pending turn.");
                }
                catch (ArgumentException)
                {
                }
                catch (InvalidOperationException)
                {
                }
            }
        }

        /// <summary>Refuse le tour en attente et arrête le transport lors de la libération du client.</summary>
        /// <returns>Tâche terminée après disposal.</returns>
        [TestMethod]
        public async Task DisposeRejectsPendingTurnAndStopsTransport()
        {
            var transport = new FakeTransport();
            var client = Client(transport);
            var turn = client.TurnAsync("first", null, null);
            await transport.TurnStarted.Task;
            client.Dispose();
            await Assert.ThrowsExceptionAsync<ObjectDisposedException>(() => turn);
            Assert.IsTrue(transport.Disposed);
        }

        /// <summary>Envoie l’identifiant exact du tour interrompu et termine sa tâche comme annulée.</summary>
        /// <returns>Tâche terminée après l’annulation du tour.</returns>
        [TestMethod]
        public async Task InterruptedTurnSendsExactTurnIdAndCompletesAsCanceled()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var turn = client.TurnAsync("Do work", "gpt-alpha", "medium");
                await transport.TurnStarted.Task;
                await client.InterruptAsync();
                var request = transport.Request("turn/interrupt");
                Assert.AreEqual("turn-1", FakeTransport.Object(request["params"])["turnId"]);
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => turn);
            }
        }

        /// <summary>Fait échouer le tour lorsqu’un transport se termine puis le libère proprement.</summary>
        /// <returns>Tâche terminée après l’arrêt du transport.</returns>
        [TestMethod]
        public async Task ServerExitFailsPendingTurnAndDisposalStopsTransport()
        {
            var transport = new FakeTransport();
            var client = Client(transport);
            try
            {
                var turn = client.TurnAsync("Do work", null, null);
                await transport.TurnStarted.Task;
                transport.Exit();
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => turn);
                StringAssert.Contains(error.Message, "stopped");
            }
            finally
            {
                client.Dispose();
            }

            Assert.IsTrue(transport.Disposed);
        }

        /// <summary>Répond method-not-found aux requêtes serveur non prises en charge sans fermer le fil.</summary>
        /// <returns>Tâche terminée après le traitement de la requête.</returns>
        [TestMethod]
        public async Task UnsupportedServerRequestGetsMethodNotFoundWithoutEndingConversation()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                await client.ListModelsAsync();
                transport.Emit(new { id = "foreign-request", method = "unknown/action", @params = new { } });
                var reply = transport.Sent.Last();
                Assert.AreEqual("foreign-request", reply["id"]);
                Assert.AreEqual(-32601, Convert.ToInt32(FakeTransport.Object(reply["error"])["code"]));
                Assert.AreEqual("thread-1", client.ThreadId);
            }
        }

        /// <summary>Refuse un appel d’outil étranger sans invoquer les outils VBE.</summary>
        /// <returns>Tâche terminée après le refus de l’appel.</returns>
        [TestMethod]
        public async Task ForeignToolCallIsRejectedWithoutInvokingVbeTools()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var turn = client.TurnAsync("Do work", null, null);
                await transport.TurnStarted.Task;
                transport.Emit(new { id = "tool-call", method = "item/tool/call", @params = new { threadId = "another-thread", tool = "status", arguments = new { } } });
                var reply = transport.Sent.Last();
                Assert.AreEqual("tool-call", reply["id"]);
                Assert.AreEqual(false, FakeTransport.Object(reply["result"])["success"]);
                transport.EmitTurnCompleted("completed", null);
                await turn;
            }
        }

        /// <summary>Refuse un prompt vide et un second tour lancé alors qu’un premier est actif.</summary>
        /// <returns>Tâche terminée après les vérifications de précondition.</returns>
        [TestMethod]
        public async Task EmptyPromptAndOverlappingTurnAreRejected()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => client.TurnAsync(" ", null, null));
                var pending = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.TurnAsync("second", null, null));
                transport.EmitTurnCompleted("completed", null);
                await pending;
            }
        }
    }
}
