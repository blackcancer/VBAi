using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class CodexAppServerClientTests
    {
        private static CodexAppServerClient Client(FakeTransport transport, string resumed = null)
        {
            var settings = new LlmSettings();
            return new CodexAppServerClient(new ImmediateContext(),
                new LlmVbeTools(null, null, settings), null, settings, resumed, transport);
        }

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
                CollectionAssert.AreEqual(new[] { "initialize", "initialized", "account/read", "thread/start", "model/list" },
                    transport.Methods.ToArray());
                var thread = transport.Request("thread/start");
                var arguments = FakeTransport.Object(thread["params"]);
                Assert.AreEqual("read-only", arguments["sandbox"]);
                Assert.AreEqual("untrusted", arguments["approvalPolicy"]);
                Assert.IsTrue(((object[])arguments["dynamicTools"]).Length > 0);
            }
            Assert.IsTrue(transport.Disposed);
        }

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

        [TestMethod]
        public async Task NonChatGptAccountStopsBeforeOpeningThread()
        {
            var transport = new FakeTransport { AccountType = "apiKey" };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "ChatGPT");
                Assert.IsFalse(transport.Methods.Contains("thread/start"));
                Assert.IsTrue(transport.Disposed);
            }
        }

        [TestMethod]
        public async Task MissingThreadIdentityFailsClosed()
        {
            var transport = new FakeTransport { ReturnEmptyThread = true };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "did not create a thread");
                Assert.IsTrue(transport.Disposed);
            }
        }

        [TestMethod]
        public async Task ModelCataloguePaginatesAndReadsEffortMetadata()
        {
            var transport = new FakeTransport { PaginateModels = true };
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

        [TestMethod]
        public async Task MissingCatalogueDataIsRejected()
        {
            var transport = new FakeTransport { MissingModelData = true };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.ListModelsAsync());
                Assert.IsFalse(string.IsNullOrWhiteSpace(error.Message));
            }
        }

        [TestMethod]
        public async Task RequestErrorPropagatesServerMessageWithoutKillingReadyThread()
        {
            var transport = new FakeTransport { ModelRequestError = true };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.ListModelsAsync());
                Assert.AreEqual("model unavailable", error.Message);
                Assert.AreEqual("thread-1", client.ThreadId);
                Assert.IsFalse(transport.Disposed);
            }
        }

        [TestMethod]
        public async Task TransportStartFailureDisposesClientAndDoesNotSendInitialize()
        {
            var transport = new FakeTransport { FailStart = true };
            using (var client = Client(transport))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.ListModelsAsync());
                StringAssert.Contains(error.Message, "start refused");
                Assert.AreEqual(0, transport.Sent.Count);
                Assert.IsTrue(transport.Disposed);
            }
        }

        [TestMethod]
        public async Task TurnPublishesStreamingUpdatesAndFinalText()
        {
            var transport = new FakeTransport { CompleteTurn = true };
            using (var client = Client(transport))
            {
                var updates = new List<string>();
                client.ChatUpdate += (kind, id, text, complete) =>
                    updates.Add(kind + ":" + id + ":" + text + ":" + complete);
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

        [TestMethod]
        public async Task NotificationsFromAnotherThreadDoNotReachChatUpdate()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                int updates = 0;
                client.ChatUpdate += (kind, id, text, complete) => updates++;
                await client.ListModelsAsync();
                transport.Emit(new { method = "item/agentMessage/delta", @params = new {
                    threadId = "thread-1", itemId = "ignored", delta = "before turn" } });
                var turn = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.Emit(new { method = "item/agentMessage/delta", @params = new {
                    threadId = "other", itemId = "ignored", delta = "foreign" } });
                transport.EmitTurnCompleted("completed", null);
                await turn;
                Assert.AreEqual(0, updates);
            }
        }

        [TestMethod]
        public async Task MalformedServerLineFailsPendingTurn()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var turn = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.EmitRaw("{bad json");
                try { await turn; Assert.Fail("Malformed JSON must fail the pending turn."); }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
            }
        }

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
            finally { client.Dispose(); }
            Assert.IsTrue(transport.Disposed);
        }

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

        [TestMethod]
        public async Task ForeignToolCallIsRejectedWithoutInvokingVbeTools()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                var turn = client.TurnAsync("Do work", null, null);
                await transport.TurnStarted.Task;
                transport.Emit(new { id = "tool-call", method = "item/tool/call", @params =
                    new { threadId = "another-thread", tool = "status", arguments = new { } } });
                var reply = transport.Sent.Last();
                Assert.AreEqual("tool-call", reply["id"]);
                Assert.AreEqual(false, FakeTransport.Object(reply["result"])["success"]);
                transport.EmitTurnCompleted("completed", null);
                await turn;
            }
        }

        [TestMethod]
        public async Task EmptyPromptAndOverlappingTurnAreRejected()
        {
            var transport = new FakeTransport();
            using (var client = Client(transport))
            {
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => client.TurnAsync(" ", null, null));
                var pending = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => client.TurnAsync("second", null, null));
                transport.EmitTurnCompleted("completed", null);
                await pending;
            }
        }

        private sealed class ImmediateContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback callback, object state) { callback(state); }
        }

        private sealed class FakeTransport : ICodexAppServerTransport
        {
            private readonly JavaScriptSerializer json = new JavaScriptSerializer();
            private bool running;
            private int modelPage;
            public event Action<string> LineReceived;
            public event Action<Exception> Exited;
            public bool IsRunning => running;
            public bool Disposed { get; private set; }
            public string AccountType { get; set; } = "chatgpt";
            public bool ReturnEmptyThread { get; set; }
            public bool PaginateModels { get; set; }
            public bool MissingModelData { get; set; }
            public bool ModelRequestError { get; set; }
            public bool FailStart { get; set; }
            public bool CompleteTurn { get; set; }
            public List<IDictionary<string, object>> Sent { get; } = new List<IDictionary<string, object>>();
            public IEnumerable<string> Methods => Sent.Where(x => x.ContainsKey("method"))
                .Select(x => Convert.ToString(x["method"]));
            public TaskCompletionSource<bool> TurnStarted { get; } =
                new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            public void Start()
            {
                if (FailStart) throw new InvalidOperationException("fake transport start refused");
                running = true;
            }
            public void Send(string line)
            {
                var message = Object(json.DeserializeObject(line));
                Sent.Add(message);
                object rawMethod;
                if (!message.TryGetValue("method", out rawMethod)) return;
                string method = Convert.ToString(rawMethod);
                if (!message.ContainsKey("id")) return;
                object id = message["id"];
                switch (method)
                {
                    case "initialize": Reply(id, new { }); break;
                    case "account/read": Reply(id, new { account = new { type = AccountType } }); break;
                    case "thread/start":
                        Reply(id, new { thread = new { id = ReturnEmptyThread ? "" : "thread-1" } }); break;
                    case "thread/resume":
                        Reply(id, new { thread = new { id = Object(message["params"])["threadId"] } }); break;
                    case "model/list":
                        if (ModelRequestError) { Emit(new { id, error = new { message = "model unavailable" } }); break; }
                        if (MissingModelData) { Reply(id, new { }); break; }
                        if (PaginateModels && modelPage++ == 0)
                            Reply(id, new { data = new object[] { new { model = "gpt-alpha", displayName = "Alpha",
                                isDefault = true, defaultReasoningEffort = "high",
                                supportedReasoningEfforts = new object[] {
                                    new { reasoningEffort = "medium", description = "Medium" },
                                    new { reasoningEffort = "high", description = "High" } } } },
                                nextCursor = "page-2" });
                        else Reply(id, new { data = PaginateModels ? new object[] {
                            new { model = "gpt-beta", displayName = "Beta" } } : new object[0],
                            nextCursor = (string)null });
                        break;
                    case "turn/start":
                        Emit(new { method = "turn/started", @params = new { threadId = "thread-1",
                            turn = new { id = "turn-1" } } });
                        Reply(id, new { turn = new { id = "turn-1" } });
                        TurnStarted.TrySetResult(true);
                        if (CompleteTurn)
                        {
                            Emit(new { method = "item/agentMessage/delta", @params = new {
                                threadId = "thread-1", itemId = "msg-1", delta = "partial" } });
                            Emit(new { method = "item/reasoning/summaryTextDelta", @params = new {
                                threadId = "thread-1", itemId = "reason-1", delta = "thinking" } });
                            Emit(new { method = "item/reasoning/summaryPartAdded", @params = new {
                                threadId = "thread-1", itemId = "reason-1", summaryIndex = 1 } });
                            Emit(new { method = "item/completed", @params = new { threadId = "thread-1",
                                item = new { type = "reasoning", id = "reason-1", summary =
                                    new object[] { new { text = "summary complete" } } } } });
                            Emit(new { method = "item/completed", @params = new { threadId = "thread-1",
                                item = new { type = "agentMessage", phase = "final", id = "msg-1", text = "Final answer" } } });
                            EmitTurnCompleted("completed", null);
                        }
                        break;
                    case "turn/interrupt": Reply(id, new { }); EmitTurnCompleted("interrupted", null); break;
                }
            }
            public void EmitTurnCompleted(string status, string error)
            {
                Emit(new { method = "turn/completed", @params = new { threadId = "thread-1",
                    turn = new { status, error = error == null ? null : new { message = error } } } });
            }
            public void Emit(object value) { LineReceived?.Invoke(json.Serialize(value)); }
            public void EmitRaw(string line) { LineReceived?.Invoke(line); }
            public void Exit() { running = false; Exited?.Invoke(new InvalidOperationException("Codex app-server stopped.")); }
            private void Reply(object id, object result) { Emit(new { id, result }); }
            public IDictionary<string, object> Request(string method)
            {
                return Sent.First(x => Convert.ToString(x["method"]) == method);
            }
            public static IDictionary<string, object> Object(object value) { return (IDictionary<string, object>)value; }
            public void Dispose() { running = false; Disposed = true; }
        }
    }
}
