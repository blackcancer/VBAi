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

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class CodexAppServerClientTests
    {
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
                transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread-1", itemId = "ignored", delta = "before turn" } });
                var turn = client.TurnAsync("first", null, null);
                await transport.TurnStarted.Task;
                transport.Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "other", itemId = "ignored", delta = "foreign" } });
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
            finally
            {
                client.Dispose();
            }

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
                transport.Emit(new { id = "tool-call", method = "item/tool/call", @params = new { threadId = "another-thread", tool = "status", arguments = new { } } });
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
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.TurnAsync("second", null, null));
                transport.EmitTurnCompleted("completed", null);
                await pending;
            }
        }
    }
}
