namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;

    public sealed partial class CodexAppServerClientTests
    {
        private static CodexAppServerClient Client(FakeTransport transport, string resumed = null, string instructionsHash = null)
        {
            var settings = new LlmSettings();
            return new CodexAppServerClient(new ImmediateContext(), new LlmVbeTools(null, null, settings), null,
                settings, resumed, instructionsHash, transport);
        }
        private static string Method(IDictionary<string, object> message) { return message.ContainsKey("method") ? Convert.ToString(message["method"]) : null; }

        private sealed class ImmediateContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback callback, object state)
            {
                callback(state);
            }
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
            public Func<IDictionary<string, object>, bool> Intercept { get; set; }
            public Action BeforeSend { get; set; }
            public List<IDictionary<string, object>> Sent { get; } = new List<IDictionary<string, object>>();
            public IEnumerable<string> Methods => Sent.Where(x => x.ContainsKey("method")).Select(x => Convert.ToString(x["method"]));
            public TaskCompletionSource<bool> TurnStarted { get; } = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            public void Start()
            {
                if (FailStart)
                    throw new InvalidOperationException("fake transport start refused");
                running = true;
            }

            public void Send(string line)
            {
                BeforeSend?.Invoke();
                var message = Object(json.DeserializeObject(line));
                Sent.Add(message);
                if (Intercept != null && Intercept(message)) return;
                object rawMethod;
                if (!message.TryGetValue("method", out rawMethod))
                    return;
                string method = Convert.ToString(rawMethod);
                if (!message.ContainsKey("id"))
                    return;
                object id = message["id"];
                switch (method)
                {
                    case "initialize":
                        Reply(id, new { });
                        break;
                    case "account/read":
                        Reply(id, new { account = new { type = AccountType } });
                        break;
                    case "thread/start":
                        Reply(id, new { thread = new { id = ReturnEmptyThread ? "" : "thread-1" } });
                        break;
                    case "thread/resume":
                        Reply(id, new { thread = new { id = Object(message["params"])["threadId"] } });
                        break;
                    case "model/list":
                        if (ModelRequestError)
                        {
                            Emit(new { id, error = new { message = "model unavailable" } });
                            break;
                        }

                        if (MissingModelData)
                        {
                            Reply(id, new { });
                            break;
                        }

                        if (PaginateModels && modelPage++ == 0)
                            Reply(id, new { data = new object[] { new { model = "gpt-alpha", displayName = "Alpha", isDefault = true, defaultReasoningEffort = "high", supportedReasoningEfforts = new object[] { new { reasoningEffort = "medium", description = "Medium" }, new { reasoningEffort = "high", description = "High" } } } }, nextCursor = "page-2" });
                        else
                            Reply(id, new { data = PaginateModels ? new object[] { new { model = "gpt-beta", displayName = "Beta" } } : new object[0], nextCursor = (string)null });
                        break;
                    case "turn/start":
                        Emit(new { method = "turn/started", @params = new { threadId = "thread-1", turn = new { id = "turn-1" } } });
                        Reply(id, new { turn = new { id = "turn-1" } });
                        TurnStarted.TrySetResult(true);
                        if (CompleteTurn)
                        {
                            Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread-1", itemId = "msg-1", delta = "partial" } });
                            Emit(new { method = "item/reasoning/summaryTextDelta", @params = new { threadId = "thread-1", itemId = "reason-1", delta = "thinking" } });
                            Emit(new { method = "item/reasoning/summaryPartAdded", @params = new { threadId = "thread-1", itemId = "reason-1", summaryIndex = 1 } });
                            Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "reasoning", id = "reason-1", summary = new object[] { new { text = "summary complete" } } } } });
                            Emit(new { method = "item/completed", @params = new { threadId = "thread-1", item = new { type = "agentMessage", phase = "final", id = "msg-1", text = "Final answer" } } });
                            EmitTurnCompleted("completed", null);
                        }

                        break;
                    case "turn/interrupt":
                        Reply(id, new { });
                        EmitTurnCompleted("interrupted", null);
                        break;
                }
            }

            public void EmitTurnCompleted(string status, string error)
            {
                Emit(new { method = "turn/completed", @params = new { threadId = "thread-1", turn = new { status, error = error == null ? null : new { message = error } } } });
            }

            public void Emit(object value)
            {
                LineReceived?.Invoke(json.Serialize(value));
            }

            public void EmitRaw(string line)
            {
                LineReceived?.Invoke(line);
            }

            public void Exit()
            {
                running = false;
                Exited?.Invoke(new InvalidOperationException("Codex app-server stopped."));
            }
            public void Stop() { running = false; }

            private void Reply(object id, object result)
            {
                Emit(new { id, result });
            }

            public IDictionary<string, object> Request(string method)
            {
                return Sent.First(x => Convert.ToString(x["method"]) == method);
            }

            public static IDictionary<string, object> Object(object value)
            {
                return (IDictionary<string, object>)value;
            }

            public void Dispose()
            {
                running = false;
                Disposed = true;
            }
        }
    }
}
