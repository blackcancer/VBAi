using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;

internal static partial class ProviderTests
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    private static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
    private static IDictionary<string, object> Obj(object x) { return ClaudeProtocol.Object(x); }
    private static string Text(IDictionary<string, object> x, string key) { return ClaudeProtocol.Text(x, key); }
    private static LlmProvider Provider(string name) { return LlmProvider.All.Single(x => x.Name == name); }
    private static readonly object[] Tools = { new { type = "function", function = new { name = "read_module", description = "Read VBA", parameters = new { type = "object", properties = new { } } } } };
    private static List<object> History() { return new List<object> { new { role = "system", content = "Read before editing." }, new { role = "user", content = "Explique la procédure été." } }; }
    private sealed class Handler : HttpMessageHandler
    {
        internal Func<HttpRequestMessage, Task<HttpResponseMessage>> Handle;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation) { return Handle(request); }
    }
    private static HttpResponseMessage Response(object payload, HttpStatusCode code = HttpStatusCode.OK)
    { return new HttpResponseMessage(code) { Content = new StringContent(Json.Serialize(payload), Encoding.UTF8, "application/json") }; }

    [STAThread]
    public static int Main(string[] args)
    {
        try {
            if (args.Contains("--headless")) { FakeCopilot(); return 0; }
            if (args.Contains("--live-openrouter")) { LiveOpenRouter().GetAwaiter().GetResult(); return 0; }
            if (args.Contains("--live-openrouter-tools")) { LiveOpenRouterTools().GetAwaiter().GetResult(); return 0; }
            SettingsUi(); Run().GetAwaiter().GetResult(); Extended().GetAwaiter().GetResult(); return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static void SettingsUi()
    {
        var settings = new LlmSettings { ProviderName = "Claude" };
        using (var form = new LlmSettingsWindow(settings)) {
            Func<string, object> field = name => typeof(LlmSettingsWindow).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
            var picker = (System.Windows.Forms.ComboBox)field("provider");
            var endpoint = (System.Windows.Forms.TextBox)field("openAiEndpoint");
            var key = (System.Windows.Forms.TextBox)field("openAiKey");
            endpoint.Text = "https://claude.example/v1/messages"; key.Text = "draft-only";
            picker.SelectedItem = Provider("Gemini");
            Assert(endpoint.Text == Provider("Gemini").Endpoint && key.Text == "", "provider draft leaked into another provider");
            picker.SelectedItem = Provider("Claude");
            Assert(endpoint.Text == "https://claude.example/v1/messages" && key.Text == "draft-only", "draft lost on provider switch");
            picker.SelectedItem = Provider("LM Studio");
            Assert(((System.Windows.Forms.TextBox)field("ollamaEndpoint")).Text == Provider("LM Studio").Endpoint, "local endpoint missing");
            picker.SelectedItem = Provider("Azure OpenAI");
            var models = (System.Windows.Forms.TextBox)field("manualModels");
            models.Text = "my-deployment";
            picker.SelectedItem = Provider("Amazon Bedrock");
            Assert(models.Text == "", "model draft leaked");
            picker.SelectedItem = Provider("Azure OpenAI");
            Assert(models.Text == "my-deployment", "deployment draft lost");
        }
        Assert(settings.ProviderEndpoints.Count == 0 && settings.EncryptedProviderKeys.Count == 0, "cancel changed persisted settings");
        // WinForms can install a synchronization context when creating controls. The protocol tests have no UI loop.
        SynchronizationContext.SetSynchronizationContext(null);
        Console.WriteLine("PASS settings form provider switching, isolated drafts, local endpoint and cancellation");
    }
    private static async Task Run()
    {
        var settings = new LlmSettings();
        foreach (var p in LlmProvider.All.Where(x => !x.IsCodex && !x.IsCopilot)) {
            settings.SetKey(p, "test-key-" + p.Name);
            settings.SetSelectedModel(p, "model-" + p.Name);
        }
        var restored = Json.Deserialize<LlmSettings>(Json.Serialize(settings));
        foreach (var p in LlmProvider.All.Where(x => !x.IsCodex && !x.IsCopilot)) {
            Assert(restored.GetKey(p) == "test-key-" + p.Name, "key isolation " + p.Name);
            Assert(restored.GetSelectedModel(p) == "model-" + p.Name, "model isolation " + p.Name);
        }
        Assert(!Json.Serialize(settings).Contains("test-key"), "plain key persisted");
        var legacy = Json.Deserialize<LlmSettings>("{\"OpenAiModel\":\"old-openai\",\"OllamaModel\":\"old-local\",\"ProviderModels\":null}");
        Assert(legacy.GetSelectedModel(Provider("OpenAI API")) == "old-openai" && legacy.GetSelectedModel(Provider("Ollama")) == "old-local", "legacy migration");
        legacy.SetSelectedModel(Provider("Gemini"), "gemini-test");
        Assert(legacy.GetSelectedModel(Provider("Gemini")) == "gemini-test", "null dictionary migration");
        Console.WriteLine("PASS provider key/model isolation, DPAPI roundtrip and legacy settings");

        foreach (var p in LlmProvider.All.Where(x => !x.IsCodex && !x.IsCopilot && !x.IsClaude && !x.ManualModels)) {
            var handler = new Handler { Handle = async request => {
                Assert(request.RequestUri.AbsoluteUri == p.Endpoint, "endpoint " + p.Name);
                Assert(request.Headers.Authorization.Parameter == "test-key-" + p.Name, "auth " + p.Name);
                var payload = Obj(Json.DeserializeObject(await request.Content.ReadAsStringAsync()));
                Assert(payload.ContainsKey("store") == (p.Name == "OpenAI API"), "provider-specific store " + p.Name);
                Assert(ClaudeProtocol.Array(payload, "tools").Length == 1, "missing tools");
                return Response(new { choices = new[] { new { message = new { role = "assistant", content = "été", reasoning_content = "opaque", tool_calls = new[] {
                    new { id = "call-1", type = "function", function = new { name = "read_module", arguments = "{}" }, extra_content = new { google = new { thought_signature = "opaque-signature" } } }
                } } } } });
            } };
            using (var client = new LlmChatClient(p, settings, "model", handler)) {
                var result = await client.CompleteAsync(History(), Tools);
                Assert(Text(result, "content") == "été", "UTF8 response");
                Assert(Json.Serialize(result).Contains("opaque-signature") && result.ContainsKey("reasoning_content"), "provider continuation metadata lost");
            }
            var catalogue = new Handler { Handle = request => {
                Assert(request.RequestUri.AbsolutePath.EndsWith(p.IsOllama ? "/api/tags" : "/models"), "catalogue path " + p.Name);
                return Task.FromResult(p.IsOllama ? Response(new { models = new[] { new { name = "local-model" } } }) : Response(new { data = new[] { new { id = "remote-model" } } }));
            } };
            Assert((await LlmChatClient.ListModelsAsync(p, settings, catalogue)).Length == 1, "catalogue " + p.Name);
        }
        Console.WriteLine("PASS all compatible providers: endpoints, auth, models, tool calls and continuation metadata");

        var claude = Provider("Claude"); int claudeCalls = 0;
        var claudeHandler = new Handler { Handle = async request => {
            Assert(request.Headers.Contains("x-api-key") && request.Headers.Contains("anthropic-version") && request.Headers.Authorization == null, "Claude auth");
            var payload = Obj(Json.DeserializeObject(await request.Content.ReadAsStringAsync()));
            Assert(Text(payload, "system") == "Read before editing.", "Claude system prompt");
            Assert(Obj(ClaudeProtocol.Array(payload, "tools")[0]).ContainsKey("input_schema"), "Claude tools schema");
            if (++claudeCalls == 1) return Response(new { content = new[] { new { type = "tool_use", id = "a", name = "read_module", input = new { } } }, stop_reason = "tool_use" });
            var last = Obj(ClaudeProtocol.Array(payload, "messages").Last());
            Assert(Text(last, "role") == "user" && Text(Obj(ClaudeProtocol.Array(last, "content")[0]), "tool_use_id") == "a", "Claude tool result mapping");
            return Response(new { content = new[] { new { type = "text", text = "Réponse complète" } }, stop_reason = "end_turn" });
        } };
        using (var client = new LlmChatClient(claude, settings, "claude-model", claudeHandler)) {
            var history = History(); history.Add(await client.CompleteAsync(history, Tools));
            history.Add(new { role = "tool", tool_call_id = "a", content = "Sub Exemple()" });
            Assert(Text(await client.CompleteAsync(history, Tools), "content") == "Réponse complète", "Claude final response");
        }
        int pages = 0;
        var paginated = new Handler { Handle = request => {
            pages++; if (pages == 2) Assert(request.RequestUri.Query == "?after_id=first", "Claude cursor");
            return Task.FromResult(Response(new { data = new[] { new { id = pages == 1 ? "first" : "second", display_name = "Claude" } }, has_more = pages == 1, last_id = "first" }));
        } };
        Assert((await LlmChatClient.ListModelsAsync(claude, settings, paginated)).Length == 2, "Claude pagination");
        Console.WriteLine("PASS native Claude system/tools/results, UTF8 and catalogue pagination");

        var failure = new Handler { Handle = r => Task.FromResult(Response(new { secret = "do-not-display" }, HttpStatusCode.Unauthorized)) };
        using (var client = new LlmChatClient(Provider("Gemini"), settings, "model", failure)) {
            try { await client.CompleteAsync(History(), Tools); throw new Exception("expected failure"); }
            catch (InvalidOperationException ex) { Assert(ex.Message.Contains("401") && !ex.Message.Contains("do-not-display"), "unsafe HTTP error"); }
        }
        settings.SetEndpoint(Provider("Gemini"), "http://example.org/chat/completions");
        try { using (new LlmChatClient(Provider("Gemini"), settings, "model")) { } throw new Exception("HTTP allowed"); }
        catch (InvalidOperationException) { }
        Console.WriteLine("PASS HTTP failures redact response bodies and remote HTTP is rejected");

        string previous = Environment.GetEnvironmentVariable("CODEXVBE_COPILOT_CLI");
        try {
            Environment.SetEnvironmentVariable("CODEXVBE_COPILOT_CLI", Assembly.GetExecutingAssembly().Location);
            foreach (string version in new[] { "2", "3" }) {
                Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_VERSION", version);
                using (var client = new CopilotClient()) Assert((await client.ListModelsAsync()).Single().Id == "test-model", "Copilot models");
                using (var client = new CopilotClient()) {
                    int toolsCalled = 0; var history = History(); var streamed = new StringBuilder();
                    client.TextDelta = text => streamed.Append(text);
                    var result = await client.CompleteAsync("test-model", history, Tools, (name, args) => {
                        Assert(name == "read_module", "unexpected native tool execution"); toolsCalled++;
                        return Task.FromResult("résultat été");
                    });
                    Assert(Text(result, "content") == "Réponse été" && toolsCalled == 1, "Copilot response or duplicate tool execution");
                    Assert(history.Count == 4, "Copilot tool history not retained");
                    Assert(streamed.ToString() == "Réponse été", "Copilot text streaming missing");
                }
            }
            Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_VERSION", "99");
            using (var client = new CopilotClient()) {
                try { await client.ListModelsAsync(); throw new Exception("unsupported protocol accepted"); }
                catch (InvalidOperationException ex) { Assert(ex.Message.Contains("incompatible"), "protocol error"); }
            }
            Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_VERSION", "3");
            Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_WAIT", "1");
            using (var client = new CopilotClient()) {
                var task = client.CompleteAsync("test-model", History(), Tools, (name, args) => Task.FromResult("unused"));
                await Task.Delay(250); client.Dispose();
                Assert(await Task.WhenAny(task, Task.Delay(3000)) == task, "Copilot cancellation hung");
                try { await task; throw new Exception("cancelled turn completed"); } catch (OperationCanceledException) { }
            }
        } finally { Environment.SetEnvironmentVariable("CODEXVBE_COPILOT_CLI", previous); Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_VERSION", null); Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_WAIT", null); }
        Console.WriteLine("PASS Copilot framed subprocess protocol v2/v3, tool dispatch/deduplication, permission scoping, foreign sessions and unsupported version");
    }

    private static void WriteFrame(object message)
    {
        byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(message));
        var stream = Console.OpenStandardOutput(); byte[] header = Encoding.ASCII.GetBytes("Content-Length: " + data.Length + "\r\n\r\n");
        stream.Write(header, 0, header.Length); stream.Write(data, 0, data.Length); stream.Flush();
    }
    private static IDictionary<string, object> ReadFrame(Stream stream)
    {
        var header = new StringBuilder();
        while (!header.ToString().EndsWith("\r\n\r\n")) { int b = stream.ReadByte(); if (b < 0) return null; header.Append((char)b); }
        int length = int.Parse(header.ToString().Substring(15).Trim()); var data = new byte[length]; int offset = 0;
        while (offset < length) { int n = stream.Read(data, offset, length - offset); if (n == 0) throw new EndOfStreamException(); offset += n; }
        return Obj(Json.DeserializeObject(Encoding.UTF8.GetString(data)));
    }
    private static void FakeCopilot()
    {
        int version = int.Parse(Environment.GetEnvironmentVariable("CODEXVBE_TEST_COPILOT_VERSION") ?? "3");
        string session = null; var input = Console.OpenStandardInput();
        Action<string, object> evt = (type, data) => WriteFrame(new { jsonrpc = "2.0", method = "session.event", @params = new { sessionId = session, @event = new { type, data } } });
        Action finish = () => { evt("assistant.message_delta", new { deltaContent = "Réponse été" }); evt("assistant.message", new { content = "Réponse été" }); evt("session.idle", new { }); };
        while (true) {
            var message = ReadFrame(input); if (message == null) return;
            string method = Text(message, "method"); object id = message["id"];
            var p = message.ContainsKey("params") ? Obj(message["params"]) : null;
            Action<object> reply = result => WriteFrame(new { jsonrpc = "2.0", id, result });
            if (method == "ping") reply(new { protocolVersion = version });
            else if (method == "models.list") reply(new { models = new[] { new { id = "test-model", name = "Test" } } });
            else if (method == "session.create") {
                session = Text(p, "sessionId");
                Assert(ClaudeProtocol.Array(p, "availableTools").Length == 1 && Equals(p["requestPermission"], true), "Copilot tool scope");
                Assert(Text(Obj(p["systemMessage"]), "content").Contains("Read before editing."), "system instructions lost");
                reply(new { sessionId = session });
            }
            else if (method == "session.send") {
                reply(new { messageId = "m1" });
                if (Environment.GetEnvironmentVariable("CODEXVBE_TEST_COPILOT_WAIT") == "1") continue;
                WriteFrame(new { jsonrpc = "2.0", method = "session.event", @params = new { sessionId = "foreign", @event = new { type = "session.idle", data = new { } } } });
                if (version == 3) evt("permission.requested", new { requestId = "deny", permissionRequest = new { kind = "shell", toolName = "read_module" } });
                else WriteFrame(new { jsonrpc = "2.0", id = "deny", method = "permission.request", @params = new { sessionId = session, permissionRequest = new { kind = "shell" } } });
            }
            else if (method == "session.permissions.handlePendingPermissionRequest" || method == null && (Convert.ToString(id) == "deny" || Convert.ToString(id) == "allow")) {
                var decision = p == null ? Obj(Obj(message["result"])["result"]) : Obj(p["result"]);
                string permissionId = p == null ? Convert.ToString(id) : Text(p, "requestId");
                Assert(Text(decision, "kind") == (permissionId == "allow" ? "approved" : "denied-by-rules"), "permission scope incorrect");
                if (p != null) reply(new { });
                if (permissionId == "deny") {
                    if (version == 3) evt("permission.requested", new { requestId = "allow", permissionRequest = new { kind = "custom-tool", toolName = "read_module" } });
                    else WriteFrame(new { jsonrpc = "2.0", id = "allow", method = "permission.request", @params = new { sessionId = session, permissionRequest = new { kind = "custom-tool", toolName = "read_module" } } });
                    continue;
                }
                if (version == 3) {
                    evt("external_tool.requested", new { requestId = "tool1", toolCallId = "c1", toolName = "read_module", arguments = new { } });
                    evt("external_tool.requested", new { requestId = "tool1", toolCallId = "c1", toolName = "read_module", arguments = new { } });
                } else WriteFrame(new { jsonrpc = "2.0", id = "tool1", method = "tool.call", @params = new { sessionId = session, toolName = "read_module", toolCallId = "c1", arguments = new { } } });
            }
            else if (method == "session.tools.handlePendingToolCall" || method == null && Convert.ToString(id) == "tool1") {
                var result = p == null ? Obj(Obj(message["result"])["result"]) : Obj(p["result"]);
                Assert(Text(result, "textResultForLlm") == "résultat été", "tool result encoding");
                if (p != null) reply(new { }); finish();
            }
            else throw new Exception("Unexpected RPC: " + method);
        }
    }
}
