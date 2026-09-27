using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // Copilot CLI's SDK stdio protocol (Content-Length framed JSON-RPC, versions 2/3).
    // Only CodexVBE tools are exposed; native permissions are denied and VBA policy remains in InvokeAsync.
    internal sealed class CopilotClient : IDisposable
    {
        private readonly object gate = new object();
        private readonly Dictionary<int, TaskCompletionSource<IDictionary<string, object>>> pending = new Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>();
        private readonly HashSet<string> handledTools = new HashSet<string>();
        private readonly SynchronizationContext ui = SynchronizationContext.Current;
        private Process process;
        private int nextId;
        private bool disposed;
        private string sessionId, answer;
        private TaskCompletionSource<string> completion;
        private Func<string, string, Task<string>> invoke;
        private HashSet<string> allowedTools;
        private IList<object> history;
        public Action<string> TextDelta { get; set; }
        private static JavaScriptSerializer Json() { return new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }; }
        private static string Text(IDictionary<string, object> o, string k) { return ClaudeProtocol.Text(o, k); }
        private static IDictionary<string, object> Object(IDictionary<string, object> o, string k)
        { object value; return o != null && o.TryGetValue(k, out value) ? value as IDictionary<string, object> : null; }
        internal static string Executable { get { return Environment.GetEnvironmentVariable("CODEXVBE_COPILOT_CLI") ?? "copilot.exe"; } }

        public static void StartLogin()
        { Process.Start(new ProcessStartInfo(Executable, "login") { UseShellExecute = true }); }

        public static async Task<string> ReadStatusAsync()
        {
            using (var client = new CopilotClient()) {
                var models = await client.ListModelsAsync();
                return "Copilot accessible · " + models.Length + " modèles. Authentification gérée par le CLI GitHub.";
            }
        }

        private async Task StartAsync()
        {
            if (disposed) throw new ObjectDisposedException(nameof(CopilotClient));
            if (process != null) return;
            var info = new ProcessStartInfo(Executable, "--headless --stdio --no-auto-update --log-level error") {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                WorkingDirectory = Path.GetTempPath()
            };
            process = new Process { StartInfo = info };
            try {
                if (!process.Start()) throw new InvalidOperationException("Impossible de démarrer Copilot.");
            } catch (Exception ex) { process.Dispose(); process = null; throw new InvalidOperationException("Installez GitHub Copilot CLI et configurez CODEXVBE_COPILOT_CLI vers son exécutable si nécessaire.", ex); }
            process.ErrorDataReceived += (s, e) => { }; // Drain without logging tokens or prompts.
            process.BeginErrorReadLine();
            _ = Task.Run(ReadLoop);
            var ping = await RequestAsync("ping", new { });
            int version;
            if (!int.TryParse(Text(ping, "protocolVersion"), out version) || version < 2 || version > 3)
                throw new InvalidOperationException("Version du protocole Copilot incompatible (versions 2 et 3 prises en charge).");
        }

        public async Task<LlmModelOption[]> ListModelsAsync()
        {
            await StartAsync();
            var result = await RequestAsync("models.list", new { });
            return ClaudeProtocol.Array(result, "models").Select(x => ClaudeProtocol.Object(x))
                .Where(x => !string.IsNullOrWhiteSpace(Text(x, "id")))
                .Select(x => new LlmModelOption(Text(x, "id"), Text(x, "name") ?? Text(x, "id"))).ToArray();
        }

        public async Task<IDictionary<string, object>> CompleteAsync(string model, IList<object> history, object[] tools,
            Func<string, string, Task<string>> toolHandler)
        {
            invoke = toolHandler ?? throw new InvalidOperationException("Les outils VBA de Copilot ne sont pas connectés.");
            this.history = history;
            await StartAsync();
            var definitions = tools.Select(x => ClaudeProtocol.Object(ClaudeProtocol.Object(x)["function"])).ToArray();
            allowedTools = new HashSet<string>(definitions.Select(x => Text(x, "name")), StringComparer.Ordinal);
            sessionId = Guid.NewGuid().ToString();
            completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var created = await RequestAsync("session.create", new {
                sessionId, model, clientName = "CodexVBE", tools = definitions, availableTools = allowedTools.ToArray(),
                requestPermission = true, streaming = TextDelta != null,
                systemMessage = new { mode = "replace", content = "You are a VBA assistant embedded in the VBE. Use only the supplied VBA tools. The prompt contains JSON conversation history; answer its last user request. Re-read live VBA before any change.\n" +
                    string.Join("\n", history.Select(ClaudeProtocol.Object).Where(x => Text(x, "role") == "system").Select(x => Text(x, "content"))) },
                enableFileHooks = false, enableSkills = false, enableHostGitOperations = false,
                enableOnDemandInstructionDiscovery = false, customAgents = new object[0],
                mcpServers = new Dictionary<string, object>()
            });
            if (Text(created, "sessionId") != sessionId) throw new InvalidOperationException("Session Copilot inattendue.");
            await RequestAsync("session.send", new { sessionId, prompt = Json().Serialize(history) });
            if (await Task.WhenAny(completion.Task, Task.Delay(TimeSpan.FromMinutes(5))) != completion.Task) {
                Dispose(); throw new TimeoutException("Copilot n’a pas terminé sa réponse dans le délai imparti.");
            }
            return new Dictionary<string, object> { ["role"] = "assistant", ["content"] = await completion.Task };
        }

        private async Task<IDictionary<string, object>> RequestAsync(string method, object parameters)
        {
            var source = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
            int id;
            lock (gate) { if (disposed) throw new ObjectDisposedException(nameof(CopilotClient)); id = ++nextId; pending.Add(id, source); }
            try {
                Send(new { jsonrpc = "2.0", id, method, @params = parameters });
                if (await Task.WhenAny(source.Task, Task.Delay(TimeSpan.FromSeconds(45))) != source.Task)
                    throw new TimeoutException("Copilot ne répond pas à " + method + ". Vérifiez la connexion du CLI.");
                return await source.Task;
            } finally { lock (gate) pending.Remove(id); }
        }

        private void Send(object message)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(Json().Serialize(message));
            byte[] header = Encoding.ASCII.GetBytes("Content-Length: " + bytes.Length + "\r\n\r\n");
            lock (gate) {
                if (disposed) throw new ObjectDisposedException(nameof(CopilotClient));
                var stream = process.StandardInput.BaseStream;
                stream.Write(header, 0, header.Length); stream.Write(bytes, 0, bytes.Length); stream.Flush();
            }
        }

        private void ReadLoop()
        {
            try {
                var stream = process.StandardOutput.BaseStream;
                while (!disposed) {
                    var header = new StringBuilder();
                    while (!header.ToString().EndsWith("\r\n\r\n", StringComparison.Ordinal)) {
                        int next = stream.ReadByte(); if (next < 0) throw new EndOfStreamException("Copilot s’est arrêté.");
                        header.Append((char)next); if (header.Length > 4096) throw new InvalidDataException("En-tête Copilot invalide.");
                    }
                    int size = 0;
                    foreach (string line in header.ToString().Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                        if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) int.TryParse(line.Substring(15).Trim(), out size);
                    if (size <= 0 || size > 10 * 1024 * 1024) throw new InvalidDataException("Taille du message Copilot invalide.");
                    var bytes = new byte[size]; int offset = 0;
                    while (offset < size) { int count = stream.Read(bytes, offset, size - offset); if (count == 0) throw new EndOfStreamException(); offset += count; }
                    Dispatch(Json().DeserializeObject(Encoding.UTF8.GetString(bytes)) as IDictionary<string, object>);
                }
            } catch (Exception ex) { Fail(ex); }
        }

        private void Dispatch(IDictionary<string, object> message)
        {
            string method = Text(message, "method");
            if (method == null) {
                int id; TaskCompletionSource<IDictionary<string, object>> source;
                if (int.TryParse(Text(message, "id"), out id)) lock (gate) if (pending.TryGetValue(id, out source)) {
                    if (message.ContainsKey("error")) source.TrySetException(new InvalidOperationException("Copilot a refusé la requête. Vérifiez la connexion, l’abonnement et le modèle dans le CLI."));
                    else source.TrySetResult(Object(message, "result") ?? new Dictionary<string, object>());
                }
                return;
            }
            var parameters = Object(message, "params");
            if (method == "session.event" && Text(parameters, "sessionId") == sessionId && completion != null && !completion.Task.IsCompleted) {
                var evt = Object(parameters, "event"); var data = Object(evt, "data"); string type = Text(evt, "type");
                if (type == "assistant.message") answer = Text(data, "content");
                else if (type == "assistant.message_delta") {
                    string fragment = Text(data, "deltaContent");
                    if (!string.IsNullOrEmpty(fragment)) {
                        if (ui != null) ui.Post(_ => { if (!disposed) TextDelta?.Invoke(fragment); }, null);
                        else TextDelta?.Invoke(fragment);
                    }
                }
                else if (type == "session.idle") completion.TrySetResult(answer ?? "");
                else if (type == "session.error") completion.TrySetException(new InvalidOperationException("La session Copilot a échoué. Vérifiez son état dans le CLI."));
                else if (type == "external_tool.requested") RunTool(data, null);
                else if (type == "permission.requested") _ = DenyAsync(data);
                return;
            }
            object requestId;
            if (message.TryGetValue("id", out requestId)) {
                if (method == "tool.call" && Text(parameters, "sessionId") == sessionId) RunTool(parameters, requestId);
                else if (method == "permission.request") Send(new { jsonrpc = "2.0", id = requestId, result = new { result = new { kind = Text(parameters, "sessionId") == sessionId ? PermissionDecision(Object(parameters, "permissionRequest")) : "denied-by-rules" } } });
                else Send(new { jsonrpc = "2.0", id = requestId, error = new { code = -32601, message = "Unsupported client method" } });
            }
        }

        private async Task DenyAsync(IDictionary<string, object> data)
        {
            try { await RequestAsync("session.permissions.handlePendingPermissionRequest", new { sessionId, requestId = Text(data, "requestId"), result = new { kind = PermissionDecision(Object(data, "permissionRequest")) } }); }
            catch (Exception ex) { Fail(ex); }
        }

        private void RunTool(IDictionary<string, object> data, object legacyId)
        {
            string requestId = Text(data, "requestId") ?? Convert.ToString(legacyId);
            lock (gate) if (!handledTools.Add(requestId)) return;
            Func<Task> run = async () => {
                try {
                    if (disposed || completion == null || completion.Task.IsCompleted) return;
                    string name = Text(data, "toolName");
                    string arguments = Json().Serialize(data.ContainsKey("arguments") ? data["arguments"] : new { });
                    string output = allowedTools != null && allowedTools.Contains(name) && invoke != null
                        ? await invoke(name, arguments)
                        : "Tool not available.";
                    string callId = Text(data, "toolCallId") ?? requestId;
                    history.Add(new { role = "assistant", content = (string)null, tool_calls = new[] { new { id = callId, type = "function", function = new { name, arguments } } } });
                    history.Add(new { role = "tool", tool_call_id = callId, content = output });
                    var result = new { textResultForLlm = output, resultType = "success" };
                    if (legacyId != null) Send(new { jsonrpc = "2.0", id = legacyId, result = new { result } });
                    else await RequestAsync("session.tools.handlePendingToolCall", new { sessionId, requestId, result });
                } catch (Exception ex) { Fail(ex); }
            };
            if (ui != null) ui.Post(async _ => await run(), null); else _ = run();
        }

        private void Fail(Exception error)
        {
            lock (gate) foreach (var request in pending.Values) request.TrySetException(error);
            completion?.TrySetException(error);
        }
        private string PermissionDecision(IDictionary<string, object> permission)
        {
            // This only routes our registered custom tool to LlmVbeTools.InvokeAsync.
            // That entry point still enforces mode/project, ReadOnly, AskEachTime and live SHA.
            // Shell, file, network, MCP and unknown tools never reach that entry point.
            return !disposed && Text(permission, "kind") == "custom-tool" && allowedTools != null &&
                allowedTools.Contains(Text(permission, "toolName")) ? "approved" : "denied-by-rules";
        }
        public void Dispose()
        {
            lock (gate) { if (disposed) return; disposed = true; }
            Fail(new OperationCanceledException());
            if (process != null) { try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { } finally { process.Dispose(); } }
        }
    }
}
