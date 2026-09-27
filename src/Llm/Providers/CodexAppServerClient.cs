using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // Owns a single Codex CLI child process. Codex itself owns ChatGPT authentication.
    internal sealed class CodexAppServerClient : IDisposable
    {
        private readonly SynchronizationContext ui;
        private readonly LlmVbeTools tools;
        private readonly Action<string> progress;
        private readonly LlmSettings settings;
        private readonly Dictionary<int, TaskCompletionSource<IDictionary<string, object>>> requests =
            new Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>();
        private readonly object gate = new object();
        private Process process;
        private int nextId;
        private string threadId;
        private TaskCompletionSource<string> turnDone;
        private string finalText;
        private bool disposed;
        private string activeTurnId;
        private bool interruptRequested;
        public string ThreadId { get { return threadId; } }
        public event Action<string> ThreadReady;
        public event Action<string, string, string, bool> ChatUpdate;

        public CodexAppServerClient(SynchronizationContext ui, LlmVbeTools tools, Action<string> progress, LlmSettings settings, string resumeThreadId = null)
        {
            this.ui = ui ?? throw new ArgumentNullException(nameof(ui));
            this.tools = tools ?? throw new ArgumentNullException(nameof(tools));
            this.progress = progress ?? (_ => { });
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
            threadId = resumeThreadId;
        }

        public async Task<LlmModelOption[]> ListModelsAsync()
        {
            if (process == null) await StartAsync();
            var result = new List<LlmModelOption>();
            string cursor = null;
            do
            {
                var response = await RequestAsync("model/list", new { limit = 100, cursor, includeHidden = false });
                var body = GetObject(response, "result");
                object raw;
                var data = body != null && body.TryGetValue("data", out raw) ? raw as object[] : null;
                if (data == null) throw new InvalidOperationException(UiText.Get("Codex did not provide its model list."));
                foreach (var entry in data)
                {
                    var item = entry as IDictionary<string, object>;
                    string id = GetString(item, "model");
                    if (!string.IsNullOrWhiteSpace(id))
                    {
                        object rawEfforts;
                        var efforts = item.TryGetValue("supportedReasoningEfforts", out rawEfforts)
                            ? rawEfforts as object[] : null;
                        var options = new List<LlmEffortOption>();
                        if (efforts != null) foreach (var effortEntry in efforts)
                        {
                            var option = effortEntry as IDictionary<string, object>;
                            string effort = GetString(option, "reasoningEffort");
                            if (!string.IsNullOrWhiteSpace(effort))
                                options.Add(new LlmEffortOption(effort, GetString(option, "description")));
                        }
                        result.Add(new LlmModelOption(id, GetString(item, "displayName"),
                            GetString(item, "isDefault") == "True", GetString(item, "defaultReasoningEffort"), options.ToArray()));
                    }
                }
                cursor = GetString(body, "nextCursor");
            } while (!string.IsNullOrWhiteSpace(cursor));
            return result.ToArray();
        }

        public async Task<string> TurnAsync(string prompt, string model, string effort)
        {
            if (string.IsNullOrWhiteSpace(prompt)) throw new ArgumentException("A prompt is required.");
            if (turnDone != null) throw new InvalidOperationException("A Codex turn is already running.");
            finalText = null;
            interruptRequested = false;
            turnDone = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                if (process == null) await StartAsync();
                if (interruptRequested) throw new OperationCanceledException();
                var started = await RequestAsync("turn/start", new {
                    threadId,
                    model,
                    effort,
                    input = new[] { new { type = "text", text = prompt } }
                });
                activeTurnId = GetString(GetObject(GetObject(started, "result"), "turn"), "id") ?? activeTurnId;
                if (interruptRequested) await InterruptAsync();
                return await turnDone.Task;
            }
            finally { turnDone = null; activeTurnId = null; }
        }

        public async Task InterruptAsync()
        {
            interruptRequested = true;
            if (turnDone != null && activeTurnId != null)
                await RequestAsync("turn/interrupt", new { threadId, turnId = activeTurnId });
        }

        private async Task StartAsync()
        {
            string executable = Environment.GetEnvironmentVariable("CODEXVBE_CODEX_CLI");
            if (string.IsNullOrWhiteSpace(executable))
            {
                string installed = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", "OpenAI", "Codex", "bin", "codex.exe");
                executable = File.Exists(installed) ? installed : "codex.exe";
            }
            var info = new ProcessStartInfo(executable, "app-server") {
                UseShellExecute = false, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                CreateNoWindow = true, StandardOutputEncoding = new UTF8Encoding(false),
                StandardErrorEncoding = new UTF8Encoding(false)
            };
            process = new Process { StartInfo = info, EnableRaisingEvents = true };
            process.Exited += (sender, args) => FailPending(new InvalidOperationException(UiText.Get("Codex app-server stopped.")));
            try
            {
                if (!ProcessInput.StartWithoutPreamble(process)) throw new InvalidOperationException(UiText.Get("Unable to start codex app-server."));
                process.OutputDataReceived += (sender, args) => { if (args.Data != null) OnLine(args.Data); };
                // Read stderr so the child cannot block on a full pipe. Diagnostics are never treated as protocol data.
                process.ErrorDataReceived += (sender, args) => { };
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                progress(UiText.Get("Codex: initializing the local server"));
                await RequestAsync("initialize", new {
                    clientInfo = new { name = "codexvbe", title = "CodexVBE", version = "0.1.0" },
                    capabilities = new { experimentalApi = true }
                });
                Send(new { method = "initialized", @params = new { } });
                progress(UiText.Get("Codex: checking the ChatGPT account"));
                var account = await RequestAsync("account/read", new { refreshToken = false });
                var accountInfo = GetObject(GetObject(account, "result"), "account");
                if (GetString(accountInfo, "type") != "chatgpt")
                    throw new InvalidOperationException(UiText.Get("Codex must be signed in through ChatGPT. No API key is used here."));

                var definitions = LlmVbeTools.Definitions.Select(raw => {
                    dynamic function = ((dynamic)raw).function;
                    return (object)new { type = "function", name = (string)function.name,
                        description = (string)function.description, inputSchema = function.parameters };
                }).ToArray();
                progress(UiText.Get("Codex: opening the VBE conversation"));
                var started = !string.IsNullOrEmpty(threadId)
                    ? await RequestAsync("thread/resume", new { threadId, approvalPolicy = "untrusted", sandbox = "read-only" })
                    : await RequestAsync("thread/start", new {
                    ephemeral = false,
                    cwd = Path.GetTempPath(),
                    sandbox = "read-only",
                    approvalPolicy = "untrusted",
                    serviceName = "codexvbe",
                    developerInstructions = LlmVbeContext.DeveloperInstructions,
                    dynamicTools = definitions
                });
                threadId = GetString(GetObject(GetObject(started, "result"), "thread"), "id");
                if (string.IsNullOrWhiteSpace(threadId)) throw new InvalidOperationException("Codex did not create a thread.");
                ThreadReady?.Invoke(threadId);
                progress(UiText.Get("Codex connected through ChatGPT"));
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        private Task<IDictionary<string, object>> RequestAsync(string method, object parameters)
        {
            if (disposed || process == null || process.HasExited)
                throw new InvalidOperationException("Codex app-server is not running.");
            int id;
            var completion = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (gate) { id = ++nextId; requests.Add(id, completion); }
            try { Send(new { method, id, @params = parameters }); }
            catch { lock (gate) requests.Remove(id); throw; }
            return completion.Task;
        }

        private void Send(object message)
        {
            string line = NewJson().Serialize(message);
            lock (gate)
            {
                if (process == null || process.HasExited) throw new InvalidOperationException("Codex app-server is not running.");
                // .NET Framework's redirected StandardInput writer can emit a BOM; JSONL requires a raw JSON byte first.
                byte[] bytes = new UTF8Encoding(false).GetBytes(line + "\n");
                process.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
                process.StandardInput.BaseStream.Flush();
            }
        }

        private void OnLine(string line)
        {
            try
            {
                var message = NewJson().DeserializeObject(line) as IDictionary<string, object>;
                if (message == null) return;
                object rawId;
                if (message.TryGetValue("id", out rawId))
                {
                    string method = GetString(message, "method");
                    if (method == "item/tool/call") { HandleToolCall(rawId, GetObject(message, "params")); return; }
                    if (method != null)
                    {
                        Send(new { id = rawId, error = new { code = -32601, message = "Unsupported request" } });
                        return;
                    }
                    int id;
                    if (!int.TryParse(Convert.ToString(rawId), out id)) return;
                    TaskCompletionSource<IDictionary<string, object>> completion;
                    lock (gate) { if (!requests.TryGetValue(id, out completion)) return; requests.Remove(id); }
                    var error = GetObject(message, "error");
                    if (error != null) completion.TrySetException(new InvalidOperationException(GetString(error, "message") ?? "Codex request failed."));
                    else completion.TrySetResult(message);
                    return;
                }
                var parameters = GetObject(message, "params");
                if (GetString(parameters, "threadId") != threadId || turnDone == null) return;
                switch (GetString(message, "method"))
                {
                    case "turn/started":
                        activeTurnId = GetString(GetObject(parameters, "turn"), "id");
                        break;
                    case "item/agentMessage/delta":
                        PublishUpdate("message", GetString(parameters, "itemId"), GetString(parameters, "delta"), false);
                        break;
                    case "item/reasoning/summaryTextDelta":
                        PublishUpdate("summary", GetString(parameters, "itemId"), GetString(parameters, "delta"), false);
                        break;
                    case "item/reasoning/summaryPartAdded":
                        PublishUpdate("summary", GetString(parameters, "itemId"), "\n\n", false);
                        break;
                    case "item/completed":
                        var item = GetObject(parameters, "item");
                        if (GetString(item, "type") == "agentMessage")
                        {
                            bool final = GetString(item, "phase") != "commentary";
                            if (final) finalText = GetString(item, "text");
                            PublishUpdate(final ? "final" : "message", GetString(item, "id"), GetString(item, "text"), true);
                        }
                        else if (GetString(item, "type") == "reasoning")
                        {
                            object summary;
                            if (item.TryGetValue("summary", out summary) && summary is object[])
                                PublishUpdate("summary", GetString(item, "id"), string.Join("\n\n", ((object[])summary).Select(Convert.ToString)), true);
                        }
                        break;
                    case "turn/completed":
                        var turn = GetObject(parameters, "turn");
                        string status = GetString(turn, "status");
                        var completion = turnDone;
                        string answer = finalText;
                        ui.Post(_ => {
                            if (status == "completed") completion.TrySetResult(answer ?? UiText.Get("Codex finished without a text response."));
                            else if (status == "interrupted") completion.TrySetException(new OperationCanceledException());
                            else completion.TrySetException(new InvalidOperationException(
                                GetString(GetObject(turn, "error"), "message") ?? "Codex turn: " + status));
                        }, null);
                        break;
                }
            }
            catch (Exception ex) { FailPending(ex); }
        }

        private void PublishUpdate(string kind, string id, string text, bool complete)
        {
            if (string.IsNullOrEmpty(id)) return;
            ui.Post(_ => { if (!disposed) ChatUpdate?.Invoke(kind, id, text, complete); }, null);
        }

        private void HandleToolCall(object requestId, IDictionary<string, object> parameters)
        {
            ui.Post(async state => {
                try
                {
                    if (GetString(parameters, "threadId") != threadId || turnDone == null)
                        throw new InvalidOperationException("Tool call does not belong to the active VBE conversation.");
                    if (interruptRequested || disposed) throw new OperationCanceledException("Conversation interrompue.");
                    string name = GetString(parameters, "tool");
                    progress("Codex appelle " + name);
                    string activityId = "tool-" + Convert.ToString(requestId);
                    ChatUpdate?.Invoke("tool", activityId, name + " · en cours", false);
                    string arguments = NewJson().Serialize(parameters["arguments"]);
                    string output = await tools.InvokeAsync(name, arguments);
                    var response = NewJson().Deserialize<Response>(output);
                    ChatUpdate?.Invoke("tool", activityId, name + (response != null && response.Ok ? UiText.Get(" · complete") : UiText.Get(" · failed")), true);
                    Send(new { id = requestId, result = new {
                        contentItems = new[] { new { type = "inputText", text = output } },
                        success = response != null && response.Ok
                    } });
                }
                catch (Exception ex)
                {
                    try { Send(new { id = requestId, result = new {
                        contentItems = new[] { new { type = "inputText", text = ex.Message } }, success = false
                    } }); }
                    catch { FailPending(ex); }
                }
            }, null);
        }

        private void FailPending(Exception error)
        {
            lock (gate)
            {
                foreach (var pending in requests.Values) pending.TrySetException(error);
                requests.Clear();
            }
            turnDone?.TrySetException(error);
        }

        private static JavaScriptSerializer NewJson() { return new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }; }
        private static IDictionary<string, object> GetObject(IDictionary<string, object> value, string key)
        {
            object raw;
            return value != null && value.TryGetValue(key, out raw) ? raw as IDictionary<string, object> : null;
        }
        private static string GetString(IDictionary<string, object> value, string key)
        {
            object raw;
            return value != null && value.TryGetValue(key, out raw) ? Convert.ToString(raw) : null;
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            FailPending(new ObjectDisposedException(nameof(CodexAppServerClient)));
            if (process == null) return;
            try { if (!process.HasExited) process.Kill(); } catch { }
            process.Dispose();
            process = null;
        }
    }
}
