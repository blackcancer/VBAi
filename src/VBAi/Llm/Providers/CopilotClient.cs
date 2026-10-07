using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{
    // Copilot CLI's SDK stdio protocol (Content-Length framed JSON-RPC, versions 2/3).
    // Only VBAi tools are exposed; native permissions are denied and VBA policy remains in InvokeAsync.
    /// <summary>Pilote le client GitHub Copilot par son protocole stdio encadré par Content-Length.</summary>
    internal sealed class CopilotClient : IDisposable
    {

        /// <summary>Verrou qui protège le processus et les requêtes en attente.</summary>
        private readonly object gate = new object();

        /// <summary>Requêtes JSON-RPC en attente, indexées par identifiant.</summary>
        private readonly Dictionary<int, TaskCompletionSource<IDictionary<string, object>>> pending = new Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>();

        /// <summary>Associates turn-bound permission/tool RPCs with their captured session and completion.</summary>
        private readonly Dictionary<int, Tuple<string, TaskCompletionSource<string>>> pendingOwners = new Dictionary<int, Tuple<string, TaskCompletionSource<string>>>();

        /// <summary>Reserves failed activity receipts before fault continuations can dispose their turn.</summary>
        private readonly Dictionary<TaskCompletionSource<string>, CodexAgentActivity[]> failedTurnActivities = new Dictionary<TaskCompletionSource<string>, CodexAgentActivity[]>();

        /// <summary>Identifiants des appels d’outils déjà traités afin d’éviter les exécutions répétées.</summary>
        private readonly HashSet<string> handledTools = new HashSet<string>();

        /// <summary>Contexte d’interface capturé à la construction pour publier les fragments et outils.</summary>
        private readonly SynchronizationContext ui = SynchronizationContext.Current;

        /// <summary>Processus Copilot CLI associé à la session.</summary>
        private Process process;

        /// <summary>Dernier identifiant local de requête JSON-RPC.</summary>
        private int nextId;

        /// <summary>Indique si le client a été libéré.</summary>
        private bool disposed;

        /// <summary>Identifiant de session Copilot et dernier texte assistant complet.</summary>
        private string sessionId, answer;

        /// <summary>Achèvement de la réponse assistant en cours.</summary>
        private TaskCompletionSource<string> completion;

        /// <summary>Fonction qui exécute un outil VBA autorisé.</summary>
        private Func<string, string, Task<string>> invoke;

        /// <summary>Noms des outils déclarés à Copilot pour la session.</summary>
        private HashSet<string> allowedTools;

        /// <summary>Historique de conversation mis à jour avec les appels et résultats d’outils.</summary>
        private IList<object> history;

        /// <summary>Callback facultatif pour les fragments de texte reçus en streaming.</summary>
        /// <value>Fonction appelée pour chaque fragment, ou null si le streaming est désactivé.</value>
        public Action<string> TextDelta { get; set; }

        /// <summary>Receives public reasoning and actual VBAi tool lifecycle events.</summary>
        /// <value>Callback receiving public lifecycle events, or null when the timeline is not subscribed.</value>
        public Action<CodexAgentActivity> ActivityUpdate { get; set; }

        /// <summary>Live activities and public snapshots are isolated by native session/block identity.</summary>
        private readonly Dictionary<string, CodexAgentActivity> pendingActivities = new Dictionary<string, CodexAgentActivity>();

        /// <summary>Retains bounded public snapshots to ignore late deltas and deduplicate message-side reasoning.</summary>
        private readonly Dictionary<string, CodexAgentActivity> activitySnapshots = new Dictionary<string, CodexAgentActivity>();

        /// <summary>Last public reasoning block for each native agent in the current native turn.</summary>
        private readonly Dictionary<string, string> latestReasoning = new Dictionary<string, string>();

        /// <summary>Monotonic native turn scope prevents reused provider block identifiers from colliding.</summary>
        private long nativeTurnSequence;

        /// <summary>Attend le délai natif des requêtes et réponses Copilot.</summary>
        internal Func<TimeSpan, Task> Delay = Task.Delay;

        /// <summary>Démarre le processus CLI sans préambule sur son entrée standard.</summary>
        internal Func<Process, bool> StartProcess = ProcessInput.StartWithoutPreamble;

        /// <summary>Crée un sérialiseur JSON configuré pour les messages de dix mégaoctets au plus.</summary>
        /// <returns>Sérialiseur de messages Copilot.</returns>
        private static JavaScriptSerializer Json() { return new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }; }

        /// <summary>Lit une valeur de dictionnaire comme texte.</summary>
        /// <param name="o">Dictionnaire source, éventuellement null.</param>
        /// <param name="k">Nom de la clé.</param>
        /// <returns>Valeur convertie, ou null si la clé manque.</returns>
        private static string Text(IDictionary<string, object> o, string k) { return ClaudeProtocol.Text(o, k); }

        /// <summary>Récupère une valeur de dictionnaire si elle-même est un dictionnaire.</summary>
        /// <param name="o">Dictionnaire source, éventuellement null.</param>
        /// <param name="k">Nom de la clé.</param>
        /// <returns>Dictionnaire associé à la clé, ou null.</returns>
        private static IDictionary<string, object> Object(IDictionary<string, object> o, string k)
        { object value; return o != null && o.TryGetValue(k, out value) ? value as IDictionary<string, object> : null; }

        /// <summary>Résout le CLI Copilot depuis sa variable de configuration ou son nom exécutable par défaut.</summary>
        /// <value>Valeur de VBAi_COPILOT_CLI, ou « copilot.exe ».</value>
        internal static string Executable { get { return Environment.GetEnvironmentVariable("VBAi_COPILOT_CLI") ?? "copilot.exe"; } }

        /// <summary>Ouvre la commande interactive de connexion Copilot.</summary>
        public static void StartLogin()
        {
            var info = new ProcessStartInfo(Executable, "login") { UseShellExecute = false, CreateNoWindow = false };
            ProviderSessionStorage.ConfigureCopilot(info); Process.Start(info);
        }

        /// <summary>Démarre brièvement le CLI, charge son catalogue de modèles et retourne un état d’accessibilité.</summary>
        /// <returns>Texte qui indique si Copilot est accessible et le nombre de modèles.</returns>
        public static async Task<string> ReadStatusAsync()
        {
            using (var client = new CopilotClient())
            {
                var models = await client.ListModelsAsync();
                return "Copilot accessible · " + models.Length + UiText.Get(" models. Authentication is managed by the GitHub CLI.");
            }
        }

        /// <summary>Démarre le CLI headless, valide la version 2 ou 3 du protocole et initialise le lecteur de messages.</summary>
        /// <returns>Tâche terminée lorsque le processus et le protocole Copilot sont initialisés.</returns>
        private async Task StartAsync()
        {
            if (disposed) throw new ObjectDisposedException(nameof(CopilotClient));
            if (process != null) return;
            var info = new ProcessStartInfo(Executable, "--headless --stdio --no-auto-update --log-level error")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = Path.GetTempPath()
            };
            ProviderSessionStorage.ConfigureCopilot(info);
            process = new Process { StartInfo = info };
            try
            {
                if (!StartProcess(process)) throw new InvalidOperationException(UiText.Get("Unable to start Copilot."));
            }
            catch (Exception ex) { process.Dispose(); process = null; throw new InvalidOperationException(UiText.Get("Install GitHub Copilot CLI and set VBAi_COPILOT_CLI to its executable if needed."), ex); }
            process.ErrorDataReceived += (s, e) => { }; // Drain without logging tokens or prompts.
            process.BeginErrorReadLine();
            _ = Task.Run(ReadLoop);
            var ping = await RequestAsync("ping", new { });
            int version;
            if (!int.TryParse(Text(ping, "protocolVersion"), out version) || version < 2 || version > 3)
                throw new InvalidOperationException(UiText.Get("Incompatible Copilot protocol version (versions 2 and 3 supported)."));
        }

        /// <summary>Demande au CLI le catalogue des modèles et ignore les entrées sans identifiant.</summary>
        /// <returns>Modèles que le CLI Copilot rend disponibles.</returns>
        public async Task<LlmModelOption[]> ListModelsAsync()
        {
            await StartAsync();
            var result = await RequestAsync("models.list", new { });
            return ClaudeProtocol.Array(result, "models").Select(x => ClaudeProtocol.Object(x))
                .Where(x => !string.IsNullOrWhiteSpace(Text(x, "id")))
                .Select(x => new LlmModelOption(Text(x, "id"), Text(x, "name") ?? Text(x, "id"), capabilities: LlmModelCapabilities.FromCopilot(x))).ToArray();
        }

        /// <summary>Crée une session, envoie l’historique et attend une réponse ou un délai maximal de cinq minutes.</summary>
        /// <param name="model">Identifiant du modèle Copilot sélectionné.</param>
        /// <param name="history">Historique de conversation transmis au modèle et enrichi par les appels d’outils.</param>
        /// <param name="tools">Définitions des outils VBE proposés.</param>
        /// <param name="toolHandler">Fonction d’exécution des outils VBE autorisés.</param>
        /// <returns>Message assistant contenant le texte final de Copilot.</returns>
        public async Task<IDictionary<string, object>> CompleteAsync(string model, IList<object> history, object[] tools,
            Func<string, string, Task<string>> toolHandler)
        {
            invoke = toolHandler ?? throw new InvalidOperationException(UiText.Get("Copilot VBA tools are not connected."));
            this.history = history;
            await StartAsync();
            var definitions = tools.Select(x => ClaudeProtocol.Object(ClaudeProtocol.Object(x)["function"])).ToArray();
            allowedTools = new HashSet<string>(definitions.Select(x => Text(x, "name")), StringComparer.Ordinal);
            EndActivities("interrupted");
            lock (gate) { pendingActivities.Clear(); activitySnapshots.Clear(); latestReasoning.Clear(); nativeTurnSequence = 0; }
            sessionId = Guid.NewGuid().ToString();
            completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var created = await RequestAsync("session.create", new
            {
                sessionId,
                model,
                clientName = "VBAi",
                tools = definitions,
                availableTools = allowedTools.ToArray(),
                requestPermission = true,
                streaming = TextDelta != null || ActivityUpdate != null,
                systemMessage = new
                {
                    mode = "replace",
                    content = "You are a VBA assistant embedded in the VBE. Use only the supplied VBA tools. The prompt contains JSON conversation history; answer its last user request. Re-read live VBA before any change.\n" +
                    string.Join("\n", history.Select(ClaudeProtocol.Object).Where(x => Text(x, "role") == "system").Select(x => Text(x, "content")))
                },
                enableFileHooks = false,
                enableSkills = false,
                enableHostGitOperations = false,
                enableOnDemandInstructionDiscovery = false,
                customAgents = new object[0],
                mcpServers = new Dictionary<string, object>()
            });
            if (Text(created, "sessionId") != sessionId) throw new InvalidOperationException("Session Copilot inattendue.");
            await RequestAsync("session.send", new { sessionId, prompt = Json().Serialize(history) });
            if (await Task.WhenAny(completion.Task, Delay(TimeSpan.FromMinutes(5))) != completion.Task)
            {
                Dispose(); throw new TimeoutException(UiText.Get("Copilot did not finish its response within the time limit."));
            }
            return new Dictionary<string, object> { ["role"] = "assistant", ["content"] = await completion.Task };
        }

        /// <summary>Envoie une requête JSON-RPC, attend sa réponse pendant au plus 45 secondes puis la retourne.</summary>
        /// <param name="method">Méthode du CLI appelée.</param>
        /// <param name="parameters">Paramètres de la méthode.</param>
        /// <returns>Résultat de la requête.</returns>
        private Task<IDictionary<string, object>> RequestAsync(string method, object parameters) => RequestCoreAsync(method, parameters, null, null);

        /// <summary>Registers an optional captured turn owner before sending its RPC.</summary>
        private async Task<IDictionary<string, object>> RequestCoreAsync(string method, object parameters, string ownerSession, TaskCompletionSource<string> ownerCompletion)
        {
            var source = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
            int id;
            lock (gate) { if (disposed) throw new ObjectDisposedException(nameof(CopilotClient)); id = ++nextId; pending.Add(id, source); if (ownerCompletion != null) pendingOwners.Add(id, Tuple.Create(ownerSession, ownerCompletion)); }
            try
            {
                Send(new { jsonrpc = "2.0", id, method, @params = parameters });
                if (await Task.WhenAny(source.Task, Delay(TimeSpan.FromSeconds(45))) != source.Task)
                    throw new TimeoutException(UiText.Get("Copilot is not responding to ") + method + UiText.Get(". Check the CLI connection."));
                return await source.Task;
            }
            finally { lock (gate) { pending.Remove(id); pendingOwners.Remove(id); } }
        }

        /// <summary>Encode puis écrit un message avec son en-tête Content-Length sur l’entrée standard.</summary>
        /// <param name="message">Message JSON-RPC à transmettre.</param>
        private void Send(object message)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(Json().Serialize(message));
            byte[] header = Encoding.ASCII.GetBytes("Content-Length: " + bytes.Length + "\r\n\r\n");
            lock (gate)
            {
                if (disposed) throw new ObjectDisposedException(nameof(CopilotClient));
                var stream = process.StandardInput.BaseStream;
                stream.Write(header, 0, header.Length); stream.Write(bytes, 0, bytes.Length); stream.Flush();
            }
        }

        /// <summary>Lit les en-têtes et corps encadrés, valide leur taille puis transmet chaque message au répartiteur.</summary>
        private void ReadLoop()
        {
            try
            {
                var stream = process.StandardOutput.BaseStream;
                while (!disposed)
                {
                    var header = new StringBuilder();
                    while (!header.ToString().EndsWith("\r\n\r\n", StringComparison.Ordinal))
                    {
                        int next = stream.ReadByte(); if (next < 0) throw new EndOfStreamException(UiText.Get("Copilot stopped."));
                        header.Append((char)next); if (header.Length > 4096) throw new InvalidDataException(UiText.Get("Invalid Copilot header."));
                    }
                    int size = 0;
                    foreach (string line in header.ToString().Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                        if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) int.TryParse(line.Substring(15).Trim(), out size);
                    if (size <= 0 || size > 10 * 1024 * 1024) throw new InvalidDataException(UiText.Get("Invalid Copilot message size."));
                    var bytes = new byte[size]; int offset = 0;
                    while (offset < size) { int count = stream.Read(bytes, offset, size - offset); if (count == 0) throw new EndOfStreamException(); offset += count; }
                    Dispatch(Json().DeserializeObject(Encoding.UTF8.GetString(bytes)) as IDictionary<string, object>);
                }
            }
            catch (Exception ex) { Fail(ex); }
        }

        /// <summary>Associe les réponses aux requêtes et traite événements, appels d’outils et décisions de permission.</summary>
        /// <param name="message">Message JSON-RPC décodé.</param>
        private void Dispatch(IDictionary<string, object> message)
        {
            string method = Text(message, "method");
            if (method == null)
            {
                int id; TaskCompletionSource<IDictionary<string, object>> source;
                Tuple<string, TaskCompletionSource<string>> owner = null;
                Exception error = null; bool turnFaulted = false;
                if (int.TryParse(Text(message, "id"), out id)) lock (gate) if (pending.TryGetValue(id, out source))
                {
                    if (message.ContainsKey("error"))
                    {
                        error = new InvalidOperationException(UiText.Get("Copilot refused the request. Check sign-in, subscription and model in the CLI."));
                        bool requestFaulted = source.TrySetException(error);
                        // Settle the captured active turn before the reader can admit a following idle event.
                        if (requestFaulted && pendingOwners.TryGetValue(id, out owner)) turnFaulted = TryFaultTurn(owner.Item1, owner.Item2, error);
                    }
                    else source.TrySetResult(Object(message, "result") ?? new Dictionary<string, object>());
                }
                if (turnFaulted) EndTurnActivities(owner.Item1, owner.Item2, error);
                return;
            }
            var parameters = Object(message, "params");
            if (method == "session.event" && Text(parameters, "sessionId") == sessionId && completion != null && !completion.Task.IsCompleted)
            {
                var evt = Object(parameters, "event"); var data = Object(evt, "data"); string type = Text(evt, "type");
                if (type == "assistant.turn_start") { EndActivities("interrupted", true); lock (gate) { latestReasoning.Clear(); nativeTurnSequence++; } }
                else if (type == "assistant.turn_end") EndActivities("interrupted", true);
                else if (type == "assistant.reasoning_delta" || type == "assistant.reasoning")
                    ReceiveReasoning(evt, data, type == "assistant.reasoning");
                else if (type == "assistant.message")
                {
                    answer = Text(data, "content");
                    ReceiveMessageReasoning(evt, data);
                }
                else if (type == "assistant.message_delta")
                {
                    string fragment = Text(data, "deltaContent");
                    if (!string.IsNullOrEmpty(fragment))
                    {
                        if (ui != null) ui.Post(_ => { if (!disposed) TextDelta?.Invoke(fragment); }, null);
                        else TextDelta?.Invoke(fragment);
                    }
                }
                else if (type == "session.idle") { EndActivities("interrupted"); completion.TrySetResult(answer ?? ""); }
                else if (type == "session.error") Fail(new InvalidOperationException(UiText.Get("The Copilot session failed. Check its status in the CLI.")));
                else if (type == "external_tool.requested") RunTool(data, null);
                else if (type == "permission.requested") _ = DenyAsync(data);
                return;
            }
            object requestId;
            if (message.TryGetValue("id", out requestId))
            {
                if (method == "tool.call" && Text(parameters, "sessionId") == sessionId) RunTool(parameters, requestId);
                else if (method == "permission.request") Send(new { jsonrpc = "2.0", id = requestId, result = new { result = new { kind = Text(parameters, "sessionId") == sessionId ? PermissionDecision(Object(parameters, "permissionRequest")) : "denied-by-rules" } } });
                else Send(new { jsonrpc = "2.0", id = requestId, error = new { code = -32601, message = "Unsupported client method" } });
            }
        }

        /// <summary>Refuse une permission Copilot selon la règle des outils personnalisés puis signale les erreurs de transport.</summary>
        /// <returns>Tâche terminée après le refus de la permission ou la gestion de son échec.</returns>
        /// <param name="data">Données de la demande de permission.</param>
        private async Task DenyAsync(IDictionary<string, object> data)
        {
            string ownerSession = sessionId; var ownerCompletion = completion;
            try { await RequestCoreAsync("session.permissions.handlePendingPermissionRequest", new { sessionId = ownerSession, requestId = Text(data, "requestId"), result = new { kind = PermissionDecision(Object(data, "permissionRequest")) } }, ownerSession, ownerCompletion); }
            catch (Exception ex) { if (TryFaultTurn(ownerSession, ownerCompletion, ex)) EndTurnActivities(ownerSession, ownerCompletion, ex); }
        }

        /// <summary>Déduplique et exécute une demande d’outil, puis renvoie son résultat au format du protocole négocié.</summary>
        /// <param name="data">Données de la demande d’outil.</param>
        /// <param name="legacyId">Identifiant JSON-RPC pour l’ancien protocole, ou null.</param>
        private void RunTool(IDictionary<string, object> data, object legacyId)
        {
            string ownerSession = sessionId;
            var ownerCompletion = completion;
            string requestId = Text(data, "requestId") ?? Convert.ToString(legacyId);
            lock (gate) if (!handledTools.Add(ownerSession + ":" + requestId)) return;
            Func<Task> run = async () =>
            {
                string name = Text(data, "toolName");
                string arguments = null;
                string callId = Text(data, "toolCallId") ?? requestId;
                string activityId = sessionId + ":tool:" + callId;
                bool settled = false;
                try
                {
                    if (disposed || ownerSession != sessionId || ownerCompletion != completion || ownerCompletion == null || ownerCompletion.Task.IsCompleted) return;
                    arguments = Json().Serialize(data.ContainsKey("arguments") ? data["arguments"] : new { });
                    PublishTurnActivity(ProviderActivityProjection.Tool(activityId, name, arguments, "inProgress"), ownerSession, ownerCompletion);
                    // A synchronous progress subscriber may dispose, settle or replace the native turn.
                    if (disposed || ownerSession != sessionId || ownerCompletion != completion || ownerCompletion.Task.IsCompleted) return;
                    var ownerTools = allowedTools; var ownerInvoke = invoke; var ownerHistory = history;
                    string output = ownerTools != null && ownerTools.Contains(name) && ownerInvoke != null
                        ? await ownerInvoke(name, arguments)
                        : "Tool not available.";
                    if (ownerSession != sessionId || ownerCompletion != completion) return;
                    PublishTurnActivity(ProviderActivityProjection.Tool(activityId, name, arguments,
                        ProviderActivityProjection.ToolOutcome(output), ProviderActivityProjection.ToolDiagnostic(output)), ownerSession, ownerCompletion);
                    settled = true;
                    ownerHistory.Add(new { role = "assistant", content = (string)null, tool_calls = new[] { new { id = callId, type = "function", function = new { name, arguments } } } });
                    ownerHistory.Add(new { role = "tool", tool_call_id = callId, content = output });
                    var result = new { textResultForLlm = output, resultType = "success" };
                    if (legacyId != null) Send(new { jsonrpc = "2.0", id = legacyId, result = new { result } });
                    else await RequestCoreAsync("session.tools.handlePendingToolCall", new { sessionId = ownerSession, requestId, result }, ownerSession, ownerCompletion);
                }
                catch (Exception ex)
                {
                    if (!TryFaultTurn(ownerSession, ownerCompletion, ex)) return;
                    if (!settled) PublishTurnActivity(ProviderActivityProjection.Tool(activityId, name, arguments, ex is OperationCanceledException ? "interrupted" : "failed", ex.Message), ownerSession, ownerCompletion, true);
                    EndTurnActivities(ownerSession, ownerCompletion, ex);
                }
            };
            if (ui != null) ui.Post(async _ => await run(), null); else _ = run();
        }

        /// <summary>Projects documented Copilot reasoning events; opaque fields are never read.</summary>
        /// <param name="evt">Native event envelope.</param>
        /// <param name="data">Event data with reasoningId and content/deltaContent strings.</param>
        /// <param name="complete">Whether the content is an authoritative final snapshot.</param>
        private void ReceiveReasoning(IDictionary<string, object> evt, IDictionary<string, object> data, bool complete)
        {
            if (data == null || !data.TryGetValue("reasoningId", out var rawId) || !(rawId is string id) || string.IsNullOrWhiteSpace(id) ||
                !data.TryGetValue(complete ? "content" : "deltaContent", out var rawText) || !(rawText is string text)) return;
            string agent = Text(evt, "agentId") ?? "main";
            string activityId = sessionId + ":turn:" + nativeTurnSequence + ":reasoning:" + agent + ":" + id;
            lock (gate) latestReasoning[agent] = activityId;
            PublishActivity(new CodexAgentActivity { Id = activityId, Kind = "reasoning", Title = UiText.Get("Reasoning"), Detail = text, Append = !complete, Status = complete ? "completed" : "inProgress" });
        }

        /// <summary>Projects the documented readable message-side fallback without repeating a completed reasoning block.</summary>
        /// <param name="evt">Native event envelope.</param>
        /// <param name="data">Message data; reasoningOpaque and encryptedContent remain untouched.</param>
        private void ReceiveMessageReasoning(IDictionary<string, object> evt, IDictionary<string, object> data)
        {
            if (data == null || !data.TryGetValue("reasoningText", out var rawText) || !(rawText is string text) || string.IsNullOrWhiteSpace(text)) return;
            string agent = Text(evt, "agentId") ?? "main", id = null;
            lock (gate)
            {
                if (latestReasoning.TryGetValue(agent, out var previousId) && activitySnapshots.TryGetValue(previousId, out var previous))
                {
                    if (previous.Status == "inProgress") id = previousId;
                    else if (previous.Detail == CodexAgentActivity.Limit(text)) return;
                }
            }
            if (id == null)
            {
                if (!data.TryGetValue("messageId", out var rawId) || !(rawId is string messageId) || string.IsNullOrWhiteSpace(messageId)) return;
                id = sessionId + ":turn:" + nativeTurnSequence + ":reasoning:" + agent + ":message:" + messageId;
            }
            lock (gate) latestReasoning[agent] = id;
            PublishActivity(new CodexAgentActivity { Id = id, Kind = "reasoning", Title = UiText.Get("Reasoning"), Detail = text, Status = "completed" });
        }

        /// <summary>Publishes bounded public state on the captured UI context, preserving terminal snapshots against late deltas.</summary>
        /// <param name="activity">Actual provider/tool event.</param>
        private void PublishActivity(CodexAgentActivity activity) => PublishTurnActivity(activity, sessionId, completion);

        /// <summary>Publishes activity only within the captured native turn, including queued callbacks.</summary>
        private void PublishTurnActivity(CodexAgentActivity activity, string ownerSession, TaskCompletionSource<string> ownerCompletion, bool isolatePublicationErrors = false)
        {
            var callback = ActivityUpdate;
            activity.Detail = CodexAgentActivity.Limit(activity.Detail);
            lock (gate)
            {
                if (ownerSession != sessionId || ownerCompletion != completion) return;
                activitySnapshots.TryGetValue(activity.Id, out var previous);
                if (activity.Append && activity.Status == "inProgress" && previous != null && previous.Status != "inProgress") return;
                var snapshot = new CodexAgentActivity
                {
                    Id = activity.Id, Kind = activity.Kind, Title = activity.Title ?? previous?.Title,
                    Detail = CodexAgentActivity.Limit(activity.Append ? (previous?.Detail ?? "") + activity.Detail : activity.Detail),
                    Status = activity.Status, DurationMs = activity.DurationMs ?? previous?.DurationMs
                };
                activitySnapshots[activity.Id] = snapshot;
                if (activity.Status == "inProgress") pendingActivities[activity.Id] = snapshot; else pendingActivities.Remove(activity.Id);
            }
            Action deliver = () =>
            {
                if (ownerSession != sessionId || ownerCompletion != completion || (disposed && activity.Status == "inProgress")) return;
                try { callback?.Invoke(activity); }
                catch (Exception error) when (isolatePublicationErrors) { LogActivityPublicationFailure(error); }
            };
            try { if (ui != null) ui.Post(_ => deliver(), null); else deliver(); }
            catch (Exception error) when (isolatePublicationErrors) { LogActivityPublicationFailure(error); }
        }

        /// <summary>Reports publication failures without replacing the primary owned operation failure.</summary>
        private static void LogActivityPublicationFailure(Exception error)
        {
            try { LoadLog.Write("Copilot activity publication failed: " + error.GetType().Name); }
            catch { /* Logging must not escape into the reader's global transport failure path. */ }
        }
        /// <summary>Settles visible partial work when the native turn ends without its own final receipt.</summary>
        /// <param name="status">Failed or interrupted terminal state.</param>
        /// <param name="reasoningOnly">Whether only native thinking blocks should settle, leaving admitted tools untouched.</param>
        private void EndActivities(string status, bool reasoningOnly = false)
        {
            CodexAgentActivity[] pending;
            lock (gate)
            {
                pending = pendingActivities.Values.Where(activity => !reasoningOnly || activity.Kind == "reasoning").ToArray();
                foreach (var activity in pending) pendingActivities.Remove(activity.Id);
            }
            foreach (var activity in pending) PublishActivity(new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind, Status = status, Append = true });
        }

        /// <summary>Closes only the still-active captured turn under the RPC admission lock.</summary>
        private bool TryFaultTurn(string ownerSession, TaskCompletionSource<string> ownerCompletion, Exception error)
        {
            lock (gate)
            {
                if (disposed || ownerCompletion == null || ownerSession != sessionId || ownerCompletion != completion || ownerCompletion.Task.IsCompleted) return false;
                string status = error is OperationCanceledException ? "interrupted" : "failed";
                var activities = pendingActivities.Values.Select(activity => new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind, Status = status, Append = true }).ToArray();
                pendingActivities.Clear();
                foreach (var activity in activities)
                    if (activitySnapshots.TryGetValue(activity.Id, out var snapshot)) snapshot.Status = status;
                failedTurnActivities[ownerCompletion] = activities;
                return ownerCompletion.TrySetException(error);
            }
        }

        /// <summary>Settles activities outside the admission lock without publishing into a replacement turn.</summary>
        private void EndTurnActivities(string ownerSession, TaskCompletionSource<string> ownerCompletion, Exception error)
        {
            CodexAgentActivity[] activities;
            lock (gate)
            {
                if (!failedTurnActivities.TryGetValue(ownerCompletion, out activities)) return;
                failedTurnActivities.Remove(ownerCompletion);
                if (ownerSession != sessionId || ownerCompletion != completion) return;
            }
            foreach (var activity in activities)
                try { PublishTurnActivity(activity, ownerSession, ownerCompletion, true); }
                catch (Exception publicationError) { LogActivityPublicationFailure(publicationError); }
        }
        /// <summary>Termine en erreur les requêtes et la réponse qui sont encore en attente.</summary>
        /// <param name="error">Erreur transmise aux opérations en attente.</param>
        private void Fail(Exception error)
        {
            EndActivities(error is OperationCanceledException ? "interrupted" : "failed");
            lock (gate) foreach (var request in pending.Values) request.TrySetException(error);
            completion?.TrySetException(error);
        }

        /// <summary>Autorise seulement un outil personnalisé déclaré et encore actif.</summary>
        /// <param name="permission">Description de la permission demandée.</param>
        /// <returns>Décision « approved » ou « denied-by-rules » du protocole Copilot.</returns>
        private string PermissionDecision(IDictionary<string, object> permission)
        {
            // This only routes our registered custom tool to LlmVbeTools.InvokeAsync.
            // That entry point still enforces mode/project, ReadOnly, AskEachTime and live SHA.
            // Shell, file, network, MCP and unknown tools never reach that entry point.
            return !disposed && Text(permission, "kind") == "custom-tool" && allowedTools != null &&
                allowedTools.Contains(Text(permission, "toolName")) ? "approved" : "denied-by-rules";
        }

        /// <summary>Annule les opérations en attente et arrête le processus Copilot CLI.</summary>
        public void Dispose()
        {
            lock (gate) { if (disposed) return; disposed = true; pendingOwners.Clear(); }
            Fail(new OperationCanceledException());
            if (process != null) { try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { } finally { process.Dispose(); } }
        }
    }
}
