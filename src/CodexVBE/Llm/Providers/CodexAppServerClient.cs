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
    /// <summary>Abstraction du processus app-server et de ses échanges par lignes.</summary>
    internal interface ICodexAppServerTransport : IDisposable
    {
        /// <summary>Survient lorsqu’une ligne est lue sur la sortie standard.</summary>
        event Action<string> LineReceived;
        /// <summary>Survient lorsque le transport cesse de fonctionner.</summary>
        event Action<Exception> Exited;
        /// <summary>Indique si le processus app-server est actif.</summary>
        /// <value>Vrai si le processus est actif.</value>
        bool IsRunning { get; }
        /// <summary>Démarre le processus et commence à lire ses sorties.</summary>
        void Start();
        /// <summary>Écrit une ligne de protocole vers le processus.</summary>
        /// <param name="line">Ligne JSON-RPC à transmettre.</param>
        void Send(string line);
    }

    /// <summary>Transport app-server qui lance le client CLI Codex et échange des lignes UTF-8.</summary>
    internal sealed class CodexProcessTransport : ICodexAppServerTransport
    {
        /// <summary>Vérifie la présence du binaire Codex installé à son emplacement local.</summary>
        internal Func<string, bool> InstalledExists = File.Exists;
        internal Func<string, string[]> GetDirectories = Directory.GetDirectories;
        internal Func<string, DateTime> GetLastWriteTimeUtc = File.GetLastWriteTimeUtc;
        /// <summary>Démarre le processus Codex sans préambule UTF-8 parasite sur l’entrée standard.</summary>
        internal Func<Process, bool> StartProcess = ProcessInput.StartWithoutPreamble;
        /// <summary>Processus CLI Codex détenu par le transport.</summary>
        private Process process;
        /// <summary>Relaye les lignes reçues sur la sortie standard.</summary>
        public event Action<string> LineReceived;
        /// <summary>Relaye la fin du processus comme une erreur de transport.</summary>
        public event Action<Exception> Exited;
        /// <summary>Indique si le processus détenu est vivant.</summary>
        /// <value>Vrai si le processus détenu est actif.</value>
        public bool IsRunning { get { return process != null && !process.HasExited; } }

        /// <summary>Lance « codex app-server », évite un préambule UTF-8 sur l’entrée standard et active la lecture asynchrone.</summary>
        public void Start()
        {
            string executable = CodexCliLocator.Resolve(InstalledExists, GetDirectories, GetLastWriteTimeUtc);
            var info = new ProcessStartInfo(executable, "app-server") {
                UseShellExecute = false, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                CreateNoWindow = true, StandardOutputEncoding = new UTF8Encoding(false),
                StandardErrorEncoding = new UTF8Encoding(false)
            };
            ProviderSessionStorage.ConfigureCodex(info);
            process = new Process { StartInfo = info, EnableRaisingEvents = true };
            process.Exited += (sender, args) => Exited?.Invoke(new InvalidOperationException(UiText.Get("Codex app-server stopped.")));
            if (!StartProcess(process))
                throw new InvalidOperationException(UiText.Get("Unable to start codex app-server."));
            process.OutputDataReceived += (sender, args) => { if (args.Data != null) LineReceived?.Invoke(args.Data); };
            // Drain stderr; diagnostics are never protocol messages.
            process.ErrorDataReceived += (sender, args) => { };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
        }

        /// <summary>Écrit la ligne suivie d’un saut de ligne sur l’entrée standard UTF-8 sans BOM.</summary>
        /// <param name="line">Ligne JSON-RPC à envoyer.</param>
        public void Send(string line)
        {
            if (!IsRunning) throw new InvalidOperationException("Codex app-server is not running.");
            byte[] bytes = new UTF8Encoding(false).GetBytes(line + "\n");
            process.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
            process.StandardInput.BaseStream.Flush();
        }

        /// <summary>Arrête le processus encore actif et libère ses ressources.</summary>
        public void Dispose()
        {
            if (process == null) return;
            try { if (!process.HasExited) process.Kill(); } catch { }
            process.Dispose();
            process = null;
        }
    }

    // Owns a single Codex CLI child process. Codex itself owns ChatGPT authentication.
    /// <summary>Gère une session app-server Codex, ses requêtes JSON-RPC et les appels d’outils du VBE.</summary>
    internal sealed class CodexAppServerClient : IDisposable
    {
        /// <summary>Contexte utilisé pour publier les mises à jour sur le fil de l’interface.</summary>
        private readonly SynchronizationContext ui;
        /// <summary>Outils VBE exécutés à la demande du processus Codex.</summary>
        private readonly LlmVbeTools tools;
        /// <summary>Exécute un outil VBE appelé par le serveur Codex avec son nom et ses arguments JSON.</summary>
        internal Func<string, string, Task<string>> InvokeTool;
        /// <summary>Callback qui signale les étapes de connexion et les appels d’outils.</summary>
        private readonly Action<string> progress;
        /// <summary>Réglages LLM fournis à cette session.</summary>
        private readonly LlmSettings settings;
        /// <summary>Réponses en attente, indexées par identifiant JSON-RPC.</summary>
        private readonly Dictionary<int, TaskCompletionSource<IDictionary<string, object>>> requests =
            new Dictionary<int, TaskCompletionSource<IDictionary<string, object>>>();
        /// <summary>Verrou qui protège les échanges et le registre des requêtes.</summary>
        private readonly object gate = new object();
        /// <summary>Transport utilisé pour lire et écrire le protocole app-server.</summary>
        private readonly ICodexAppServerTransport transport;
        /// <summary>Indique si le transport a été démarré.</summary>
        private bool transportStarted;
        /// <summary>Dernier identifiant JSON-RPC attribué aux requêtes.</summary>
        private int nextId;
        /// <summary>Identifiant du fil Codex créé ou repris.</summary>
        private string threadId;
        /// <summary>Achèvement de la réponse du tour actif.</summary>
        private TaskCompletionSource<string> turnDone;
        /// <summary>Dernier texte final reçu pour le tour actif.</summary>
        private string finalText;
        /// <summary>Indique si le client a été libéré.</summary>
        private bool disposed;
        /// <summary>Identifiant du tour Codex en cours.</summary>
        private string activeTurnId;
        /// <summary>Indique qu’une interruption a été demandée pendant le démarrage ou le tour.</summary>
        private bool interruptRequested;
        /// <summary>Identifiant du fil courant.</summary>
        /// <value>Identifiant du fil, ou null avant sa création.</value>
        public string ThreadId { get { return threadId; } }
        /// <summary>Survient quand le fil Codex est prêt à recevoir des tours.</summary>
        public event Action<string> ThreadReady;
        /// <summary>Publie le texte, le raisonnement et l’état des appels d’outils dans la conversation.</summary>
        public event Action<string, string, string, bool> ChatUpdate;

        /// <summary>Crée un client qui utilise le transport CLI standard.</summary>
        /// <param name="ui">Contexte du fil de l’interface.</param>
        /// <param name="tools">Outils disponibles au processus Codex.</param>
        /// <param name="progress">Callback de progression.</param>
        /// <param name="settings">Réglages de session.</param>
        /// <param name="resumeThreadId">Fil à reprendre, ou null pour en créer un.</param>
        public CodexAppServerClient(SynchronizationContext ui, LlmVbeTools tools, Action<string> progress, LlmSettings settings, string resumeThreadId = null)
            : this(ui, tools, progress, settings, resumeThreadId, new CodexProcessTransport()) { }

        /// <summary>Crée un client avec un transport injecté, notamment pour les transports contrôlés par l’appelant.</summary>
        /// <param name="ui">Contexte du fil de l’interface.</param>
        /// <param name="tools">Outils disponibles au processus Codex.</param>
        /// <param name="progress">Callback de progression, remplacé par un callback vide si null.</param>
        /// <param name="settings">Réglages de session.</param>
        /// <param name="resumeThreadId">Fil à reprendre, ou null pour en créer un.</param>
        /// <param name="transport">Transport app-server à utiliser.</param>
        internal CodexAppServerClient(SynchronizationContext ui, LlmVbeTools tools, Action<string> progress,
            LlmSettings settings, string resumeThreadId, ICodexAppServerTransport transport)
        {
            this.ui = ui ?? throw new ArgumentNullException(nameof(ui));
            this.tools = tools ?? throw new ArgumentNullException(nameof(tools));
            InvokeTool = tools.InvokeAsync;
            this.progress = progress ?? (_ => { });
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
            this.transport = transport ?? throw new ArgumentNullException(nameof(transport));
            this.transport.LineReceived += OnLine;
            this.transport.Exited += FailPending;
            threadId = resumeThreadId;
        }

        /// <summary>Interroge toutes les pages du catalogue Codex et convertit les efforts pris en charge.</summary>
        /// <returns>Modèles exposés par le compte Codex.</returns>
        public async Task<LlmModelOption[]> ListModelsAsync()
        {
            if (!transportStarted) await StartAsync();
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

        /// <summary>Démarre un tour dans le fil et attend son texte final ou son échec.</summary>
        /// <param name="prompt">Texte transmis comme entrée du tour.</param>
        /// <param name="model">Identifiant de modèle sélectionné.</param>
        /// <param name="effort">Niveau d’effort demandé.</param>
        /// <returns>Texte final du tour.</returns>
        /// <exception cref="ArgumentException">Le prompt est vide ou ne contient que des espaces.</exception>
        /// <exception cref="InvalidOperationException">Un tour est déjà actif.</exception>
        public async Task<string> TurnAsync(string prompt, string model, string effort)
        {
            if (string.IsNullOrWhiteSpace(prompt)) throw new ArgumentException("A prompt is required.");
            if (turnDone != null) throw new InvalidOperationException("A Codex turn is already running.");
            finalText = null;
            interruptRequested = false;
            turnDone = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                if (!transportStarted) await StartAsync();
                if (interruptRequested) throw new OperationCanceledException();
                var started = await RequestAsync("turn/start", new {
                    threadId,
                    model,
                    effort,
                    summary = "auto",
                    input = new[] { new { type = "text", text = prompt } }
                });
                activeTurnId = GetString(GetObject(GetObject(started, "result"), "turn"), "id") ?? activeTurnId;
                if (interruptRequested) await InterruptAsync();
                return await turnDone.Task;
            }
            finally { turnDone = null; activeTurnId = null; }
        }

        /// <summary>Demande l’interruption du tour actif ; mémorise la demande si son identifiant n’est pas encore disponible.</summary>
        /// <returns>Tâche terminée après l’envoi de la demande, s’il y a un tour identifié.</returns>
        public async Task InterruptAsync()
        {
            interruptRequested = true;
            if (turnDone != null && activeTurnId != null)
                await RequestAsync("turn/interrupt", new { threadId, turnId = activeTurnId });
        }

        /// <summary>Démarre le transport, initialise le protocole, vérifie le compte ChatGPT et crée ou reprend le fil.</summary>
        /// <returns>Tâche terminée lorsque le transport, le compte et le fil sont prêts.</returns>
        private async Task StartAsync()
        {
            try
            {
                transport.Start();
                transportStarted = true;
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

        /// <summary>Envoie une requête JSON-RPC et retourne une tâche complétée à la réception de sa réponse.</summary>
        /// <param name="method">Méthode app-server appelée.</param>
        /// <param name="parameters">Paramètres sérialisés dans la requête.</param>
        /// <returns>Tâche contenant la réponse du serveur.</returns>
        private Task<IDictionary<string, object>> RequestAsync(string method, object parameters)
        {
            if (disposed || !transportStarted || !transport.IsRunning)
                throw new InvalidOperationException("Codex app-server is not running.");
            int id;
            var completion = new TaskCompletionSource<IDictionary<string, object>>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (gate) { id = ++nextId; requests.Add(id, completion); }
            try { Send(new { method, id, @params = parameters }); }
            catch { lock (gate) requests.Remove(id); throw; }
            return completion.Task;
        }

        /// <summary>Sérialise un message JSON-RPC et l’envoie sous le verrou du transport.</summary>
        /// <param name="message">Objet sérialisable à transmettre.</param>
        private void Send(object message)
        {
            string line = NewJson().Serialize(message);
            lock (gate)
            {
                if (!transportStarted || !transport.IsRunning) throw new InvalidOperationException("Codex app-server is not running.");
                transport.Send(line);
            }
        }

        /// <summary>Traite une réponse JSON-RPC, un appel d’outil ou une notification de tour reçue.</summary>
        /// <param name="line">Ligne reçue sur le transport.</param>
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
                        object summaryIndex;
                        if (parameters.TryGetValue("summaryIndex", out summaryIndex) &&
                            Convert.ToInt32(summaryIndex) > 0)
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
                            {
                                var parts = ((object[])summary).Select(part => {
                                    var value = part as IDictionary<string, object>;
                                    return value != null ? GetString(value, "text") : part as string;
                                }).Where(part => !string.IsNullOrWhiteSpace(part));
                                var text = string.Join("\n\n", parts);
                                PublishUpdate("summary", GetString(item, "id"),
                                    string.IsNullOrWhiteSpace(text) ? null : text, true);
                            }
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

        /// <summary>Publie une mise à jour de conversation sur le contexte de l’interface.</summary>
        /// <param name="kind">Catégorie du contenu publié.</param>
        /// <param name="id">Identifiant stable de l’élément mis à jour.</param>
        /// <param name="text">Texte ajouté ou contenu final, éventuellement null.</param>
        /// <param name="complete">Indique si l’élément est terminé.</param>
        private void PublishUpdate(string kind, string id, string text, bool complete)
        {
            if (string.IsNullOrEmpty(id)) return;
            ui.Post(_ => { if (!disposed) ChatUpdate?.Invoke(kind, id, text, complete); }, null);
        }

        /// <summary>Exécute un outil VBE demandé par le serveur puis lui renvoie son résultat.</summary>
        /// <param name="requestId">Identifiant de la requête d’outil à répondre.</param>
        /// <param name="parameters">Paramètres de l’appel d’outil.</param>
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
                    string output = await InvokeTool(name, arguments);
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

        /// <summary>Termine en erreur toutes les requêtes et le tour encore en attente.</summary>
        /// <param name="error">Erreur transmise aux opérations en attente.</param>
        private void FailPending(Exception error)
        {
            lock (gate)
            {
                foreach (var pending in requests.Values) pending.TrySetException(error);
                requests.Clear();
            }
            turnDone?.TrySetException(error);
        }

        /// <summary>Crée un sérialiseur configuré pour accepter des messages jusqu’à dix mégaoctets.</summary>
        /// <returns>Sérialiseur JSON utilisé par le protocole.</returns>
        private static JavaScriptSerializer NewJson() { return new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }; }
        /// <summary>Récupère une propriété uniquement si elle contient un dictionnaire objet.</summary>
        /// <param name="value">Dictionnaire source, éventuellement null.</param>
        /// <param name="key">Nom de la propriété.</param>
        /// <returns>Dictionnaire contenu dans la propriété, ou null.</returns>
        private static IDictionary<string, object> GetObject(IDictionary<string, object> value, string key)
        {
            object raw;
            return value != null && value.TryGetValue(key, out raw) ? raw as IDictionary<string, object> : null;
        }
        /// <summary>Récupère une propriété et convertit sa valeur en chaîne.</summary>
        /// <param name="value">Dictionnaire source, éventuellement null.</param>
        /// <param name="key">Nom de la propriété.</param>
        /// <returns>Valeur convertie, ou null si la propriété est absente.</returns>
        private static string GetString(IDictionary<string, object> value, string key)
        {
            object raw;
            return value != null && value.TryGetValue(key, out raw) ? Convert.ToString(raw) : null;
        }

        /// <summary>Libère le client, termine les opérations en attente et ferme le transport.</summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            FailPending(new ObjectDisposedException(nameof(CodexAppServerClient)));
            transport.Dispose();
        }
    }
}
