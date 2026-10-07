using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{
    // This client speaks the Chat Completions function-call format. No credential is persisted.
    /// <summary>Client HTTP compatible avec les fournisseurs LLM et leur protocole de conversation.</summary>
    internal sealed class LlmChatClient : IDisposable
    {

        /// <summary>Client HTTP utilisé pour appeler le fournisseur courant.</summary>
        private readonly HttpClient http;

        /// <summary>Sérialiseur JSON des requêtes et réponses fournisseur.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };

        /// <summary>URL validée du point de terminaison de conversation.</summary>
        private readonly Uri endpoint;

        /// <summary>Identifiant du modèle sélectionné.</summary>
        private readonly string model;

        /// <summary>Immutable provider connection for this client lifetime.</summary>
        private readonly LlmConnectionProfile connection;

        /// <summary>Explicit catalog capabilities captured with this model connection.</summary>
        private readonly LlmModelCapabilities modelCapabilities;

        /// <summary>Fournisseur et protocole associés au client.</summary>
        private readonly LlmProvider provider;

        /// <summary>Client Copilot utilisé lorsque ce fournisseur est sélectionné.</summary>
        private readonly CopilotClient copilot;

        /// <summary>Validated Ollama-only sampling snapshot for this client's lifetime.</summary>
        private readonly double? ollamaTemperature, ollamaTopP;

        /// <summary>Source d’annulation liée à la durée de vie du client.</summary>
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();

        /// <summary>Empêche la libération répétée des clients et jetons.</summary>
        private bool disposed;

        /// <summary>Reçoit les fragments de texte émis pendant une réponse en flux.</summary>
        /// <value>Action appelée pour chaque fragment de texte reçu en flux, ou null si le flux est désactivé.</value>
        public Action<string> TextDelta { get; set; }

        /// <summary>Observes request-scoped public reasoning and tool activity without changing provider messages.</summary>
        public Action<CodexAgentActivity> ActivityUpdate { get; set; }

        /// <summary>Last streamed response metadata; contains no request or response content.</summary>
        /// <value>Current last stream diagnostics exposed by llm chat client.</value>
        internal StreamDiagnostics LastStreamDiagnostics { get; private set; }

        /// <summary>Traite un appel d’outil retourné par le fournisseur.</summary>
        /// <value>Délégué qui reçoit le nom et les arguments JSON d’un outil, ou null si aucun outil n’est disponible.</value>
        public Func<string, string, Task<string>> ToolHandler { get; set; }

        /// <summary>Crée le gestionnaire HTTP utilisé par défaut, remplaçable dans les tests.</summary>
        internal static Func<HttpMessageHandler> HttpHandlerFactory = CreateHttpHandler;

        /// <summary>Crée un transport HTTP qui refuse les redirections automatiques.</summary>
        /// <returns>Gestionnaire configuré pour ne pas suivre les redirections.</returns>
        private static HttpMessageHandler CreateHttpHandler() { return new HttpClientHandler { AllowAutoRedirect = false }; }

        /// <summary>Crée un client pour le fournisseur, ses réglages et le modèle choisi.</summary>
        /// <param name="provider">Fournisseur LLM sélectionné.</param>
        /// <param name="settings">Paramètres contenant les adresses, clés et options d’authentification.</param>
        /// <param name="selectedModel">Identifiant du modèle demandé.</param>
        public LlmChatClient(LlmProvider provider, LlmSettings settings, string selectedModel) : this(provider, settings, selectedModel, null) { }

        /// <summary>Crée un client et permet d’injecter le transport HTTP.</summary>
        /// <param name="provider">Fournisseur LLM sélectionné.</param>
        /// <param name="settings">Paramètres contenant les adresses, clés et options d’authentification.</param>
        /// <param name="selectedModel">Identifiant du modèle demandé.</param>
        /// <param name="handler">Gestionnaire HTTP facultatif fourni pour le transport.</param>
        internal LlmChatClient(LlmProvider provider, LlmSettings settings, string selectedModel, HttpMessageHandler handler)
            : this(provider, settings, selectedModel, handler, null) { }

        /// <summary>Creates one immutable transport/model capability snapshot.</summary>
        internal LlmChatClient(LlmProvider provider, LlmSettings settings, string selectedModel, HttpMessageHandler handler,
            LlmModelCapabilities capabilities)
        {
            if (provider == null || !provider.Available || provider.IsCodex || settings == null)
                throw new InvalidOperationException(UiText.Get("This provider is not implemented yet."));
            this.provider = provider;
            modelCapabilities = capabilities ?? LlmModelCapabilities.Unknown;
            model = selectedModel;
            if (string.IsNullOrWhiteSpace(model)) throw new InvalidOperationException(UiText.Get("Select a model in the conversation."));
            if (provider.IsOllama)
            {
                ollamaTemperature = settings.OllamaTemperature;
                ollamaTopP = settings.OllamaTopP;
                RequireOllamaSampling(ollamaTemperature, 0, 2, false, nameof(LlmSettings.OllamaTemperature));
                RequireOllamaSampling(ollamaTopP, 0, 1, true, nameof(LlmSettings.OllamaTopP));
            }
            if (provider.IsCopilot) { copilot = new CopilotClient(); return; }
            connection = LlmConnectionProfile.ForModel(provider, settings, model, modelCapabilities);
            endpoint = connection.Endpoint;
            http = handler == null ? new HttpClient(HttpHandlerFactory()) : new HttpClient(handler);
            http.Timeout = TimeSpan.FromSeconds(120);
        }

        /// <summary>Refuses nonfinite or out-of-range optional sampling before constructing an HTTP transport.</summary>
        /// <param name="configured">Optional user value; null leaves the provider default unchanged.</param>
        /// <param name="minimum">Lower bound, excluded only when <paramref name="excludeMinimum"/> is true.</param>
        /// <param name="maximum">Inclusive upper bound.</param>
        /// <param name="excludeMinimum">Whether the lower bound itself is invalid, as for top-p.</param>
        /// <param name="name">Setting name included in the validation error.</param>
        /// <exception cref="InvalidOperationException">The configured value is nonfinite or outside the supported range.</exception>
        private static void RequireOllamaSampling(double? configured, double minimum, double maximum, bool excludeMinimum, string name)
        {
            if (configured.HasValue && (Double.IsNaN(configured.Value) || Double.IsInfinity(configured.Value) ||
                configured.Value < minimum || configured.Value > maximum || (excludeMinimum && configured.Value == minimum)))
                throw new InvalidOperationException(name + " must be finite and within the supported Ollama sampling range.");
        }

        /// <summary>Nom du modèle accompagné de l’hôte ou du fournisseur.</summary>
        /// <value>Identifiant du modèle suivi de GitHub Copilot ou de l’hôte du point de terminaison.</value>
        public string DisplayName { get { return model + " @ " + (provider.IsCopilot ? "GitHub Copilot" : endpoint.Host); } }

        /// <summary>Liste les modèles accessibles ou les identifiants manuels configurés.</summary>
        /// <param name="provider">Fournisseur LLM sélectionné.</param>
        /// <param name="settings">Paramètres contenant les adresses, clés et options d’authentification.</param>
        /// <returns>Options de modèles disponibles pour ce fournisseur.</returns>
        public static async Task<LlmModelOption[]> ListModelsAsync(LlmProvider provider, LlmSettings settings)
        { return await ListModelsAsync(provider, settings, null); }

        /// <summary>Liste les modèles accessibles ou les identifiants manuels configurés.</summary>
        /// <param name="provider">Fournisseur LLM sélectionné.</param>
        /// <param name="settings">Paramètres contenant les adresses, clés et options d’authentification.</param>
        /// <param name="handler">Gestionnaire HTTP facultatif fourni pour le transport.</param>
        /// <returns>Options de modèles disponibles pour ce fournisseur.</returns>
        internal static async Task<LlmModelOption[]> ListModelsAsync(LlmProvider provider, LlmSettings settings, HttpMessageHandler handler)
        {
            if (provider == null || !provider.Available || provider.IsCodex)
                throw new InvalidOperationException(UiText.Get("Model list unavailable for this provider."));
            if (provider.IsCopilot) { using (var client = new CopilotClient()) return await client.ListModelsAsync(); }
            if (provider.ManualModels)
            {
                var ids = (settings.GetManualModels(provider) ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0).Distinct(StringComparer.Ordinal).ToArray();
                if (ids.Length == 0) throw new InvalidOperationException(UiText.Get("Enter the ") + (provider.IsAzure ? UiText.Get("Azure deployments") : UiText.Get("model identifiers")) + UiText.Get(" in provider settings."));
                return ids.Select(id => new LlmModelOption(id, id)).ToArray();
            }
            var connection = LlmConnectionProfile.ForCatalogue(provider, settings);
            Uri catalogue = connection.CatalogueEndpoint();
            Uri catalogueBase = catalogue;
            var models = new List<LlmModelOption>();
            var cursors = new HashSet<string>();
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20)))
            using (var client = new HttpClient(handler ?? HttpHandlerFactory()) { Timeout = TimeSpan.FromSeconds(20) })
                for (int page = 0; page < 100; page++)
                {
                    using (var request = new HttpRequestMessage(HttpMethod.Get, catalogue))
                    {
                        connection.Authenticate(request);
                        using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(false))
                        {
                            if (!response.IsSuccessStatusCode)
                                throw new InvalidOperationException("Catalogue HTTP " + (int)response.StatusCode + ": " +
                                    UiText.Get("Check credentials, URL and provider limits."));
                            string body = await ChatStreamReader.ReadBodyAsync(await response.Content.ReadAsStreamAsync().ConfigureAwait(false), timeout.Token).ConfigureAwait(false);
                            var root = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.DeserializeObject(body) as IDictionary<string, object>;
                            var data = (root != null && root.TryGetValue(provider.IsOllama ? "models" : "data", out object raw)
                                ? raw as object[] : null) ?? throw new InvalidOperationException(UiText.Get("The provider did not return a model list."));
                            foreach (var entry in data)
                            {
                                if (!(entry is IDictionary<string, object> item)) continue;
                                string field = provider.IsOllama ? "name" : "id";
                                string id = item.ContainsKey(field) ? Convert.ToString(item[field]) : null;
                                if (!string.IsNullOrWhiteSpace(id)) models.Add(new LlmModelOption(id, ClaudeProtocol.Text(item, "display_name") ?? ClaudeProtocol.Text(item, "name") ?? id, capabilities: LlmModelCapabilities.FromCatalogue(provider, item)));
                            }
                            if (!provider.IsClaude || !root.ContainsKey("has_more") || !Equals(root["has_more"], true)) return models.ToArray();
                            string cursor = ClaudeProtocol.Text(root, "last_id");
                            if (string.IsNullOrWhiteSpace(cursor) || !cursors.Add(cursor)) throw new InvalidOperationException(UiText.Get("Invalid Claude model list pagination."));
                            var next = new UriBuilder(catalogueBase)
                            {
                                Query = connection.CatalogueQuery(cursor)
                            }; catalogue = next.Uri;
                        }
                    }
                }
            throw new InvalidOperationException(UiText.Get("The model list exceeds the pagination limit."));
        }

        /// <summary>Envoie la conversation au fournisseur, gère le flux de texte et normalise sa réponse.</summary>
        /// <param name="messages">Historique des messages de la conversation.</param>
        /// <param name="tools">Définitions des outils disponibles pour le modèle.</param>
        /// <returns>Message JSON normalisé produit par le fournisseur, avec les appels d’outil éventuels.</returns>
        public async Task<IDictionary<string, object>> CompleteAsync(IList<object> messages, object[] tools)
        {
            if (modelCapabilities.ToolCalling == false) tools = new object[0];
            LastStreamDiagnostics = null;
            if (copilot != null) { copilot.TextDelta = TextDelta; copilot.ActivityUpdate = ActivityUpdate; return RequireModelTools(await copilot.CompleteAsync(model, messages, tools, modelCapabilities.ToolCalling == false ? null : ToolHandler)); }
            object payload;
            bool streaming = (TextDelta != null || ActivityUpdate != null) && !provider.IsBedrock;
            string activityPrefix = Guid.NewGuid().ToString("N");
            if (provider.IsBedrock) payload = BedrockProtocol.Request(messages, tools);
            else if (provider.IsClaude) { var data = ClaudeProtocol.Object(ClaudeProtocol.Request(model, messages, tools)); if (streaming) data["stream"] = true; payload = data; }
            else
            {
                var data = new Dictionary<string, object> { ["model"] = model, ["messages"] = messages, ["tools"] = tools };
                if (!provider.Local) data["tool_choice"] = "auto";
                if (provider.Name == "OpenAI API") data["store"] = false;
                if (streaming) data["stream"] = true;
                if (provider.IsOllama)
                {
                    if (ollamaTemperature.HasValue) data["temperature"] = ollamaTemperature.Value;
                    if (ollamaTopP.HasValue) data["top_p"] = ollamaTopP.Value;
                }
                payload = data;
            }
            if (modelCapabilities.ToolCalling == false)
            {
                var textOnlyPayload = ClaudeProtocol.Object(payload);
                textOnlyPayload.Remove("tools");
                textOnlyPayload.Remove("tool_choice");
                textOnlyPayload.Remove("toolConfig");
                payload = textOnlyPayload;
            }
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(lifetime.Token))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(120));
                request.Content = new StringContent(json.Serialize(payload), Encoding.UTF8, "application/json");
                connection.Authenticate(request);
                using (var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException(provider.Name + " HTTP " + (int)response.StatusCode + UiText.Get(": check the key, model, URL and provider limits."));
                    if (streaming && response.Content.Headers.ContentType?.MediaType == "text/event-stream")
                    {
                        LastStreamDiagnostics = new StreamDiagnostics();
                        return RequireModelTools(await ChatStreamReader.ReadAsync(await response.Content.ReadAsStreamAsync(), provider.IsClaude, TextDelta, timeout.Token, LastStreamDiagnostics, ActivityUpdate, activityPrefix));
                    }
                    string body = await ChatStreamReader.ReadBodyAsync(await response.Content.ReadAsStreamAsync(), timeout.Token);

                    var root = json.DeserializeObject(body) as IDictionary<string, object>;
                    if (provider.IsBedrock) return RequireModelTools(BedrockProtocol.Response(root, ActivityUpdate, activityPrefix));
                    if (provider.IsClaude) return RequireModelTools(ClaudeProtocol.Response(root, ActivityUpdate, activityPrefix));
                    // JavaScriptSerializer uses object[] for JSON arrays.
                    var array = root != null && root.ContainsKey("choices") ? root["choices"] as object[] : null;
                    if (array == null || array.Length == 0)
                        throw new InvalidOperationException("The LLM returned no choice.");
                    var choice = array[0] as IDictionary<string, object>;
                    var reasoning = new ProviderReasoning(ActivityUpdate, activityPrefix);
                    reasoning.OpenAi(choice != null && choice.TryGetValue("message", out var rawReasoningMessage) ? rawReasoningMessage as IDictionary<string, object> : null, true);
                    try
                    {
                        string finish = ClaudeProtocol.Text(choice, "finish_reason");
                        if (finish == "length" || finish == "content_filter") throw new InvalidOperationException(UiText.Get("Response truncated or filtered; tool calls were not executed."));
                        if (choice == null || !choice.ContainsKey("message"))
                            throw new InvalidOperationException("The LLM response has no message.");
                        if (!(choice["message"] is IDictionary<string, object> message)) throw new InvalidOperationException("Invalid LLM message.");
                        if (message.TryGetValue("tool_calls", out var proposedCalls) && proposedCalls is object[] calls && calls.Length > 0 &&
                            finish != "tool_calls" && finish != "stop")
                            throw new InvalidOperationException(UiText.Get("Response truncated or filtered; tool calls were not executed."));
                        RequireModelTools(message);
                        reasoning.Finish("completed");
                        return message;
                    }
                    catch
                    {
                        reasoning.Finish(timeout.IsCancellationRequested ? "interrupted" : "failed");
                        throw;
                    }
                }
            }
        }

        /// <summary>Rejects tool calls contradicting an explicit model declaration before any native dispatch.</summary>
        private IDictionary<string, object> RequireModelTools(IDictionary<string, object> message)
        {
            if (modelCapabilities.ToolCalling == false && message != null &&
                message.TryGetValue("tool_calls", out var raw) && raw is object[] calls && calls.Length != 0)
                throw new InvalidOperationException(UiText.Get("This model does not support tool calling."));
            return message;
        }

        /// <summary>Annule la requête active et libère les clients détenus.</summary>
        public void Dispose() { if (disposed) return; disposed = true; lifetime.Cancel(); http?.Dispose(); copilot?.Dispose(); lifetime.Dispose(); }

    }
}
