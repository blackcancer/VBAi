using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // This client speaks the Chat Completions function-call format. No credential is persisted.
    internal sealed class LlmChatClient : IDisposable
    {
        private readonly HttpClient http;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private readonly Uri endpoint;
        private readonly string model;
        private readonly string key;
        private readonly LlmProvider provider;
        private readonly CopilotClient copilot;
        private readonly bool azureEntra;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private bool disposed;
        public Action<string> TextDelta { get; set; }
        public Func<string, string, Task<string>> ToolHandler { get; set; }
        internal static Func<HttpMessageHandler> HttpHandlerFactory = CreateHttpHandler;
        private static HttpMessageHandler CreateHttpHandler() { return new HttpClientHandler { AllowAutoRedirect = false }; }

        public LlmChatClient(LlmProvider provider, LlmSettings settings, string selectedModel) : this(provider, settings, selectedModel, null) { }

        internal LlmChatClient(LlmProvider provider, LlmSettings settings, string selectedModel, HttpMessageHandler handler)
        {
            if (provider == null || !provider.Available || provider.IsCodex || settings == null)
                throw new InvalidOperationException(UiText.Get("This provider is not implemented yet."));
            this.provider = provider;
            azureEntra = settings.AzureUseEntraToken;
            model = selectedModel;
            if (string.IsNullOrWhiteSpace(model)) throw new InvalidOperationException(UiText.Get("Select a model in the conversation."));
            if (provider.IsCopilot) { copilot = new CopilotClient(); return; }
            string raw = settings.ResolveEndpoint(provider);
            if (string.IsNullOrWhiteSpace(raw)) throw new InvalidOperationException("Configurez l’URL de " + provider.Name + UiText.Get(" in settings."));
            if (provider.IsBedrock) raw = raw.TrimEnd('/') + "/model/" + Uri.EscapeDataString(model) + "/converse";
            endpoint = new Uri(raw, UriKind.Absolute);
            if (endpoint.Scheme != Uri.UriSchemeHttps &&
                !(endpoint.Scheme == Uri.UriSchemeHttp && endpoint.IsLoopback))
                throw new InvalidOperationException("The LLM endpoint must use HTTPS (or HTTP on localhost).");
            key = settings.GetKey(provider);
            if (string.IsNullOrWhiteSpace(key) && provider.RequiresKey)
                throw new InvalidOperationException(UiText.Get("Configure the API key for ") + provider.Name + UiText.Get(" in VBAi settings."));
            http = handler == null ? new HttpClient(HttpHandlerFactory()) : new HttpClient(handler);
            http.Timeout = TimeSpan.FromSeconds(120);
        }

        public string DisplayName { get { return model + " @ " + (provider.IsCopilot ? "GitHub Copilot" : endpoint.Host); } }

        public static async Task<LlmModelOption[]> ListModelsAsync(LlmProvider provider, LlmSettings settings)
        { return await ListModelsAsync(provider, settings, null); }

        internal static async Task<LlmModelOption[]> ListModelsAsync(LlmProvider provider, LlmSettings settings, HttpMessageHandler handler)
        {
            if (provider == null || !provider.Available || provider.IsCodex)
                throw new InvalidOperationException(UiText.Get("Model list unavailable for this provider."));
            if (provider.IsCopilot) { using (var client = new CopilotClient()) return await client.ListModelsAsync(); }
            if (provider.ManualModels) {
                var ids = (settings.GetManualModels(provider) ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0).Distinct(StringComparer.Ordinal).ToArray();
                if (ids.Length == 0) throw new InvalidOperationException(UiText.Get("Enter the ") + (provider.IsAzure ? UiText.Get("Azure deployments") : UiText.Get("model identifiers")) + UiText.Get(" in provider settings."));
                return ids.Select(id => new LlmModelOption(id, id)).ToArray();
            }
            var chatEndpoint = new Uri(settings.ResolveEndpoint(provider), UriKind.Absolute);
            if (chatEndpoint.Scheme != Uri.UriSchemeHttps &&
                !(chatEndpoint.Scheme == Uri.UriSchemeHttp && chatEndpoint.IsLoopback))
                throw new InvalidOperationException("L'URL doit utiliser HTTPS, ou HTTP sur localhost.");
            // OpenAI-compatible chat endpoints expose /models next to /chat/completions.
            Uri catalogue = provider.IsOllama
                ? new Uri(chatEndpoint.GetLeftPart(UriPartial.Authority) + "/api/tags")
                : new Uri(chatEndpoint.AbsoluteUri.Replace(provider.IsClaude ? "/messages" : "/chat/completions", "/models"));
            var models = new List<LlmModelOption>();
            var cursors = new HashSet<string>();
            using (var client = new HttpClient(handler ?? HttpHandlerFactory()) { Timeout = TimeSpan.FromSeconds(20) })
            for (int page = 0; page < 100; page++)
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, catalogue)) {
                string key = settings.GetKey(provider);
                if (string.IsNullOrWhiteSpace(key) && provider.RequiresKey)
                    throw new InvalidOperationException(UiText.Get("Configure the API key for ") + provider.Name + UiText.Get(" to load models."));
                Authenticate(request, provider, key, settings.AzureUseEntraToken);
                using (var response = await client.SendAsync(request).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException("Catalogue HTTP " + (int)response.StatusCode + ": " +
                            UiText.Get("Check credentials, URL and provider limits."));
                    var root = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.DeserializeObject(body) as IDictionary<string, object>;
                    object raw;
                    var data = root != null && root.TryGetValue(provider.IsOllama ? "models" : "data", out raw)
                        ? raw as object[] : null;
                    if (data == null) throw new InvalidOperationException(UiText.Get("The provider did not return a model list."));
                    foreach (var entry in data)
                    {
                        var item = entry as IDictionary<string, object>;
                        if (item == null) continue;
                        string field = provider.IsOllama ? "name" : "id";
                        string id = item.ContainsKey(field) ? Convert.ToString(item[field]) : null;
                        if (!string.IsNullOrWhiteSpace(id)) models.Add(new LlmModelOption(id, ClaudeProtocol.Text(item, "display_name") ?? ClaudeProtocol.Text(item, "name") ?? id));
                    }
                    if (!provider.IsClaude || !root.ContainsKey("has_more") || !Equals(root["has_more"], true)) return models.ToArray();
                    string cursor = ClaudeProtocol.Text(root, "last_id");
                    if (string.IsNullOrWhiteSpace(cursor) || !cursors.Add(cursor)) throw new InvalidOperationException(UiText.Get("Invalid Claude model list pagination."));
                    var next = new UriBuilder(catalogue); next.Query = "after_id=" + Uri.EscapeDataString(cursor);
                    catalogue = next.Uri;
                }
                }
            }
            throw new InvalidOperationException(UiText.Get("The model list exceeds the pagination limit."));
        }

        public async Task<IDictionary<string, object>> CompleteAsync(IList<object> messages, object[] tools)
        {
            if (copilot != null) { copilot.TextDelta = TextDelta; return await copilot.CompleteAsync(model, messages, tools, ToolHandler); }
            object payload;
            bool streaming = TextDelta != null && !provider.IsBedrock;
            if (provider.IsBedrock) payload = BedrockProtocol.Request(messages, tools);
            else if (provider.IsClaude) { var data = ClaudeProtocol.Object(ClaudeProtocol.Request(model, messages, tools)); if (streaming) data["stream"] = true; payload = data; }
            else {
                var data = new Dictionary<string, object> { ["model"] = model, ["messages"] = messages, ["tools"] = tools };
                if (!provider.Local) data["tool_choice"] = "auto";
                if (provider.Name == "OpenAI API") data["store"] = false;
                if (streaming) data["stream"] = true;
                payload = data;
            }
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(lifetime.Token))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(120));
                request.Content = new StringContent(json.Serialize(payload), Encoding.UTF8, "application/json");
                Authenticate(request, provider, key, azureEntra);
                using (var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException(provider.Name + " HTTP " + (int)response.StatusCode + UiText.Get(": check the key, model, URL and provider limits."));
                    if (streaming && response.Content.Headers.ContentType?.MediaType == "text/event-stream")
                        return await ChatStreamReader.ReadAsync(await response.Content.ReadAsStreamAsync(), provider.IsClaude, TextDelta, timeout.Token);
                    string body;
                    using (var bodyStream = await response.Content.ReadAsStreamAsync())
                    using (timeout.Token.Register(() => bodyStream.Dispose()))
                    using (var reader = new System.IO.StreamReader(bodyStream, Encoding.UTF8)) body = await reader.ReadToEndAsync();
                    timeout.Token.ThrowIfCancellationRequested();
                    var root = json.DeserializeObject(body) as IDictionary<string, object>;
                    if (provider.IsBedrock) return BedrockProtocol.Response(root);
                    if (provider.IsClaude) return ClaudeProtocol.Response(root);
                    // JavaScriptSerializer uses object[] for JSON arrays.
                    var array = root != null && root.ContainsKey("choices") ? root["choices"] as object[] : null;
                    if (array == null || array.Length == 0)
                        throw new InvalidOperationException("The LLM returned no choice.");
                    var choice = array[0] as IDictionary<string, object>;
                    string finish = ClaudeProtocol.Text(choice, "finish_reason");
                    if (finish == "length" || finish == "content_filter") throw new InvalidOperationException(UiText.Get("Response truncated or filtered; tool calls were not executed."));
                    if (choice == null || !choice.ContainsKey("message"))
                        throw new InvalidOperationException("The LLM response has no message.");
                    var message = choice["message"] as IDictionary<string, object>;
                    if (message == null) throw new InvalidOperationException("Invalid LLM message.");
                    return message;
                }
            }
        }

        public void Dispose() { if (disposed) return; disposed = true; lifetime.Cancel(); http?.Dispose(); copilot?.Dispose(); lifetime.Dispose(); }

        private static void Authenticate(HttpRequestMessage request, LlmProvider provider, string key, bool azureEntra)
        {
            if (provider.IsAzure && !azureEntra) { if (!string.IsNullOrWhiteSpace(key)) request.Headers.Add("api-key", key); }
            else if (provider.IsClaude) {
                request.Headers.Add("anthropic-version", "2023-06-01");
                if (!string.IsNullOrWhiteSpace(key)) request.Headers.Add("x-api-key", key);
            }
            else if (!string.IsNullOrWhiteSpace(key)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);
        }
    }
}
