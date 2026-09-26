using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // This client speaks the Chat Completions function-call format. No credential is persisted.
    internal sealed class LlmChatClient : IDisposable
    {
        private readonly HttpClient http = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private readonly Uri endpoint;
        private readonly string model;
        private readonly string key;
        private readonly LlmProvider provider;

        public LlmChatClient(LlmProvider provider, LlmSettings settings)
        {
            if (provider == null || !provider.Available || provider.IsCodex || settings == null)
                throw new InvalidOperationException("Ce fournisseur n'est pas encore implémenté.");
            this.provider = provider;
            string raw = settings.ResolveEndpoint(provider);
            endpoint = new Uri(raw, UriKind.Absolute);
            if (endpoint.Scheme != Uri.UriSchemeHttps &&
                !(endpoint.Scheme == Uri.UriSchemeHttp && endpoint.IsLoopback))
                throw new InvalidOperationException("The LLM endpoint must use HTTPS (or HTTP on localhost).");
            model = settings.ResolveModel(provider);
            key = provider.Local ? null : settings.GetOpenAiKey();
            if (string.IsNullOrWhiteSpace(key) && !provider.Local)
                throw new InvalidOperationException("Configurez la clé OpenAI API dans les paramètres CodexVBE.");
        }

        public string DisplayName { get { return model + " @ " + endpoint.Host; } }

        public async Task<IDictionary<string, object>> CompleteAsync(IList<object> messages, object[] tools)
        {
            object payload = provider.Local
                ? (object)new { model, messages, tools }
                : new { model, messages, tools, tool_choice = "auto", store = false };
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(json.Serialize(payload), Encoding.UTF8, "application/json");
                if (!string.IsNullOrWhiteSpace(key)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);
                using (var response = await http.SendAsync(request).ConfigureAwait(true))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException("LLM HTTP " + (int)response.StatusCode + ": " + body.Substring(0, Math.Min(body.Length, 500)));
                    var root = json.DeserializeObject(body) as IDictionary<string, object>;
                    // JavaScriptSerializer uses object[] for JSON arrays.
                    var array = root != null && root.ContainsKey("choices") ? root["choices"] as object[] : null;
                    if (array == null || array.Length == 0)
                        throw new InvalidOperationException("The LLM returned no choice.");
                    var choice = array[0] as IDictionary<string, object>;
                    if (choice == null || !choice.ContainsKey("message"))
                        throw new InvalidOperationException("The LLM response has no message.");
                    var message = choice["message"] as IDictionary<string, object>;
                    if (message == null) throw new InvalidOperationException("Invalid LLM message.");
                    return message;
                }
            }
        }

        public void Dispose() { http.Dispose(); }
    }
}
