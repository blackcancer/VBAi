using System;
using System.Net.Http;
using System.Net.Http.Headers;

namespace VBAi
{
    internal enum LlmConnectionProtocol { CodexAppServer, CopilotSdk, OpenAiCompletions, AnthropicMessages, BedrockConverse }
    internal enum LlmConnectionAuthentication { ManagedAccount, Bearer, AnthropicApiKey, AzureApiKey }

    /// <summary>Binds one provider, protocol, endpoint and credential snapshot to a model connection.</summary>
    internal sealed class LlmConnectionProfile
    {
        private readonly string credential;
        internal LlmProvider Provider { get; }
        internal LlmConnectionProtocol Protocol { get; }
        internal LlmConnectionAuthentication Authentication { get; }
        internal string Model { get; }
        internal Uri Endpoint { get; }
        internal LlmModelCapabilities Capabilities { get; }

        private LlmConnectionProfile(LlmProvider provider, LlmSettings settings, string model, bool requireModel, LlmModelCapabilities capabilities = null)
        {
            if (provider == null || !provider.Available || settings == null)
                throw new InvalidOperationException(UiText.Get("This provider is not implemented yet."));
            if (requireModel && string.IsNullOrWhiteSpace(model))
                throw new InvalidOperationException(UiText.Get("Select a model in the conversation."));
            Provider = provider;
            Model = model;
            Capabilities = capabilities ?? LlmModelCapabilities.Unknown;
            Protocol = provider.IsCodex ? LlmConnectionProtocol.CodexAppServer :
                provider.IsCopilot ? LlmConnectionProtocol.CopilotSdk :
                provider.IsClaude ? LlmConnectionProtocol.AnthropicMessages :
                provider.IsBedrock ? LlmConnectionProtocol.BedrockConverse : LlmConnectionProtocol.OpenAiCompletions;
            Authentication = provider.IsCodex || provider.IsCopilot ? LlmConnectionAuthentication.ManagedAccount :
                provider.IsClaude ? LlmConnectionAuthentication.AnthropicApiKey :
                provider.IsAzure && !settings.AzureUseEntraToken ? LlmConnectionAuthentication.AzureApiKey : LlmConnectionAuthentication.Bearer;
            if (Authentication == LlmConnectionAuthentication.ManagedAccount) return;
            string raw = settings.ResolveEndpoint(provider);
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidOperationException(UiText.Get("Configure ") + provider.Name + UiText.Get(" in settings."));
            Endpoint = RequireHttpEndpoint(raw);
            if (provider.IsBedrock && requireModel)
            {
                var builder = new UriBuilder(Endpoint);
                builder.Path = builder.Path.TrimEnd('/') + "/model/" + Uri.EscapeDataString(model) + "/converse";
                Endpoint = builder.Uri;
            }
            credential = settings.GetKey(provider);
            if (string.IsNullOrWhiteSpace(credential) && provider.RequiresKey)
                throw new InvalidOperationException(UiText.Get("Configure the API key for ") + provider.Name + UiText.Get(" in VBAi settings."));
        }

        internal static LlmConnectionProfile ForModel(LlmProvider provider, LlmSettings settings, string model, LlmModelCapabilities capabilities = null) =>
            new LlmConnectionProfile(provider, settings, model, true, capabilities);

        internal static LlmConnectionProfile ForCatalogue(LlmProvider provider, LlmSettings settings) =>
            new LlmConnectionProfile(provider, settings, null, false);

        /// <summary>Validates the transport address without treating an URL as an authentication store.</summary>
        internal static Uri RequireHttpEndpoint(string raw)
        {
            if (!Uri.TryCreate(raw, UriKind.Absolute, out var endpoint) ||
                (endpoint.Scheme != Uri.UriSchemeHttps && !(endpoint.Scheme == Uri.UriSchemeHttp && endpoint.IsLoopback)) ||
                !string.IsNullOrEmpty(endpoint.UserInfo) || !string.IsNullOrEmpty(endpoint.Fragment))
                throw new InvalidOperationException(UiText.Get("The LLM endpoint must use HTTPS (or HTTP on localhost), without embedded credentials or a fragment."));
            return endpoint;
        }

        /// <summary>Derives only the documented catalog path, preserving the configured origin and query.</summary>
        internal Uri CatalogueEndpoint()
        {
            if (Endpoint == null || Provider.ManualModels)
                throw new InvalidOperationException(UiText.Get("Model list unavailable for this provider."));
            var builder = new UriBuilder(Endpoint);
            if (Provider.IsOllama) builder.Path = "/api/tags";
            else
            {
                string suffix = Provider.IsClaude ? "/messages" : "/chat/completions";
                if (!builder.Path.EndsWith(suffix, StringComparison.Ordinal))
                    throw new InvalidOperationException(UiText.Get("The configured chat endpoint has no recognized model catalog path."));
                builder.Path = builder.Path.Substring(0, builder.Path.Length - suffix.Length) + "/models";
            }
            return builder.Uri;
        }

        /// <summary>Preserves configured catalog query fields while replacing an explicit pagination cursor.</summary>
        /// <param name="cursor">New cursor, or null to retain the initial configured query.</param>
        /// <returns>Escaped query without its leading question mark.</returns>
        internal string CatalogueQuery(string cursor)
        {
            string query = CatalogueEndpoint().Query.TrimStart('?');
            if (cursor == null) return query;
            var fields = new System.Collections.Generic.List<string>();
            foreach (string field in query.Split('&'))
            {
                if (field.Length == 0) continue;
                int separator = field.IndexOf('=');
                string key = separator < 0 ? field : field.Substring(0, separator);
                if (Uri.UnescapeDataString(key.Replace("+", " ")) != "after_id") fields.Add(field);
            }
            fields.Add("after_id=" + Uri.EscapeDataString(cursor));
            return string.Join("&", fields);
        }

        /// <summary>Applies the captured credential only to the reviewed connection origin.</summary>
        internal void Authenticate(HttpRequestMessage request)
        {
            if (request == null || Endpoint == null || request.RequestUri == null ||
                !request.RequestUri.IsAbsoluteUri || request.RequestUri.Scheme != Endpoint.Scheme ||
                !string.Equals(request.RequestUri.IdnHost, Endpoint.IdnHost, StringComparison.OrdinalIgnoreCase) ||
                request.RequestUri.Port != Endpoint.Port || !string.IsNullOrEmpty(request.RequestUri.UserInfo) ||
                !string.IsNullOrEmpty(request.RequestUri.Fragment))
                throw new InvalidOperationException(UiText.Get("The request does not belong to its model connection."));
            if (Authentication == LlmConnectionAuthentication.AnthropicApiKey)
            {
                request.Headers.Add("anthropic-version", "2023-06-01");
                if (!string.IsNullOrWhiteSpace(credential)) request.Headers.Add("x-api-key", credential);
            }
            else if (Authentication == LlmConnectionAuthentication.AzureApiKey)
            {
                if (!string.IsNullOrWhiteSpace(credential)) request.Headers.Add("api-key", credential);
            }
            else if (!string.IsNullOrWhiteSpace(credential)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", credential);
        }
    }
}
