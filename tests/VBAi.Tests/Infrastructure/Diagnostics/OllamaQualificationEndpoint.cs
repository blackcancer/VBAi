using System;

namespace VBAi.Tests.Integration
{
    /// <summary>Allows an explicitly selected local test port without expanding the synthetic wire boundary.</summary>
    internal static class OllamaQualificationEndpoint
    {
        internal const string EnvironmentName = "VBAi_TEST_OLLAMA_ENDPOINT";
        internal const string Default = "http://127.0.0.1:11434/v1/chat/completions";

        internal static Uri Resolve() => Parse(Environment.GetEnvironmentVariable(EnvironmentName));

        internal static Uri Parse(string configured)
        {
            string text = configured ?? Default;
            Uri uri;
            if (!Uri.TryCreate(text, UriKind.Absolute, out uri) || uri.Scheme != "http" ||
                uri.Host != "127.0.0.1" || uri.Port <= 0 || uri.UserInfo.Length != 0 ||
                uri.Query.Length != 0 || uri.Fragment.Length != 0 ||
                uri.AbsolutePath != "/v1/chat/completions" || text != uri.AbsoluteUri)
                throw new InvalidOperationException("The qualification endpoint must be a canonical literal IPv4 loopback chat-completions URL, without credentials, query or fragment.");
            return uri;
        }

        internal static void RequireWireUri(Uri request, Uri selected)
        {
            if (request == null || selected == null || !request.IsAbsoluteUri ||
                request.Scheme != selected.Scheme || request.Host != selected.Host ||
                request.Port != selected.Port || request.UserInfo.Length != 0 ||
                request.Query.Length != 0 || request.Fragment.Length != 0 ||
                (request.AbsolutePath != "/api/tags" && request.AbsolutePath != selected.AbsolutePath))
                throw new InvalidOperationException("Synthetic wire capture requires the exact selected loopback server and the catalogue/chat routes only.");
        }
    }
}
