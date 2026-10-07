using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Net.Http;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmConnectionProfileTests
    {
        private static LlmSettings Settings(LlmProvider provider, string endpoint = "https://fixture.invalid/v1/chat/completions")
        {
            var settings = new LlmSettings();
            settings.SetEndpoint(provider, endpoint);
            if (provider.RequiresKey) settings.SetKey(provider, "fixture-key");
            return settings;
        }

        [TestMethod]
        public void EveryProviderBindsItsProtocolAndAuthenticationWithoutNativeActivation()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (var provider in LlmProvider.All)
                {
                    var connection = LlmConnectionProfile.ForModel(provider, Settings(provider), "fixture-model");
                    Assert.AreSame(provider, connection.Provider); Assert.AreEqual("fixture-model", connection.Model);
                    var protocol = provider.IsCodex ? LlmConnectionProtocol.CodexAppServer : provider.IsCopilot ? LlmConnectionProtocol.CopilotSdk :
                        provider.IsClaude ? LlmConnectionProtocol.AnthropicMessages : provider.IsBedrock ? LlmConnectionProtocol.BedrockConverse : LlmConnectionProtocol.OpenAiCompletions;
                    var auth = provider.IsCodex || provider.IsCopilot ? LlmConnectionAuthentication.ManagedAccount :
                        provider.IsClaude ? LlmConnectionAuthentication.AnthropicApiKey : provider.IsAzure ? LlmConnectionAuthentication.AzureApiKey : LlmConnectionAuthentication.Bearer;
                    Assert.AreEqual(protocol, connection.Protocol, provider.Name); Assert.AreEqual(auth, connection.Authentication, provider.Name);
                    Assert.AreEqual(provider.IsCodex || provider.IsCopilot, connection.Endpoint == null, provider.Name);
                }
        }

        [TestMethod]
        public void ConnectionCapturesEndpointAndCredentialInsteadOfRereadingMutableSettings()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Mistral"); var settings = Settings(provider);
                var connection = LlmConnectionProfile.ForModel(provider, settings, "first-model");
                settings.SetKey(provider, "replacement-key"); settings.SetEndpoint(provider, "https://other.invalid/v1/chat/completions");
                using (var request = new HttpRequestMessage(HttpMethod.Get, connection.CatalogueEndpoint()))
                {
                    connection.Authenticate(request);
                    Assert.AreEqual("Bearer", request.Headers.Authorization.Scheme); Assert.AreEqual("fixture-key", request.Headers.Authorization.Parameter);
                    Assert.AreEqual("fixture.invalid", request.RequestUri.Host); Assert.AreEqual("first-model", connection.Model);
                }
            }
        }

        [TestMethod]
        public void AzureEntraCapturesTheSelectedEnvironmentTokenAtAdmission()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Azure OpenAI"); var settings = Settings(provider);
                settings.SetKey(provider, null); settings.AzureUseEntraToken = true;
                Environment.SetEnvironmentVariable("AZURE_OPENAI_ENTRA_TOKEN", "first-token");
                var connection = LlmConnectionProfile.ForModel(provider, settings, "deployment");
                Environment.SetEnvironmentVariable("AZURE_OPENAI_ENTRA_TOKEN", "replacement-token"); settings.AzureUseEntraToken = false;
                using (var request = new HttpRequestMessage(HttpMethod.Post, connection.Endpoint))
                {
                    connection.Authenticate(request);
                    Assert.AreEqual(LlmConnectionAuthentication.Bearer, connection.Authentication);
                    Assert.AreEqual("first-token", request.Headers.Authorization.Parameter); Assert.IsFalse(request.Headers.Contains("api-key"));
                }
            }
        }

        [DataTestMethod]
        [DataRow("https://fixture.invalid/v1/chat/completions?api-version=v1")]
        [DataRow("http://localhost:1234/v1/chat/completions")]
        [DataRow("http://127.0.0.1:1234/v1/chat/completions")]
        [DataRow("http://[::1]:1234/v1/chat/completions")]
        public void EndpointValidationAcceptsHttpsAndExplicitHttpLoopback(string endpoint)
        {
            Assert.AreEqual(new Uri(endpoint), LlmConnectionProfile.RequireHttpEndpoint(endpoint));
        }

        [DataTestMethod]
        [DataRow(null), DataRow(""), DataRow("/relative/chat/completions")]
        [DataRow("ftp://localhost/path"), DataRow("http://remote.invalid/v1/chat/completions")]
        [DataRow("https://user:password@fixture.invalid/v1/chat/completions")]
        [DataRow("https://fixture.invalid/v1/chat/completions#section")]
        public void EndpointValidationRejectsUnsafeAddressBeforeCredentialsAreApplied(string endpoint)
        {
            Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.RequireHttpEndpoint(endpoint));
        }

        [DataTestMethod]
        [DataRow("https://other.invalid/v1/models"), DataRow("https://fixture.invalid:444/v1/models")]
        [DataRow("http://fixture.invalid/v1/models"), DataRow("https://user@fixture.invalid/v1/models")]
        [DataRow("https://fixture.invalid/v1/models#section"), DataRow("/v1/models")]
        public void AuthenticationRejectsAnotherOriginOrUnreviewedUrlWithoutAddingHeaders(string address)
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Mistral"); var connection = LlmConnectionProfile.ForModel(provider, Settings(provider), "model");
                using (var request = new HttpRequestMessage(HttpMethod.Get, new Uri(address, UriKind.RelativeOrAbsolute)))
                {
                    Assert.ThrowsException<InvalidOperationException>(() => connection.Authenticate(request));
                    Assert.IsNull(request.Headers.Authorization); Assert.AreEqual(0, request.Headers.Count());
                }
                Assert.ThrowsException<InvalidOperationException>(() => connection.Authenticate(null));
            }
        }

        [TestMethod]
        public void CatalogueChangesOnlyRecognizedPathAndPreservesOriginAndEncodedQuery()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (string name in new[] { "Mistral", "Claude", "Ollama" })
                {
                    var provider = LlmBoundaryScope.Provider(name);
                    string path = provider.IsClaude ? "/proxy/v1/messages" : "/proxy/v1/chat/completions";
                    var connection = LlmConnectionProfile.ForCatalogue(provider, Settings(provider, "https://fixture.invalid:443" + path + "?scope=chat%2Fcompletions&target=%2Fmessages"));
                    var catalogue = connection.CatalogueEndpoint();
                    Assert.AreEqual(provider.IsOllama ? "/api/tags" : "/proxy/v1/models", catalogue.AbsolutePath);
                    Assert.AreEqual("https", catalogue.Scheme); Assert.AreEqual("fixture.invalid", catalogue.Host); Assert.AreEqual(443, catalogue.Port);
                    Assert.AreEqual(connection.Endpoint.Query, catalogue.Query);
                }
        }

        [TestMethod]
        public void CatalogueQueryReplacesOnlyCursorAndEscapesItsValue()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Claude");
                var connection = LlmConnectionProfile.ForCatalogue(provider, Settings(provider, "https://fixture.invalid/v1/messages?api-version=v1&after_id=old&%61fter_id=old2&scope=a%2Fb&empty="));
                Assert.AreEqual(connection.CatalogueEndpoint().Query.TrimStart('?'), connection.CatalogueQuery(null));
                string query = connection.CatalogueQuery("a b/#&=+?");
                Assert.AreEqual("api-version=v1&scope=a%2Fb&empty=&after_id=" + Uri.EscapeDataString("a b/#&=+?"), query);
                var paged = new UriBuilder(connection.CatalogueEndpoint()) { Query = query }.Uri;
                Assert.AreEqual("fixture.invalid", paged.Host); Assert.AreEqual("/v1/models", paged.AbsolutePath); Assert.AreEqual("", paged.Fragment);
            }
        }

        [TestMethod]
        public void UnrecognizedCataloguePathCannotBeRecoveredFromQueryContent()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Mistral");
                var connection = LlmConnectionProfile.ForCatalogue(provider, Settings(provider, "https://fixture.invalid/other?next=/chat/completions"));
                Assert.ThrowsException<InvalidOperationException>(() => connection.CatalogueEndpoint());
                foreach (var manual in LlmProvider.All.Where(p => p.ManualModels || p.IsCodex || p.IsCopilot))
                {
                    var manualConnection = LlmConnectionProfile.ForCatalogue(manual, Settings(manual));
                    Assert.ThrowsException<InvalidOperationException>(() => manualConnection.CatalogueEndpoint());
                }
            }
        }

        [TestMethod]
        public void BedrockEscapesModelAsOnePathSegmentAndKeepsConfiguredQuery()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var provider = LlmBoundaryScope.Provider("Amazon Bedrock");
                string model = "arn:aws:bedrock:eu-west-1:123:inference-profile/model /#?";
                var connection = LlmConnectionProfile.ForModel(provider, Settings(provider, "https://fixture.invalid/runtime/?trace=enabled"), model);
                StringAssert.Contains(connection.Endpoint.AbsoluteUri, "/runtime/model/" + Uri.EscapeDataString(model) + "/converse");
                Assert.AreEqual("?trace=enabled", connection.Endpoint.Query); Assert.AreEqual("", connection.Endpoint.Fragment);
            }
        }

        [TestMethod]
        public void MissingProviderModelEndpointOrRequiredKeyFailsBeforeTransport()
        {
            using (var scope = new LlmBoundaryScope())
            {
                Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.ForModel(null, new LlmSettings(), "model"));
                var provider = LlmBoundaryScope.Provider("Mistral");
                Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.ForModel(provider, null, "model"));
                foreach (string model in new[] { null, "", " " })
                    Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.ForModel(provider, Settings(provider), model));
                Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.ForModel(provider, new LlmSettings(), "model"));
                Assert.ThrowsException<InvalidOperationException>(() => LlmConnectionProfile.ForCatalogue(LlmBoundaryScope.Provider("Azure OpenAI"), new LlmSettings()));
                foreach (string name in new[] { "LM Studio", "Ollama", "Personnalisé (OpenAI)" })
                {
                    var optional = LlmBoundaryScope.Provider(name); var settings = Settings(optional);
                    settings.SetKey(optional, null); var connection = LlmConnectionProfile.ForModel(optional, settings, "model");
                    using (var request = new HttpRequestMessage(HttpMethod.Post, connection.Endpoint))
                    {
                        connection.Authenticate(request); Assert.IsNull(request.Headers.Authorization);
                    }
                }
            }
        }
    }
}
