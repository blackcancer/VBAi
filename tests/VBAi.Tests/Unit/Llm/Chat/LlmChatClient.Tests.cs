using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmChatClientCoverageTests
    {
        private const string Answer = "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"answer\"}}]}";
        private static IList<object> History() { return new List<object> { new { role = "system", content = "Read before editing." }, new { role = "user", content = "request" } }; }
        private static object[] Tools() { return new object[] { new { type = "function", function = new { name = "read_module", description = "fixture", parameters = new { type = "object", properties = new { } } } } }; }
        private static LlmSettings Settings(LlmProvider provider)
        {
            var settings = new LlmSettings(); settings.SetEndpoint(provider, provider.IsClaude ? "https://fixture.invalid/v1/messages" : "https://fixture.invalid/v1/chat/completions"); if (provider.RequiresKey) settings.SetKey(provider, "fixture-only-key"); return settings;
        }
        [TestMethod]
        public void ConstructorsValidateProviderSettingsModelEndpointAndKeysBeforeAnyRequest()
        {
            using (var native = LlmChatClient.HttpHandlerFactory()) { Assert.IsFalse(((HttpClientHandler)native).AllowAutoRedirect); }
            using (var scope = new LlmBoundaryScope())
            {
                var local = LlmBoundaryScope.Provider("Ollama"); var unavailable = (LlmProvider)Activator.CreateInstance(typeof(LlmProvider), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "Unavailable", false, false, null, null, null }, null);
                foreach (var provider in new[] { null, LlmProvider.All[0], unavailable }) Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(provider, new LlmSettings(), "model")); Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(local, null, "model")); foreach (var model in new[] { null, "", " " }) Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(local, new LlmSettings(), model));
                var custom = LlmBoundaryScope.Provider("Personnalisé (OpenAI)"); Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(custom, new LlmSettings(), "model"));
                foreach (var endpoint in new[] { "http://remote.invalid/v1", "ftp://localhost/a", "file:///fixture", "relative" }) { var settings = new LlmSettings { OllamaEndpoint = endpoint }; Exception error = null; try { new LlmChatClient(local, settings, "model"); } catch (Exception ex) { error = ex; } Assert.IsNotNull(error, endpoint); }
                Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(LlmBoundaryScope.Provider("OpenAI API"), new LlmSettings(), "model"));
                foreach (var endpoint in new[] { "http://localhost:11434/v1/chat/completions", "http://127.0.0.1:1/v1", "https://fixture.invalid/v1" }) using (var client = new LlmChatClient(local, new LlmSettings { OllamaEndpoint = endpoint }, "model", new LlmHttpFixture(Answer))) { StringAssert.Contains(client.DisplayName, "model @ "); client.Dispose(); client.Dispose(); }
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("GitHub Copilot"), new LlmSettings(), "model")) { Assert.AreEqual("model @ GitHub Copilot", client.DisplayName); client.Dispose(); }
            }
        }
        [TestMethod]
        public async Task EveryHttpProviderUsesItsNativePayloadAndAuthenticationWithOnlyMemoryHandlers()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (var provider in LlmProvider.All.Where(p => !p.IsCodex && !p.IsCopilot))
                {
                    var settings = Settings(provider); string response = provider.IsClaude ? "{\"content\":[{\"type\":\"text\",\"text\":\"answer\"}],\"stop_reason\":\"end_turn\"}" : provider.IsBedrock ? "{\"output\":{\"message\":{\"content\":[{\"text\":\"answer\"}]}},\"stopReason\":\"end_turn\"}" : Answer; var handler = new LlmHttpFixture(response);
                    using (var client = new LlmChatClient(provider, settings, "model", handler)) { Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]); var body = LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(handler.Bodies.Single())); Assert.IsTrue(body.ContainsKey("messages")); if (!provider.IsClaude && !provider.IsBedrock) { Assert.AreEqual(!provider.Local, body.ContainsKey("tool_choice")); if (provider.Name == "OpenAI API") Assert.AreEqual(false, body["store"]); } if (provider.IsAzure) Assert.IsTrue(handler.Headers[0].ContainsKey("api-key")); else if (provider.IsClaude) { Assert.IsTrue(handler.Headers[0].ContainsKey("anthropic-version")); Assert.IsTrue(handler.Headers[0].ContainsKey("x-api-key")); } else if (provider.RequiresKey) StringAssert.StartsWith(handler.Headers[0]["Authorization"], "Bearer "); }
                }
        }
        [TestMethod]
        public async Task StreamingClaudeOpenAiBedrockAndDefaultHandlerFollowNativeResponseContracts()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var provider in new[] { LlmBoundaryScope.Provider("Ollama"), LlmBoundaryScope.Provider("Claude") })
                {
                    string stream = provider.IsClaude ? "data: {\"type\":\"content_block_start\",\"index\":0,\"content_block\":{\"type\":\"text\",\"text\":\"answer\"}}\n\ndata: {\"type\":\"message_delta\",\"delta\":{\"stop_reason\":\"end_turn\"}}\n\ndata: {\"type\":\"message_stop\"}\n\n" : "data: {\"choices\":[{\"delta\":{\"content\":\"answer\"},\"finish_reason\":\"stop\"}]}\n\ndata: [DONE]\n\n";
                    var handler = new LlmHttpFixture(); handler.Replies.Enqueue(new LlmHttpFixture.Reply(stream) { MediaType = "text/event-stream" }); using (var client = new LlmChatClient(provider, Settings(provider), "model", handler)) { var chunks = new List<string>(); client.TextDelta = chunks.Add; Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]); CollectionAssert.AreEqual(new[] { "answer" }, chunks); }
                }
                var bedrock = LlmBoundaryScope.Provider("Amazon Bedrock"); using (var client = new LlmChatClient(bedrock, Settings(bedrock), "profile / ARN", new LlmHttpFixture("{\"output\":{\"message\":{\"content\":[{\"text\":\"answer\"}]}},\"stopReason\":\"end_turn\"}"))) { client.TextDelta = s => Assert.Fail("Bedrock JSON must not emit SSE deltas."); Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]); }
                var factory = new LlmHttpFixture(Answer); LlmChatClient.HttpHandlerFactory = () => factory; using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model")) { Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]); }
                var noType = new LlmHttpFixture(); noType.Replies.Enqueue(new LlmHttpFixture.Reply(Answer) { OmitContentType = true }); using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", noType)) { client.TextDelta = s => Assert.Fail(); Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]); }
                scope.UseCopilot(); using (var client = new LlmChatClient(LlmBoundaryScope.Provider("GitHub Copilot"), new LlmSettings(), "model")) { client.TextDelta = s => { }; client.ToolHandler = (n, a) => Task.FromResult("result"); Assert.AreEqual("answer été", (await client.CompleteAsync(History(), Tools()))["content"]); }
            }
        }
        [TestMethod]
        public async Task MalformedChoicesAndHttpErrorsFailWithoutDisclosingProviderBody()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var body in new[] { "null", "{}", "{\"choices\":[]}", "{\"choices\":null}", "{\"choices\":[7]}", "{\"choices\":[{}]}", "{\"choices\":[{\"message\":null}]}", "{\"choices\":[{\"finish_reason\":\"length\",\"message\":{}}]}", "{\"choices\":[{\"finish_reason\":\"content_filter\",\"message\":{}}]}" }) using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", new LlmHttpFixture(body))) { await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync(History(), Tools())); }
                var failure = new LlmHttpFixture(); failure.Replies.Enqueue(new LlmHttpFixture.Reply("private-token-body") { Status = HttpStatusCode.Unauthorized }); using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", failure)) { var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync(History(), Tools())); StringAssert.Contains(error.Message, "401"); Assert.IsFalse(error.Message.Contains("private-token-body")); }
            }
        }
        [TestMethod]
        public async Task CataloguesCoverManualNativeCopilotPaginationMalformedDataAndLimits()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var unavailable = (LlmProvider)Activator.CreateInstance(typeof(LlmProvider), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { "Unavailable", false, false, null, null, null }, null); foreach (var provider in new[] { null, LlmProvider.All[0], unavailable }) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(provider, new LlmSettings()));
                foreach (var provider in LlmProvider.All.Where(p => p.ManualModels)) { var settings = new LlmSettings(); await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(provider, settings)); settings.ManualModelLists[provider.Name] = " first \r\nsecond\nfirst\n \n"; CollectionAssert.AreEqual(new[] { "first", "second" }, (await LlmChatClient.ListModelsAsync(provider, settings)).Select(x => x.Id).ToArray()); }
                scope.UseCopilot(); Assert.AreEqual(2, (await LlmChatClient.ListModelsAsync(LlmBoundaryScope.Provider("GitHub Copilot"), new LlmSettings())).Length);
                foreach (var provider in new[] { LlmBoundaryScope.Provider("Ollama"), LlmBoundaryScope.Provider("Mistral"), LlmBoundaryScope.Provider("Claude") })
                {
                    string body = new JavaScriptSerializer().Serialize(new Dictionary<string, object> { { provider.IsOllama ? "models" : "data", new object[] { null, new { }, new Dictionary<string, object> { { provider.IsOllama ? "name" : "id", "" } }, new Dictionary<string, object> { { provider.IsOllama ? "name" : "id", "m1" }, { "display_name", "Shown" } }, new Dictionary<string, object> { { provider.IsOllama ? "name" : "id", "m2" }, { provider.IsOllama ? "display_name" : "name", "Named" } }, new Dictionary<string, object> { { provider.IsOllama ? "name" : "id", "m3" } } } } }); var handler = new LlmHttpFixture(body); var models = await LlmChatClient.ListModelsAsync(provider, Settings(provider), handler); Assert.AreEqual(3, models.Length); Assert.AreEqual("Shown", models[0].Label); Assert.AreEqual("Named", models[1].Label); Assert.AreEqual("m3", models[2].Label); StringAssert.Contains(handler.Uris[0].AbsoluteUri, provider.IsOllama ? "/api/tags" : "models");
                }
                var local = LlmBoundaryScope.Provider("Ollama"); foreach (var endpoint in new[] { "ftp://localhost/data", "http://remote.invalid/api" }) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(local, new LlmSettings { OllamaEndpoint = endpoint }, new LlmHttpFixture("{}")));
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(LlmBoundaryScope.Provider("Mistral"), new LlmSettings(), new LlmHttpFixture("{}")));
                foreach (var body in new[] { "null", "{}", "{\"data\":{}}" }) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(LlmBoundaryScope.Provider("Mistral"), Settings(LlmBoundaryScope.Provider("Mistral")), new LlmHttpFixture(body)));
                var failure = new LlmHttpFixture(); failure.Replies.Enqueue(new LlmHttpFixture.Reply("private-body") { Status = HttpStatusCode.Forbidden }); var failed = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(local, new LlmSettings(), failure)); Assert.IsFalse(failed.Message.Contains("private-body"));
                var empty = new LlmHttpFixture("{\"models\":[]}"); LlmChatClient.HttpHandlerFactory = () => empty; Assert.AreEqual(0, (await LlmChatClient.ListModelsAsync(local, new LlmSettings())).Length);
                var claude = LlmBoundaryScope.Provider("Claude"); var settingsClaude = Settings(claude); var paginated = new LlmHttpFixture("{\"data\":[{\"id\":\"first\"}],\"has_more\":true,\"last_id\":\"a b/#\"}", "{\"data\":[{\"id\":\"last\"}],\"has_more\":false}"); Assert.AreEqual(2, (await LlmChatClient.ListModelsAsync(claude, settingsClaude, paginated)).Length); StringAssert.Contains(paginated.Uris[1].Query, "after_id=a%20b%2F%23");
                foreach (var body in new[] { "{\"data\":[],\"has_more\":true}", "{\"data\":[],\"has_more\":true,\"last_id\":\" \"}" }) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(claude, settingsClaude, new LlmHttpFixture(body)));
                string repeated = "{\"data\":[],\"has_more\":true,\"last_id\":\"same\"}"; await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(claude, settingsClaude, new LlmHttpFixture(repeated, repeated)));
                var pages = Enumerable.Range(0, 100).Select(i => "{\"data\":[],\"has_more\":true,\"last_id\":\"p" + i + "\"}").ToArray(); var limit = new LlmHttpFixture(pages); await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(claude, settingsClaude, limit)); Assert.AreEqual(100, limit.Uris.Count);
            }
        }
        [TestMethod]
        public void AuthenticationHelperPreservesRequiredHeadersAndOptionalKeyContracts()
        {
            var authenticate = typeof(LlmChatClient).GetMethod("Authenticate", BindingFlags.NonPublic | BindingFlags.Static);
            foreach (var provider in new[] { LlmBoundaryScope.Provider("Azure OpenAI"), LlmBoundaryScope.Provider("Claude"), LlmBoundaryScope.Provider("Mistral") }) foreach (bool entra in new[] { false, true }) foreach (var key in new[] { null, "", "fixture-key" }) using (var request = new HttpRequestMessage(HttpMethod.Post, "https://fixture.invalid")) { authenticate.Invoke(null, new object[] { request, provider, key, entra }); Assert.AreEqual(provider.IsAzure && !entra && !string.IsNullOrWhiteSpace(key), request.Headers.Contains("api-key")); Assert.AreEqual(provider.IsClaude, request.Headers.Contains("anthropic-version")); Assert.AreEqual(provider.IsClaude && !string.IsNullOrWhiteSpace(key), request.Headers.Contains("x-api-key")); Assert.AreEqual(!provider.IsClaude && !(provider.IsAzure && !entra) && !string.IsNullOrWhiteSpace(key), request.Headers.Authorization != null); }
        }
        [TestMethod]
        public async Task DisposalAtTheHttpBoundaryClosesTheBodyBeforeReturningProviderData()
        {
            using (var scope = new LlmBoundaryScope())
            {
                LlmChatClient client = null; var handler = new LlmHttpFixture(Answer); handler.BeforeResponse = () => client.Dispose(); using (client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", handler)) { await Assert.ThrowsExceptionAsync<ArgumentException>(() => client.CompleteAsync(History(), Tools())); Assert.AreEqual(1, handler.Uris.Count); }
            }
        }
    }
}
