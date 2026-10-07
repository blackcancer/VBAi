using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

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
        [DataTestMethod]
        [DataRow(null, null, false), DataRow(null, null, true)]
        [DataRow(0.0, null, false), DataRow(0.0, null, true)]
        [DataRow(null, 0.8, false), DataRow(null, 0.8, true)]
        [DataRow(2.0, 1.0, false), DataRow(2.0, 1.0, true)]
        [DataRow(0.0, 0.00001, false), DataRow(0.0, 0.00001, true)]
        [DataRow(0.7, 0.8, false), DataRow(0.7, 0.8, true)]
        public async Task OllamaSamplingPayloadPreservesNullPartialBoundsAndStreaming(double? temperature, double? topP, bool streaming)
        {
            using (var scope = new LlmBoundaryScope())
            {
                var handler = new LlmHttpFixture();
                handler.Replies.Enqueue(new LlmHttpFixture.Reply(streaming ?
                    "data: {\"choices\":[{\"delta\":{\"content\":\"answer\"},\"finish_reason\":\"stop\"}]}\n\ndata: [DONE]\n\n" : Answer)
                {
                    MediaType = streaming ? "text/event-stream" : "application/json"
                });
                var settings = new LlmSettings { OllamaTemperature = temperature, OllamaTopP = topP };
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), settings, "synthetic-model", handler))
                {
                    var chunks = new List<string>();
                    if (streaming) client.TextDelta = chunks.Add;
                    Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]);
                    Assert.AreEqual(streaming, chunks.Count > 0);
                    var body = LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(handler.Bodies.Single()));
                    Assert.AreEqual(temperature.HasValue, body.ContainsKey("temperature"));
                    Assert.AreEqual(topP.HasValue, body.ContainsKey("top_p"));
                    if (temperature.HasValue) Assert.AreEqual(temperature.Value, Convert.ToDouble(body["temperature"]));
                    if (topP.HasValue) Assert.AreEqual(topP.Value, Convert.ToDouble(body["top_p"]));
                    Assert.AreEqual(streaming, body.ContainsKey("stream"));
                    Assert.IsFalse(body.ContainsKey("logprobs"));
                    Assert.IsFalse(body.ContainsKey("top_logprobs"));
                    Assert.IsFalse(body.ContainsKey("tool_choice"));
                    CollectionAssert.AreEquivalent(new[] { "model", "messages", "tools" }
                        .Concat(temperature.HasValue ? new[] { "temperature" } : new string[0])
                        .Concat(topP.HasValue ? new[] { "top_p" } : new string[0])
                        .Concat(streaming ? new[] { "stream" } : new string[0]).ToArray(), body.Keys.ToArray());
                }
            }
        }

        [TestMethod]
        public async Task OllamaSamplingUsesTheValidatedClientSnapshotAcrossLaterSettingsChanges()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var settings = new LlmSettings { OllamaTemperature = 0, OllamaTopP = 0.8 };
                var handler = new LlmHttpFixture(Answer, Answer);
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), settings, "model", handler))
                {
                    settings.OllamaTemperature = Double.NaN; settings.OllamaTopP = Double.PositiveInfinity;
                    await client.CompleteAsync(History(), Tools());
                    settings.OllamaTemperature = null; settings.OllamaTopP = null;
                    await client.CompleteAsync(History(), Tools());
                    Assert.AreEqual(2, handler.Bodies.Count);
                    foreach (string request in handler.Bodies)
                    {
                        var body = LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(request));
                        Assert.AreEqual(0.0, Convert.ToDouble(body["temperature"]));
                        Assert.AreEqual(0.8, Convert.ToDouble(body["top_p"]));
                    }
                }
                var absent = new LlmSettings(); var unchanged = new LlmHttpFixture(Answer);
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), absent, "model", unchanged))
                {
                    absent.OllamaTemperature = 0; absent.OllamaTopP = 0.8;
                    await client.CompleteAsync(History(), Tools());
                    var body = LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(unchanged.Bodies.Single()));
                    Assert.IsFalse(body.ContainsKey("temperature")); Assert.IsFalse(body.ContainsKey("top_p"));
                }
            }
        }

        [DataTestMethod]
        [DataRow(-0.001, null), DataRow(2.001, null), DataRow(Double.NaN, null)]
        [DataRow(Double.PositiveInfinity, null), DataRow(Double.NegativeInfinity, null)]
        [DataRow(null, 0.0), DataRow(null, -0.001), DataRow(null, 1.001)]
        [DataRow(null, Double.NaN), DataRow(null, Double.PositiveInfinity), DataRow(null, Double.NegativeInfinity)]
        public void InvalidOllamaSamplingIsRefusedBeforeFactoryOrSuppliedHandlerEmission(double? temperature, double? topP)
        {
            using (var scope = new LlmBoundaryScope())
            {
                var settings = new LlmSettings { OllamaTemperature = temperature, OllamaTopP = topP };
                int factoryCalls = 0;
                LlmChatClient.HttpHandlerFactory = () => { factoryCalls++; throw new InvalidOperationException("Factory must not be reached."); };
                var provider = LlmBoundaryScope.Provider("Ollama");
                var error = Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(provider, settings, "model"));
                StringAssert.Contains(error.Message, temperature.HasValue ? "OllamaTemperature" : "OllamaTopP");
                Assert.AreEqual(0, factoryCalls);
                using (var handler = new LlmHttpFixture(Answer))
                {
                    Assert.ThrowsException<InvalidOperationException>(() => new LlmChatClient(provider, settings, "model", handler));
                    Assert.AreEqual(0, handler.Bodies.Count);
                }
            }
        }

        [TestMethod]
        public async Task OtherHttpProvidersIgnoreEvenInvalidOllamaSamplingWithoutChangingNativePayloads()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (var provider in LlmProvider.All.Where(item => !item.IsCodex && !item.IsCopilot && !item.IsOllama))
                {
                    string response = provider.IsClaude ? "{\"content\":[{\"type\":\"text\",\"text\":\"answer\"}],\"stop_reason\":\"end_turn\"}" :
                        provider.IsBedrock ? "{\"output\":{\"message\":{\"content\":[{\"text\":\"answer\"}]}},\"stopReason\":\"end_turn\"}" : Answer;
                    var baseline = Settings(provider); var withOllama = Settings(provider);
                    withOllama.OllamaTemperature = Double.NaN; withOllama.OllamaTopP = Double.PositiveInfinity;
                    var original = new LlmHttpFixture(response); var configured = new LlmHttpFixture(response);
                    using (var client = new LlmChatClient(provider, baseline, "model", original)) await client.CompleteAsync(History(), Tools());
                    using (var client = new LlmChatClient(provider, withOllama, "model", configured)) await client.CompleteAsync(History(), Tools());
                    Assert.AreEqual(original.Bodies.Single(), configured.Bodies.Single(), provider.Name);
                    var body = LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(configured.Bodies.Single()));
                    Assert.IsFalse(body.ContainsKey("temperature"), provider.Name); Assert.IsFalse(body.ContainsKey("top_p"), provider.Name);
                }
            using (var scope = new LlmBoundaryScope())
            {
                scope.UseCopilot();
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("GitHub Copilot"),
                    new LlmSettings { OllamaTemperature = Double.NaN, OllamaTopP = Double.PositiveInfinity }, "model"))
                    Assert.AreEqual("model @ GitHub Copilot", client.DisplayName);
            }
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
        public async Task StreamMetadataSurvivesProtocolFailureWithoutRetryAndResetsForJsonResponse()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var handler = new LlmHttpFixture();
                handler.Replies.Enqueue(new LlmHttpFixture.Reply("data: {\"choices\":[{\"delta\":{\"content\":\"private-text\"},\"finish_reason\":\"private-reason\"}]}\n\ndata: [DONE]\n\n") { MediaType = "text/event-stream" });
                handler.Replies.Enqueue(new LlmHttpFixture.Reply(Answer));
                using (var client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", handler))
                {
                    client.TextDelta = _ => { };
                    await Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => client.CompleteAsync(History(), Tools()));
                    Assert.AreEqual(1, handler.Bodies.Count, "A failed stream must never replay the request.");
                    Assert.AreEqual("protocol-error", client.LastStreamDiagnostics.Outcome);
                    Assert.AreEqual("unknown", client.LastStreamDiagnostics.TerminalReason);
                    Assert.AreEqual(1, client.LastStreamDiagnostics.TextChunks);
                    Assert.IsFalse(new JavaScriptSerializer().Serialize(client.LastStreamDiagnostics).Contains("private"));
                    Assert.AreEqual("answer", (await client.CompleteAsync(History(), Tools()))["content"]);
                    Assert.IsNull(client.LastStreamDiagnostics, "A later JSON response must not reuse prior stream metadata.");
                    Assert.AreEqual(2, handler.Bodies.Count);
                }
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
                LlmChatClient client = null;
                var stream = new GeneratedHttpBody();
                var handler = new StreamBodyHandler(stream, HttpStatusCode.OK) { BeforeResponse = () => client.Dispose() };
                using (client = new LlmChatClient(LlmBoundaryScope.Provider("Ollama"), new LlmSettings(), "model", handler))
                {
                    client.TextDelta = _ => Assert.Fail("A cancelled request must not publish provider content.");
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.CompleteAsync(History(), Tools()));
                    Assert.AreEqual(1, handler.RequestCount, "Cancellation must not retry the request.");
                    Assert.AreEqual(0, stream.BytesRead, "Cancellation at the headers boundary must be observed before reading the body.");
                    Assert.IsTrue(stream.Disposed, "The HTTP response must close its body even when cancellation precedes reading.");
                }
            }
        }
        [TestMethod]
        public async Task OversizedCompletionAndCatalogueBodiesStopBeforeTransportBuffering()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (bool catalogue in new[] { false, true })
                {
                    var stream = new GeneratedHttpBody();
                    using (var handler = new StreamBodyHandler(stream, HttpStatusCode.OK))
                    {
                        var provider = LlmBoundaryScope.Provider("Ollama");
                        if (catalogue)
                            await Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => LlmChatClient.ListModelsAsync(provider, new LlmSettings(), handler));
                        else
                            using (var client = new LlmChatClient(provider, new LlmSettings(), "model", handler))
                                await Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => client.CompleteAsync(History(), Tools()));
                        Assert.IsTrue(stream.BytesRead <= 10 * 1024 * 1024 + 1, "An oversized HTTP body must not be buffered before the limit is checked.");
                        Assert.IsTrue(stream.Disposed);
                    }
                }
        }

        [TestMethod]
        public async Task FailedHttpStatusDoesNotReadAnUntrustedBody()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (bool catalogue in new[] { false, true })
                {
                    var stream = new GeneratedHttpBody();
                    using (var handler = new StreamBodyHandler(stream, HttpStatusCode.Forbidden))
                    {
                        var provider = LlmBoundaryScope.Provider("Ollama");
                        if (catalogue)
                            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => LlmChatClient.ListModelsAsync(provider, new LlmSettings(), handler));
                        else
                            using (var client = new LlmChatClient(provider, new LlmSettings(), "model", handler))
                                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.CompleteAsync(History(), Tools()));
                        Assert.AreEqual(0, stream.BytesRead);
                        Assert.IsTrue(stream.Disposed);
                    }
                }
        }

        private sealed class StreamBodyHandler : HttpMessageHandler
        {
            private readonly System.IO.Stream stream;
            private readonly HttpStatusCode status;
            internal Action BeforeResponse;
            internal int RequestCount;
            internal StreamBodyHandler(System.IO.Stream stream, HttpStatusCode status) { this.stream = stream; this.status = status; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, System.Threading.CancellationToken cancellationToken)
            {
                RequestCount++; BeforeResponse?.Invoke();
                return Task.FromResult(new HttpResponseMessage(status) { Content = new StreamContent(stream) });
            }
        }

        private sealed class GeneratedHttpBody : System.IO.Stream
        {
            internal int BytesRead;
            internal bool Disposed;
            public override bool CanRead => !Disposed;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => 10 * 1024 * 1024 + 8192;
            public override long Position { get => BytesRead; set => throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count)
            {
                int read = (int)Math.Min(count, Length - BytesRead);
                for (int i = 0; i < read; i++) buffer[offset + i] = (byte)'x';
                BytesRead += read; return read;
            }
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, System.Threading.CancellationToken cancellationToken)
            { cancellationToken.ThrowIfCancellationRequested(); return Task.FromResult(Read(buffer, offset, count)); }
            protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
            public override void Flush() => throw new NotSupportedException();
            public override long Seek(long offset, System.IO.SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
