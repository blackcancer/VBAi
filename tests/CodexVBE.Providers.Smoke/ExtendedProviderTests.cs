using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CodexVBE;

/// <summary>Scénarios smoke supplémentaires pour Azure, Bedrock et les réponses en flux.</summary>
internal static partial class ProviderTests
{
    /// <summary>Formate un événement Server-Sent Events contenant un objet JSON.</summary>
    /// <param name="value">Valeur de l’événement.</param>
    /// <returns>Bloc SSE terminé par une ligne vide.</returns>
    private static string Event(object value) { return "data: " + Json.Serialize(value) + "\n\n"; }
    /// <summary>Crée une réponse HTTP de type event-stream avec le contenu fourni.</summary>
    /// <param name="data">Flux SSE simulé.</param>
    /// <returns>Réponse HTTP 200 contenant le flux.</returns>
    private static HttpResponseMessage Sse(string data) { return new HttpResponseMessage(System.Net.HttpStatusCode.OK) { Content = new StringContent(data, Encoding.UTF8, "text/event-stream") }; }
    /// <summary>Vérifie les modèles manuels, l’authentification Azure, Bedrock Converse et le décodage des flux.</summary>
    /// <returns>Tâche terminée après les scénarios asynchrones.</returns>
    private static async Task Extended()
    {
        var settings = new LlmSettings();
        var custom = Provider("Personnalisé (OpenAI)"); var azure = Provider("Azure OpenAI"); var bedrock = Provider("Amazon Bedrock");
        foreach (var p in new[] { custom, azure, bedrock }) {
            settings.ManualModelLists[p.Name] = "deployment-test\r\ndeployment-test\n second-model ";
            var list = await LlmChatClient.ListModelsAsync(p, settings);
            Assert(list.Length == 2 && list[0].Id == "deployment-test", "manual models deduplication");
        }
        settings.SetEndpoint(custom, "http://localhost:5678/v1/chat/completions");
        var customHandler = new Handler { Handle = request => { Assert(request.Headers.Authorization == null, "custom key should be optional"); return Task.FromResult(Response(new { choices = new[] { new { message = new { role = "assistant", content = "ok" } } } })); } };
        string oldKey = Environment.GetEnvironmentVariable("CODEXVBE_CUSTOM_API_KEY"); Environment.SetEnvironmentVariable("CODEXVBE_CUSTOM_API_KEY", null);
        try { using (var client = new LlmChatClient(custom, settings, "deployment-test", customHandler)) Assert(Text(await client.CompleteAsync(History(), Tools), "content") == "ok", "custom response"); }
        finally { Environment.SetEnvironmentVariable("CODEXVBE_CUSTOM_API_KEY", oldKey); }
        settings.SetEndpoint(azure, "https://example.openai.azure.com/openai/v1/chat/completions"); settings.SetKey(azure, "azure-test-key");
        foreach (bool entra in new[] { false, true }) {
            settings.AzureUseEntraToken = entra;
            var handler = new Handler { Handle = async request => {
                Assert(entra ? request.Headers.Authorization.Parameter == "azure-test-key" && !request.Headers.Contains("api-key") : request.Headers.Contains("api-key") && request.Headers.Authorization == null, "Azure auth mode");
                Assert(Text(Obj(Json.DeserializeObject(await request.Content.ReadAsStringAsync())), "model") == "deployment-test", "Azure deployment name");
                return Response(new { choices = new[] { new { message = new { role = "assistant", content = "ok" } } } });
            } };
            using (var client = new LlmChatClient(azure, settings, "deployment-test", handler)) await client.CompleteAsync(History(), Tools);
        }
        Console.WriteLine("PASS custom optional auth and manual models; Azure deployment and key/Entra authentication");

        settings.SetEndpoint(bedrock, "https://bedrock-runtime.eu-west-3.amazonaws.com"); settings.SetKey(bedrock, "bedrock-test");
        int round = 0; var bedrockHandler = new Handler { Handle = async request => {
            Assert(request.RequestUri.AbsolutePath.EndsWith("/converse") && request.Headers.Authorization.Parameter == "bedrock-test", "Bedrock endpoint/auth");
            var payload = Obj(Json.DeserializeObject(await request.Content.ReadAsStringAsync()));
            Assert(!payload.ContainsKey("model") && payload.ContainsKey("toolConfig") && payload.ContainsKey("system"), "Bedrock protocol");
            if (++round == 1) return Response(new { output = new { message = new { role = "assistant", content = new object[] {
                new { reasoningContent = new { reasoningText = new { text = "opaque", signature = "signed" } } },
                new { toolUse = new { toolUseId = "b1", name = "read_module", input = new { } } }
            } } }, stopReason = "tool_use" });
            Assert(Json.Serialize(payload).Contains("signed") && Json.Serialize(payload).Contains("toolResult"), "Bedrock roundtrip metadata/results");
            return Response(new { output = new { message = new { role = "assistant", content = new[] { new { text = "Vérifié" } } } }, stopReason = "end_turn" });
        } };
        using (var client = new LlmChatClient(bedrock, settings, "arn:aws:bedrock:eu-west-3:123456789012:inference-profile/test", bedrockHandler)) {
            var history = History(); history.Add(await client.CompleteAsync(history, Tools)); history.Add(new { role = "tool", tool_call_id = "b1", content = "Sub Test()" });
            Assert(Text(await client.CompleteAsync(history, Tools), "content") == "Vérifié", "Bedrock final");
        }
        Console.WriteLine("PASS Bedrock Converse tool request/result and signed content roundtrip");

        string stream = ": keep-alive\n\n" + Event(new { choices = new[] { new { delta = new { role = "assistant", content = "Bon" } } } }) +
            Event(new { choices = new[] { new { delta = new { content = "jour été", tool_calls = new[] { new { index = 0, id = "t1", type = "function", function = new { name = "read_module", arguments = "{\"a\":" }, extra_content = new { google = new { thought_signature = "signature" } } } } } } } }) +
            Event(new { choices = new[] { new { delta = new { tool_calls = new[] { new { index = 0, function = new { arguments = "1}" } } } }, finish_reason = "tool_calls" } } }) + "data: [DONE]\n\n";
        settings.SetKey(Provider("Groq"), "test"); var progress = new StringBuilder();
        var streamingHandler = new Handler { Handle = async request => { Assert(Equals(Obj(Json.DeserializeObject(await request.Content.ReadAsStringAsync()))["stream"], true), "stream flag missing"); return Sse(stream); } };
        using (var client = new LlmChatClient(Provider("Groq"), settings, "test", streamingHandler)) {
            client.TextDelta = text => progress.Append(text); var result = await client.CompleteAsync(History(), Tools);
            Assert(progress.ToString() == "Bonjour été" && Text(result, "content") == progress.ToString(), "SSE progressive text");
            var call = Obj(ClaudeProtocol.Array(result, "tool_calls")[0]);
            Assert(Text(Obj(call["function"]), "arguments") == "{\"a\":1}" && Json.Serialize(call).Contains("signature"), "SSE fragmented arguments/signature");
        }
        string claudeStream = Event(new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "" } }) +
            Event(new { type = "content_block_delta", index = 0, delta = new { type = "text_delta", text = "Salut" } }) +
            Event(new { type = "content_block_start", index = 1, content_block = new { type = "tool_use", id = "c1", name = "read_module", input = new { } } }) +
            Event(new { type = "content_block_delta", index = 1, delta = new { type = "input_json_delta", partial_json = "{}" } }) +
            Event(new { type = "message_delta", delta = new { stop_reason = "tool_use" } }) + Event(new { type = "message_stop" });
        using (var memory = new MemoryStream(Encoding.UTF8.GetBytes(claudeStream))) {
            var result = await ChatStreamReader.ReadAsync(memory, true, text => { }, CancellationToken.None);
            Assert(Text(result, "content") == "Salut" && ClaudeProtocol.Array(result, "tool_calls").Length == 1, "Claude SSE tools");
        }
        foreach (var invalid in new[] { stream.Replace("data: [DONE]\n\n", ""), stream.Replace("tool_calls\"}", "length\"}"), "data: {\"error\":{}}\n\n" }) {
            bool refused = false;
            try { using (var memory = new MemoryStream(Encoding.UTF8.GetBytes(invalid))) await ChatStreamReader.ReadAsync(memory, false, text => { }, CancellationToken.None); }
            catch (InvalidDataException) { refused = true; } catch (InvalidOperationException) { refused = true; }
            Assert(refused, "incomplete/error stream accepted");
        }
        using (var cancelled = new CancellationTokenSource()) {
            cancelled.Cancel(); bool stopped = false;
            try { using (var memory = new MemoryStream()) await ChatStreamReader.ReadAsync(memory, false, text => { }, cancelled.Token); } catch (OperationCanceledException) { stopped = true; }
            Assert(stopped, "stream cancellation ignored");
        }
        Console.WriteLine("PASS compatible/Claude streaming, partial tool reconstruction, signatures, interrupted streams and cancellation");
    }

    /// <summary>Exécute un test live synthétique du catalogue et du flux OpenRouter si une clé est disponible.</summary>
    /// <returns>Tâche terminée après le test live, ou immédiatement s’il est ignoré.</returns>
    private static async Task LiveOpenRouter()
    {
        // Only a synthetic prompt is transmitted. No project, user settings file or VBA source is read.
        var settings = new LlmSettings(); var provider = Provider("OpenRouter");
        if (string.IsNullOrWhiteSpace(settings.GetKey(provider))) { Console.WriteLine("NOT_RUN OpenRouter: no environment key"); return; }
        var catalogue = await LlmChatClient.ListModelsAsync(provider, settings);
        Assert(catalogue.Length > 0, "empty live catalogue");
        Console.WriteLine("PASS live OpenRouter catalogue: " + catalogue.Length + " models");
        using (var client = new LlmChatClient(provider, settings, "openrouter/free")) {
            int fragments = 0; client.TextDelta = text => fragments++;
            var messages = new List<object> { new { role = "user", content = "This is a connectivity test. Reply with exactly OK. Do not call tools." } };
            var result = await client.CompleteAsync(messages, Tools);
            Assert(!string.IsNullOrWhiteSpace(Text(result, "content")) && fragments > 0, "live streamed response missing");
            Console.WriteLine("PASS live OpenRouter streamed synthetic response: " + fragments + " fragments");
        }
    }
    /// <summary>Vérifie en live un appel d’outil synthétique OpenRouter et l’envoi de son résultat.</summary>
    /// <returns>Tâche terminée après le test live, ou immédiatement s’il est ignoré.</returns>
    private static async Task LiveOpenRouterTools()
    {
        var settings = new LlmSettings(); var provider = Provider("OpenRouter");
        if (string.IsNullOrWhiteSpace(settings.GetKey(provider))) { Console.WriteLine("NOT_RUN OpenRouter: no environment key"); return; }
        object[] tools = { new { type = "function", function = new { name = "get_test_number", description = "Return the connectivity test number.", parameters = new { type = "object", properties = new { } } } } };
        var history = new List<object> { new { role = "user", content = "Call get_test_number once, then answer with the number returned by the tool. This is a synthetic connectivity test." } };
        using (var client = new LlmChatClient(provider, settings, "openrouter/free")) {
            int fragments = 0; client.TextDelta = text => fragments++;
            var response = await client.CompleteAsync(history, tools); history.Add(response);
            var calls = ClaudeProtocol.Array(response, "tool_calls"); Assert(calls.Length == 1, "Expected one synthetic tool call");
            var call = Obj(calls[0]); Assert(Text(Obj(call["function"]), "name") == "get_test_number", "Unexpected synthetic tool");
            history.Add(new { role = "tool", tool_call_id = Text(call, "id"), content = "{\"number\":42}" });
            var final = await client.CompleteAsync(history, tools);
            Assert((Text(final, "content") ?? "").Contains("42") && fragments > 0 && ClaudeProtocol.Array(final, "tool_calls").Length == 0, "Tool-result continuation failed");
            Console.WriteLine("PASS live OpenRouter streamed tool call and synthetic result continuation (no VBA execution)");
        }
    }
}
