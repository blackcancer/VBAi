using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Convertit les messages et outils internes au format Converse d’Amazon Bedrock.</summary>
    internal static class BedrockProtocol
    {
        /// <summary>Adapte l’historique, les résultats d’outils et leurs schémas au format Bedrock.</summary>
        /// <param name="history">Messages conservés au format Chat Completions.</param>
        /// <param name="tools">Définitions des outils disponibles.</param>
        /// <returns>Objet sérialisable correspondant au corps de requête Bedrock.</returns>
        public static object Request(IList<object> history, object[] tools)
        {
            var messages = new List<object>(); var system = new List<object>();
            List<object> results = null;
            foreach (var raw in history) {
                var m = ClaudeProtocol.Object(raw); string role = ClaudeProtocol.Text(m, "role"), text = ClaudeProtocol.Text(m, "content");
                if (role == "system") { system.Add(new { text }); continue; }
                if (role == "tool") {
                    if (results == null) { results = new List<object>(); messages.Add(new { role = "user", content = results }); }
                    results.Add(new { toolResult = new { toolUseId = ClaudeProtocol.Text(m, "tool_call_id"), content = new[] { new { text = text ?? "" } } } });
                    continue;
                }
                results = null;
                if (m.ContainsKey("_bedrock_content")) { messages.Add(new { role, content = m["_bedrock_content"] }); continue; }
                var content = new List<object>();
                if (!string.IsNullOrEmpty(text)) content.Add(new { text });
                foreach (var callRaw in ClaudeProtocol.Array(m, "tool_calls")) {
                    var call = ClaudeProtocol.Object(callRaw); var f = ClaudeProtocol.Object(call["function"]);
                    content.Add(new { toolUse = new { toolUseId = ClaudeProtocol.Text(call, "id"), name = ClaudeProtocol.Text(f, "name"),
                        input = new JavaScriptSerializer().DeserializeObject(ClaudeProtocol.Text(f, "arguments") ?? "{}") } });
                }
                if (content.Count > 0) messages.Add(new { role, content });
            }
            return new { system, messages, inferenceConfig = new { maxTokens = 8192 }, toolConfig = new {
                tools = tools.Select(raw => { var f = ClaudeProtocol.Object(ClaudeProtocol.Object(raw)["function"]); return new { toolSpec = new {
                    name = ClaudeProtocol.Text(f, "name"), description = ClaudeProtocol.Text(f, "description"), inputSchema = new { json = f["parameters"] } } }; }).ToArray() } };
        }
        /// <summary>Convertit le message Converse de Bedrock en message assistant au format interne.</summary>
        /// <param name="root">Réponse Bedrock décodée.</param>
        /// <returns>Message assistant avec le contenu Bedrock d’origine conservé et les appels d’outils normalisés.</returns>
        public static IDictionary<string, object> Response(IDictionary<string, object> root)
        {
            string stop = ClaudeProtocol.Text(root, "stopReason");
            if (stop != "end_turn" && stop != "tool_use" && stop != "stop_sequence")
                throw new InvalidOperationException(UiText.Get("Incomplete or filtered Bedrock response: ") + stop);
            var message = ClaudeProtocol.Object(ClaudeProtocol.Object(root["output"])["message"]);
            var blocks = ClaudeProtocol.Array(message, "content"); var text = new List<string>(); var calls = new List<object>();
            foreach (var raw in blocks) {
                var block = ClaudeProtocol.Object(raw);
                if (block.ContainsKey("text")) text.Add(ClaudeProtocol.Text(block, "text"));
                if (block.ContainsKey("toolUse")) { var tool = ClaudeProtocol.Object(block["toolUse"]);
                    calls.Add(new Dictionary<string, object> { ["id"] = tool["toolUseId"], ["type"] = "function",
                        ["function"] = new Dictionary<string, object> { ["name"] = tool["name"], ["arguments"] = new JavaScriptSerializer().Serialize(tool["input"]) } });
                }
            }
            var result = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = string.Join("\n", text), ["_bedrock_content"] = blocks };
            if (calls.Count > 0) result["tool_calls"] = calls.ToArray();
            return result;
        }
    }
}
