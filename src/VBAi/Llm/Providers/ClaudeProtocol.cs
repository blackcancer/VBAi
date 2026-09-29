using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{
    // Adapt the persisted Chat Completions history to Anthropic's native Messages API.
    /// <summary>Convertit entre l’historique Chat Completions conservé par le complément et l’API Messages de Claude.</summary>
    internal static class ClaudeProtocol
    {
        /// <summary>Convertit une valeur sérialisée en dictionnaire objet, avec aller-retour JSON si nécessaire.</summary>
        /// <param name="value">Valeur à convertir.</param>
        /// <returns>Dictionnaire de propriétés, ou null si la valeur ne peut pas être convertie.</returns>
        internal static IDictionary<string, object> Object(object value)
        {
            return value as IDictionary<string, object> ??
                new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(value)) as IDictionary<string, object>;
        }
        /// <summary>Lit une propriété et la convertit en texte.</summary>
        /// <param name="value">Dictionnaire source, éventuellement null.</param>
        /// <param name="key">Nom de la propriété à lire.</param>
        /// <returns>Valeur convertie en chaîne, ou null si la clé est absente.</returns>
        internal static string Text(IDictionary<string, object> value, string key)
        { object raw; return value != null && value.TryGetValue(key, out raw) ? Convert.ToString(raw) : null; }
        /// <summary>Lit une propriété tableau du dictionnaire.</summary>
        /// <param name="value">Dictionnaire source, éventuellement null.</param>
        /// <param name="key">Nom de la propriété à lire.</param>
        /// <returns>Tableau d’objets, ou tableau vide si la propriété est absente ou de type différent.</returns>
        internal static object[] Array(IDictionary<string, object> value, string key)
        { object raw; return value != null && value.TryGetValue(key, out raw) ? raw as object[] ?? new object[0] : new object[0]; }

        /// <summary>Adapte l’historique persistant et les outils au format de l’API Messages de Claude.</summary>
        /// <param name="model">Identifiant du modèle Claude.</param>
        /// <param name="history">Messages conservés au format Chat Completions.</param>
        /// <param name="definitions">Définitions des outils disponibles.</param>
        /// <returns>Objet sérialisable correspondant au corps de requête Claude.</returns>
        public static object Request(string model, IList<object> history, object[] definitions)
        {
            var messages = new List<object>();
            var system = new List<string>();
            List<object> pendingResults = null;
            foreach (var raw in history)
            {
                var message = Object(raw);
                string role = Text(message, "role"), text = Text(message, "content");
                if (role == "system") { system.Add(text); continue; }
                if (role == "tool")
                {
                    if (pendingResults == null) { pendingResults = new List<object>(); messages.Add(new { role = "user", content = pendingResults }); }
                    pendingResults.Add(new { type = "tool_result", tool_use_id = Text(message, "tool_call_id"), content = text ?? "" });
                    continue;
                }
                pendingResults = null;
                if (message.ContainsKey("_claude_content")) { messages.Add(new { role, content = message["_claude_content"] }); continue; }
                var content = new List<object>();
                if (!string.IsNullOrEmpty(text)) content.Add(new { type = "text", text });
                foreach (var callRaw in Array(message, "tool_calls"))
                {
                    var call = Object(callRaw); var function = Object(call["function"]);
                    content.Add(new { type = "tool_use", id = Text(call, "id"), name = Text(function, "name"),
                        input = new JavaScriptSerializer().DeserializeObject(Text(function, "arguments") ?? "{}") });
                }
                if (content.Count > 0) messages.Add(new { role, content });
            }
            return new { model, max_tokens = 8192, system = string.Join("\n\n", system), messages,
                tools = definitions.Select(raw => { var f = Object(Object(raw)["function"]); return new {
                    name = Text(f, "name"), description = Text(f, "description"), input_schema = f["parameters"] }; }).ToArray() };
        }

        /// <summary>Convertit les blocs Claude en message assistant et en appels d’outils au format interne.</summary>
        /// <param name="response">Réponse Claude décodée.</param>
        /// <returns>Message assistant au format Chat Completions, avec le contenu Claude d’origine conservé.</returns>
        public static IDictionary<string, object> Response(IDictionary<string, object> response)
        {
            var text = new List<string>(); var calls = new List<object>();
            foreach (var raw in Array(response, "content"))
            {
                var block = Object(raw);
                if (Text(block, "type") == "text") text.Add(Text(block, "text"));
                if (Text(block, "type") == "tool_use") calls.Add(new Dictionary<string, object> {
                    ["id"] = Text(block, "id"), ["type"] = "function", ["function"] = new Dictionary<string, object> {
                        ["name"] = Text(block, "name"), ["arguments"] = new JavaScriptSerializer().Serialize(block["input"]) } });
            }
            if (Text(response, "stop_reason") == "max_tokens") throw new InvalidOperationException(UiText.Get("Claude reached its response limit. Narrow the scope of your request."));
            var result = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = string.Join("\n", text), ["_claude_content"] = Array(response, "content") };
            if (calls.Count > 0) result["tool_calls"] = calls.ToArray();
            return result;
        }
    }
}
