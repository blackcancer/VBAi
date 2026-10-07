using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{
    /// <summary>Immutable catalog declarations; null means the provider has not declared that capability.</summary>
    internal sealed class LlmModelCapabilities
    {
        internal static LlmModelCapabilities Unknown { get; } = new LlmModelCapabilities();
        internal bool? ToolCalling { get; }
        internal bool? Reasoning { get; }
        internal bool? Vision { get; }

        internal LlmModelCapabilities(bool? toolCalling = null, bool? reasoning = null, bool? vision = null)
        {
            ToolCalling = toolCalling; Reasoning = reasoning; Vision = vision;
        }

        /// <summary>Reads only the selected provider's documented catalog fields, without model-name inference.</summary>
        internal static LlmModelCapabilities FromCatalogue(LlmProvider provider, IDictionary<string, object> item)
        {
            if (provider == null || item == null) return Unknown;
            if (provider.Name == "OpenRouter")
            {
                string[] parameters = Strings(item, "supported_parameters");
                string[] inputs = Strings(Object(item, "architecture"), "input_modalities");
                return new LlmModelCapabilities(Contains(parameters, "tools"),
                    parameters == null ? (bool?)null : parameters.Contains("reasoning", StringComparer.Ordinal) || parameters.Contains("include_reasoning", StringComparer.Ordinal),
                    Contains(inputs, "image"));
            }
            if (provider.Name == "Mistral")
            {
                var declarations = Object(item, "capabilities");
                return new LlmModelCapabilities(Boolean(declarations, "function_calling"), null, Boolean(declarations, "vision"));
            }
            return Unknown;
        }

        /// <summary>Copilot declares vision and reasoning-effort support; its SDK catalog has no tool-calling support flag.</summary>
        internal static LlmModelCapabilities FromCopilot(IDictionary<string, object> item)
        {
            var supports = Object(Object(item, "capabilities"), "supports");
            return new LlmModelCapabilities(null, Boolean(supports, "reasoningEffort"), Boolean(supports, "vision"));
        }

        /// <summary>Reads declared Codex reasoning-effort options and input modalities without filling schema defaults.</summary>
        internal static LlmModelCapabilities FromCodex(IDictionary<string, object> item)
        {
            if (item == null) return Unknown;
            bool? reasoning = null;
            if (item.TryGetValue("supportedReasoningEfforts", out var raw) && raw is object[] options)
            {
                bool positive = false, malformed = false;
                foreach (object option in options)
                {
                    if (!(option is IDictionary<string, object> declaration) || !declaration.TryGetValue("reasoningEffort", out var effort) ||
                        !(effort is string text) || string.IsNullOrWhiteSpace(text)) { malformed = true; continue; }
                    if (text != "none") positive = true;
                }
                reasoning = positive ? true : malformed ? (bool?)null : false;
            }
            return new LlmModelCapabilities(null, reasoning, Contains(Strings(item, "inputModalities"), "image"));
        }

        private static IDictionary<string, object> Object(IDictionary<string, object> item, string key) =>
            item != null && item.TryGetValue(key, out var value) ? value as IDictionary<string, object> : null;

        private static bool? Boolean(IDictionary<string, object> item, string key) =>
            item != null && item.TryGetValue(key, out var value) && value is bool flag ? (bool?)flag : null;

        private static string[] Strings(IDictionary<string, object> item, string key)
        {
            if (item == null || !item.TryGetValue(key, out var raw) || !(raw is object[] values) || values.Any(value => !(value is string))) return null;
            return values.Cast<string>().ToArray();
        }

        private static bool? Contains(string[] values, string member) =>
            values == null ? (bool?)null : values.Contains(member, StringComparer.Ordinal);
    }
}
