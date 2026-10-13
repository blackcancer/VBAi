using System;

namespace VBAi
{
    /// <summary>Ollama's documented default/off controls without guessing named levels from a model name.</summary>
    internal static class OllamaReasoning
    {
        /// <summary>Local selection that deliberately omits a reasoning field from the request.</summary>
        internal const string Default = "default";

        /// <summary>OpenAI-compatible value requesting no thinking output.</summary>
        internal const string None = "none";

        /// <summary>Returns controls only when the catalogue explicitly declares thinking support.</summary>
        /// <param name="capabilities">Selected model's immutable catalogue declarations.</param>
        /// <returns>Default/off options, or an empty list when support is false or unknown.</returns>
        internal static LlmEffortOption[] Options(LlmModelCapabilities capabilities) =>
            capabilities?.Reasoning == true ? new[]
            {
                new LlmEffortOption(Default, UiText.Get("Use the model's default reasoning."), UiText.Get("Model default")),
                new LlmEffortOption(None, UiText.Get("Disable reasoning for this conversation."), UiText.Get("Reasoning off"))
            } : new LlmEffortOption[0];

        /// <summary>Validates one explicit selection before transport creation; never coerces an unsupported level.</summary>
        /// <param name="selected">Saved or current selection; null and default retain provider behavior.</param>
        /// <param name="capabilities">Declarations for this exact model connection.</param>
        /// <returns>Null to omit the field, or the documented none value.</returns>
        /// <exception cref="InvalidOperationException">The selection is unsupported or its model has not declared thinking.</exception>
        internal static string Validate(string selected, LlmModelCapabilities capabilities)
        {
            if (selected == null || selected == Default) return null;
            if (selected != None || capabilities?.Reasoning != true)
                throw new InvalidOperationException(UiText.Get("The selected reasoning option is not supported by this model."));
            return None;
        }
    }
}
