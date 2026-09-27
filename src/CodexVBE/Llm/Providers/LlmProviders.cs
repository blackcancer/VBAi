using System;

namespace CodexVBE
{
    internal sealed class LlmProvider
    {
        public string Name { get; private set; }
        public bool Available { get; private set; }
        public bool Local { get; private set; }
        public bool IsCodex { get { return Name == "Codex"; } }
        public bool IsCopilot { get { return Name == "GitHub Copilot"; } }
        public bool IsClaude { get { return Name == "Claude"; } }
        public bool IsOllama { get { return Name == "Ollama"; } }
        public bool IsAzure { get { return Name == "Azure OpenAI"; } }
        public bool IsBedrock { get { return Name == "Amazon Bedrock"; } }
        public bool IsCustom { get { return Name == "Personnalisé (OpenAI)"; } }
        public bool ManualModels { get { return IsAzure || IsBedrock || IsCustom; } }
        public bool RequiresKey { get { return !Local && !IsCopilot && !IsCustom; } }
        public string CustomDisplayName { get; set; }
        public string EndpointVariable { get { return ModelVariable == null ? null : ModelVariable.Replace("_MODEL", "_ENDPOINT"); } }
        public string Endpoint { get; private set; }
        public string ModelVariable { get; private set; }
        public string KeyVariable { get; private set; }

        private LlmProvider(string name, bool available, bool local, string endpoint, string modelVariable, string keyVariable)
        {
            Name = name; Available = available; Local = local; Endpoint = endpoint;
            ModelVariable = modelVariable; KeyVariable = keyVariable;
        }

        public static readonly LlmProvider[] All = {
            new LlmProvider("Codex", true, true, null, null, null),
            new LlmProvider("OpenAI API", true, false, "https://api.openai.com/v1/chat/completions", "CODEXVBE_OPENAI_MODEL", "OPENAI_API_KEY"),
            new LlmProvider("Ollama", true, true, "http://localhost:11434/v1/chat/completions", "CODEXVBE_OLLAMA_MODEL", null),
            new LlmProvider("Claude", true, false, "https://api.anthropic.com/v1/messages", "CODEXVBE_CLAUDE_MODEL", "ANTHROPIC_API_KEY"),
            new LlmProvider("GitHub Copilot", true, false, null, "CODEXVBE_COPILOT_MODEL", null),
            new LlmProvider("Gemini", true, false, "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions", "CODEXVBE_GEMINI_MODEL", "GEMINI_API_KEY"),
            new LlmProvider("Mistral", true, false, "https://api.mistral.ai/v1/chat/completions", "CODEXVBE_MISTRAL_MODEL", "MISTRAL_API_KEY"),
            new LlmProvider("DeepSeek", true, false, "https://api.deepseek.com/chat/completions", "CODEXVBE_DEEPSEEK_MODEL", "DEEPSEEK_API_KEY"),
            new LlmProvider("OpenRouter", true, false, "https://openrouter.ai/api/v1/chat/completions", "CODEXVBE_OPENROUTER_MODEL", "OPENROUTER_API_KEY"),
            new LlmProvider("LM Studio", true, true, "http://localhost:1234/v1/chat/completions", "CODEXVBE_LMSTUDIO_MODEL", "LM_STUDIO_API_KEY"),
            new LlmProvider("Personnalisé (OpenAI)", true, false, null, "CODEXVBE_CUSTOM_MODEL", "CODEXVBE_CUSTOM_API_KEY"),
            new LlmProvider("Azure OpenAI", true, false, null, "CODEXVBE_AZURE_MODEL", "AZURE_OPENAI_API_KEY"),
            new LlmProvider("Grok", true, false, "https://api.x.ai/v1/chat/completions", "CODEXVBE_GROK_MODEL", "XAI_API_KEY"),
            new LlmProvider("Groq", true, false, "https://api.groq.com/openai/v1/chat/completions", "CODEXVBE_GROQ_MODEL", "GROQ_API_KEY"),
            new LlmProvider("Amazon Bedrock", true, false, null, "CODEXVBE_BEDROCK_MODEL", "AWS_BEARER_TOKEN_BEDROCK")
        };

        public override string ToString() { return (IsCustom ? (!string.IsNullOrWhiteSpace(CustomDisplayName) ? CustomDisplayName + " (OpenAI compatible)" : UiText.Get("Custom (OpenAI)")) : Name) + (Available ? "" : UiText.Get(" (coming soon)")); }

        public string ResolveModel()
        {
            string model = Environment.GetEnvironmentVariable(ModelVariable);
            if (string.IsNullOrWhiteSpace(model))
                throw new InvalidOperationException(UiText.Get("Configure ") + ModelVariable + UiText.Get(" then restart the host application."));
            return model;
        }
    }
}
