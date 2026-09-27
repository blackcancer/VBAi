using System;

namespace CodexVBE
{
    internal sealed class LlmProvider
    {
        public string Name { get; private set; }
        public bool Available { get; private set; }
        public bool Local { get; private set; }
        public bool IsCodex { get { return Name == "Codex"; } }
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
            new LlmProvider("Claude", false, false, null, null, null),
            new LlmProvider("GitHub Copilot", false, false, null, null, null),
            new LlmProvider("Gemini", false, false, null, null, null)
        };

        public override string ToString() { return Name + (Available ? "" : " (à venir)"); }

        public string ResolveModel()
        {
            string model = Environment.GetEnvironmentVariable(ModelVariable);
            if (string.IsNullOrWhiteSpace(model))
                throw new InvalidOperationException("Configurez " + ModelVariable + " puis redémarrez l'application hôte.");
            return model;
        }
    }
}
