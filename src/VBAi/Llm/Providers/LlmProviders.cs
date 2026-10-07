using System;

namespace VBAi
{

    /// <summary>Décrit un fournisseur LLM et les réglages nécessaires pour l’utiliser.</summary>
    internal sealed class LlmProvider
    {

        /// <summary>Nom canonique du fournisseur.</summary>
        /// <value>Nom canonique du fournisseur.</value>
        public string Name { get; private set; }

        /// <summary>Indique si le fournisseur peut être sélectionné.</summary>
        /// <value>Vrai si le fournisseur est activé.</value>
        public bool Available { get; private set; }

        /// <summary>Indique si le fournisseur est exécuté localement.</summary>
        /// <value>Vrai si le service fonctionne localement.</value>
        public bool Local { get; private set; }

        /// <summary>Indique si le fournisseur est Codex.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est Codex.</value>
        public bool IsCodex { get { return Name == "Codex"; } }

        /// <summary>Indique si le fournisseur est GitHub Copilot.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est GitHub Copilot.</value>
        public bool IsCopilot { get { return Name == "GitHub Copilot"; } }

        /// <summary>Indique si le fournisseur est Claude.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est Claude.</value>
        public bool IsClaude { get { return Name == "Claude"; } }

        /// <summary>Indique si le fournisseur est Ollama.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est Ollama.</value>
        public bool IsOllama { get { return Name == "Ollama"; } }

        /// <summary>Indique si le fournisseur est Azure OpenAI.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est Azure OpenAI.</value>
        public bool IsAzure { get { return Name == "Azure OpenAI"; } }

        /// <summary>Indique si le fournisseur est Amazon Bedrock.</summary>
        /// <value>Vrai lorsque le nom du fournisseur est Amazon Bedrock.</value>
        public bool IsBedrock { get { return Name == "Amazon Bedrock"; } }

        /// <summary>Indique si le fournisseur utilise une configuration OpenAI personnalisée.</summary>
        /// <value>Vrai lorsque le fournisseur utilise la configuration personnalisée.</value>
        public bool IsCustom { get { return Name == "Personnalisé (OpenAI)"; } }

        /// <summary>Indique si les modèles sont saisis manuellement pour ce fournisseur.</summary>
        /// <value>Vrai lorsque les modèles sont saisis manuellement.</value>
        public bool ManualModels { get { return IsAzure || IsBedrock || IsCustom; } }

        /// <summary>Indique si le fournisseur requiert une clé d’API configurée.</summary>
        /// <value>Vrai lorsque la configuration du fournisseur requiert une clé.</value>
        public bool RequiresKey { get { return !Local && !IsCopilot && !IsCustom; } }

        /// <summary>Nom d’affichage choisi pour le fournisseur personnalisé.</summary>
        /// <value>Nom d’affichage facultatif de la configuration personnalisée.</value>
        public string CustomDisplayName { get; set; }

        /// <summary>Nom de la variable d’environnement qui configure le point de terminaison.</summary>
        /// <value>Nom dérivé de la variable du modèle, ou null si elle est absente.</value>
        public string EndpointVariable { get { return ModelVariable?.Replace("_MODEL", "_ENDPOINT"); } }

        /// <summary>Adresse de requête du fournisseur, si elle est fixe.</summary>
        /// <value>Adresse de requête fixe, ou null pour les fournisseurs qui la composent.</value>
        public string Endpoint { get; private set; }

        /// <summary>Nom de la variable d’environnement du modèle.</summary>
        /// <value>Nom de la variable d’environnement du modèle.</value>
        public string ModelVariable { get; private set; }

        /// <summary>Nom de la variable d’environnement de la clé d’API.</summary>
        /// <value>Nom de la variable d’environnement de la clé.</value>
        public string KeyVariable { get; private set; }

        /// <summary>Crée la description d’un fournisseur avec ses paramètres de disponibilité et de connexion.</summary>
        /// <param name="name">Nom canonique du fournisseur.</param>
        /// <param name="available">Indique si la sélection est activée.</param>
        /// <param name="local">Indique si le service est local.</param>
        /// <param name="endpoint">Point de terminaison fixe, ou null si sa construction dépend du compte.</param>
        /// <param name="modelVariable">Variable d’environnement du modèle, ou null.</param>
        /// <param name="keyVariable">Variable d’environnement de la clé, ou null.</param>
        private LlmProvider(string name, bool available, bool local, string endpoint, string modelVariable, string keyVariable)
        {
            Name = name; Available = available; Local = local; Endpoint = endpoint;
            ModelVariable = modelVariable; KeyVariable = keyVariable;
        }

        /// <summary>Catalogue des fournisseurs et paramètres connus par l’interface.</summary>
        public static readonly LlmProvider[] All = {
            new LlmProvider("Codex", true, true, null, null, null),
            new LlmProvider("OpenAI API", true, false, "https://api.openai.com/v1/chat/completions", "VBAi_OPENAI_MODEL", "OPENAI_API_KEY"),
            new LlmProvider("Ollama", true, true, "http://localhost:11434/v1/chat/completions", "VBAi_OLLAMA_MODEL", null),
            new LlmProvider("Claude", true, false, "https://api.anthropic.com/v1/messages", "VBAi_CLAUDE_MODEL", "ANTHROPIC_API_KEY"),
            new LlmProvider("GitHub Copilot", true, false, null, "VBAi_COPILOT_MODEL", null),
            new LlmProvider("Gemini", true, false, "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions", "VBAi_GEMINI_MODEL", "GEMINI_API_KEY"),
            new LlmProvider("Mistral", true, false, "https://api.mistral.ai/v1/chat/completions", "VBAi_MISTRAL_MODEL", "MISTRAL_API_KEY"),
            new LlmProvider("DeepSeek", true, false, "https://api.deepseek.com/chat/completions", "VBAi_DEEPSEEK_MODEL", "DEEPSEEK_API_KEY"),
            new LlmProvider("OpenRouter", true, false, "https://openrouter.ai/api/v1/chat/completions", "VBAi_OPENROUTER_MODEL", "OPENROUTER_API_KEY"),
            new LlmProvider("LM Studio", true, true, "http://localhost:1234/v1/chat/completions", "VBAi_LMSTUDIO_MODEL", "LM_STUDIO_API_KEY"),
            new LlmProvider("Personnalisé (OpenAI)", true, false, null, "VBAi_CUSTOM_MODEL", "VBAi_CUSTOM_API_KEY"),
            new LlmProvider("Azure OpenAI", true, false, null, "VBAi_AZURE_MODEL", "AZURE_OPENAI_API_KEY"),
            new LlmProvider("Grok", true, false, "https://api.x.ai/v1/chat/completions", "VBAi_GROK_MODEL", "XAI_API_KEY"),
            new LlmProvider("Groq", true, false, "https://api.groq.com/openai/v1/chat/completions", "VBAi_GROQ_MODEL", "GROQ_API_KEY"),
            new LlmProvider("Amazon Bedrock", true, false, null, "VBAi_BEDROCK_MODEL", "AWS_BEARER_TOKEN_BEDROCK")
        };

        /// <summary>Retourne le nom affiché, avec la mention « bientôt disponible » si le fournisseur est désactivé.</summary>
        /// <returns>Nom traduit ou nom personnalisé du fournisseur.</returns>
        public override string ToString() { return (IsCustom ? (!string.IsNullOrWhiteSpace(CustomDisplayName) ? CustomDisplayName + " (OpenAI compatible)" : UiText.Get("Custom (OpenAI)")) : Name) + (Available ? "" : UiText.Get(" (coming soon)")); }

        /// <summary>Résout l’identifiant du modèle depuis sa variable d’environnement et signale son absence.</summary>
        /// <returns>Identifiant du modèle configuré.</returns>
        /// <exception cref="InvalidOperationException">La variable du modèle est absente ou vide.</exception>
        public string ResolveModel()
        {
            string model = Environment.GetEnvironmentVariable(ModelVariable);
            if (string.IsNullOrWhiteSpace(model))
                throw new InvalidOperationException(UiText.Get("Configure ") + ModelVariable + UiText.Get(" then restart the host application."));
            return model;
        }
    }
}
