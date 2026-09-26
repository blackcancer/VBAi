using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class LlmSettings
    {
        public string ProviderName { get; set; } = "Codex";
        public string CodexModel { get; set; }
        public string OpenAiModel { get; set; }
        public string OllamaModel { get; set; }
        public string OpenAiEndpoint { get; set; }
        public string OllamaEndpoint { get; set; }
        public string EncryptedOpenAiKey { get; set; }

        private static string FilePath { get {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CodexVBE", "settings.json");
        } }

        public static LlmSettings Load()
        {
            if (!File.Exists(FilePath)) return new LlmSettings();
            return new JavaScriptSerializer().Deserialize<LlmSettings>(File.ReadAllText(FilePath, Encoding.UTF8))
                ?? new LlmSettings();
        }

        public void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            File.WriteAllText(FilePath, new JavaScriptSerializer().Serialize(this), new UTF8Encoding(false));
        }

        public string GetOpenAiKey()
        {
            if (!string.IsNullOrWhiteSpace(EncryptedOpenAiKey))
            {
                byte[] cipher = Convert.FromBase64String(EncryptedOpenAiKey);
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser));
            }
            return Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        }

        public void SetOpenAiKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            EncryptedOpenAiKey = Convert.ToBase64String(ProtectedData.Protect(
                Encoding.UTF8.GetBytes(key), null, DataProtectionScope.CurrentUser));
        }

        public string ResolveModel(LlmProvider provider)
        {
            string value = provider.Name == "OpenAI API" ? OpenAiModel : OllamaModel;
            if (string.IsNullOrWhiteSpace(value)) value = Environment.GetEnvironmentVariable(provider.ModelVariable);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException("Configurez un modèle pour " + provider.Name + " dans les paramètres CodexVBE.");
            return value;
        }

        public string GetSelectedModel(LlmProvider provider)
        {
            return provider.IsCodex ? CodexModel : provider.Name == "OpenAI API" ? OpenAiModel : OllamaModel;
        }

        public void SetSelectedModel(LlmProvider provider, string model)
        {
            if (provider.IsCodex) CodexModel = model;
            else if (provider.Name == "OpenAI API") OpenAiModel = model;
            else if (provider.Name == "Ollama") OllamaModel = model;
        }

        public string ResolveEndpoint(LlmProvider provider)
        {
            string value = provider.Name == "OpenAI API" ? OpenAiEndpoint : OllamaEndpoint;
            if (string.IsNullOrWhiteSpace(value))
                value = Environment.GetEnvironmentVariable(provider.Local ? "CODEXVBE_OLLAMA_ENDPOINT" : "CODEXVBE_OPENAI_ENDPOINT");
            return string.IsNullOrWhiteSpace(value) ? provider.Endpoint : value;
        }
    }
}
