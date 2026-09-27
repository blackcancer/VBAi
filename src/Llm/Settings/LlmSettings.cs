using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class LlmSettings
    {
        public string ProviderName { get; set; } = "Codex";
        // Automatic is the requested default for this prototype; unknown persisted values fail closed in LlmVbeTools.
        public string VbeEditApproval { get; set; } = "Automatic";
        public string CodexModel { get; set; }
        public string OpenAiModel { get; set; }
        public string OllamaModel { get; set; }
        public string OpenAiEndpoint { get; set; }
        public string OllamaEndpoint { get; set; }
        public string EncryptedOpenAiKey { get; set; }
        public string CustomProviderName { get; set; }
        public bool AzureUseEntraToken { get; set; }
        public Dictionary<string, string> ManualModelLists { get; set; } = new Dictionary<string, string>();
        public string GetManualModels(LlmProvider provider)
        {
            string value;
            return ManualModelLists != null && ManualModelLists.TryGetValue(provider.Name, out value) ? value :
                (provider.ModelVariable == null ? null : Environment.GetEnvironmentVariable(provider.ModelVariable));
        }
        public void ApplyProviderLabels() { foreach (var provider in LlmProvider.All) if (provider.IsCustom) provider.CustomDisplayName = CustomProviderName; }
        public Dictionary<string, string> ReasoningEfforts { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> ProviderModels { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> ProviderEndpoints { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> EncryptedProviderKeys { get; set; } = new Dictionary<string, string>();

        public string GetKey(LlmProvider provider)
        {
            if (provider.Name == "OpenAI API") return GetOpenAiKey();
            string cipher;
            if (EncryptedProviderKeys != null && EncryptedProviderKeys.TryGetValue(provider.Name, out cipher) && !string.IsNullOrEmpty(cipher))
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(cipher), null, DataProtectionScope.CurrentUser));
            return provider.IsAzure && AzureUseEntraToken ? Environment.GetEnvironmentVariable("AZURE_OPENAI_ENTRA_TOKEN") :
                provider.KeyVariable == null ? null : Environment.GetEnvironmentVariable(provider.KeyVariable);
        }

        public void SetKey(LlmProvider provider, string key)
        {
            if (provider.Name == "OpenAI API") { EncryptedOpenAiKey = null; SetOpenAiKey(key); return; }
            if (EncryptedProviderKeys == null) EncryptedProviderKeys = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(key)) { EncryptedProviderKeys.Remove(provider.Name); return; }
            EncryptedProviderKeys[provider.Name] = Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(key.Trim()), null, DataProtectionScope.CurrentUser));
        }

        public void SetEndpoint(LlmProvider provider, string endpoint)
        {
            if (provider.Name == "OpenAI API") OpenAiEndpoint = endpoint;
            else if (provider.IsOllama) OllamaEndpoint = endpoint;
            else { if (ProviderEndpoints == null) ProviderEndpoints = new Dictionary<string, string>(); ProviderEndpoints[provider.Name] = endpoint; }
        }

        private static string FilePath { get {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CodexVBE", "settings.json");
        } }

        public static LlmSettings Load()
        {
            if (!File.Exists(FilePath)) return new LlmSettings();
            string content = File.ReadAllText(FilePath, Encoding.UTF8);
            var stored = new JavaScriptSerializer().DeserializeObject(content) as System.Collections.Generic.IDictionary<string, object>;
            var loaded = new JavaScriptSerializer().Deserialize<LlmSettings>(content) ?? new LlmSettings();
            // Existing installations previously confirmed every edit. Preserve that behavior until changed in Settings.
            if (stored == null || !stored.ContainsKey("VbeEditApproval")) loaded.VbeEditApproval = "AskEachTime";
            loaded.ApplyProviderLabels();
            return loaded;
        }

        public void Save()
        {
            ApplyProviderLabels();
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
            string value = GetSelectedModel(provider);
            if (string.IsNullOrWhiteSpace(value) && provider.ModelVariable != null) value = Environment.GetEnvironmentVariable(provider.ModelVariable);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException("Configurez un modèle pour " + provider.Name + " dans les paramètres CodexVBE.");
            return value;
        }

        public string GetSelectedModel(LlmProvider provider)
        {
            if (provider.IsCodex) return CodexModel;
            if (provider.Name == "OpenAI API") return OpenAiModel;
            if (provider.IsOllama) return OllamaModel;
            string value;
            return ProviderModels != null && ProviderModels.TryGetValue(provider.Name, out value) ? value : null;
        }

        public void SetSelectedModel(LlmProvider provider, string model)
        {
            if (provider.IsCodex) CodexModel = model;
            else if (provider.Name == "OpenAI API") OpenAiModel = model;
            else if (provider.Name == "Ollama") OllamaModel = model;
            else { if (ProviderModels == null) ProviderModels = new Dictionary<string, string>(); ProviderModels[provider.Name] = model; }
        }

        public string GetReasoningEffort(LlmProvider provider, string model)
        {
            string value;
            return ReasoningEfforts != null && ReasoningEfforts.TryGetValue(provider.Name + ":" + model, out value)
                ? value : null;
        }

        public void SetReasoningEffort(LlmProvider provider, string model, string effort)
        {
            if (ReasoningEfforts == null) ReasoningEfforts = new Dictionary<string, string>();
            ReasoningEfforts[provider.Name + ":" + model] = effort;
        }

        public string ResolveEndpoint(LlmProvider provider)
        {
            string value = null;
            if (provider.Name == "OpenAI API") value = OpenAiEndpoint;
            else if (provider.IsOllama) value = OllamaEndpoint;
            else if (ProviderEndpoints != null) ProviderEndpoints.TryGetValue(provider.Name, out value);
            if (string.IsNullOrWhiteSpace(value) && provider.EndpointVariable != null)
                value = Environment.GetEnvironmentVariable(provider.EndpointVariable);
            return string.IsNullOrWhiteSpace(value) ? provider.Endpoint : value;
        }
    }
}
