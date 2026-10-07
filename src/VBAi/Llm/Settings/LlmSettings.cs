using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace VBAi
{

    /// <summary>Stocke les fournisseurs, modèles, points de terminaison et secrets de configuration LLM.</summary>
    internal sealed partial class LlmSettings
    {

        /// <summary>Remplacement facultatif du chemin de stockage, principalement utilisé par les tests isolés.</summary>
        internal static string StoragePathOverride;

        /// <summary>Obtient ou définit le fournisseur sélectionné par défaut.</summary>
        /// <value>Nom du fournisseur, « Codex » par défaut.</value>
        public string ProviderName { get; set; } = "Codex";

        /// <summary>Obtient ou définit le compte GitHub utilisé par l’intégration GitHub.</summary>
        /// <value>Nom de compte, ou nul si aucun compte n’est choisi.</value>
        public string GitHubAccount { get; set; }

        /// <summary>Obtient ou définit la stratégie de confirmation des modifications VBE.</summary>
        /// <value>Identifiant de stratégie persisté.</value>
        public string VbeEditApproval { get; set; } = "Automatic";

        /// <summary>Obtient ou définit l’habillage sombre expérimental des fenêtres natives du VBE.</summary>
        /// <value><see langword="true"/> pour activer l’habillage natif dans le processus hôte.</value>
        public bool NativeVbeDarkTheme { get; set; }

        /// <summary>Obtient ou définit le modèle Codex choisi.</summary>
        /// <value>Identifiant du modèle Codex.</value>
        public string CodexModel { get; set; }

        /// <summary>Obtient ou définit le modèle sélectionné pour l’API OpenAI.</summary>
        /// <value>Identifiant du modèle OpenAI.</value>
        public string OpenAiModel { get; set; }

        /// <summary>Obtient ou définit le modèle sélectionné pour Ollama.</summary>
        /// <value>Identifiant du modèle Ollama.</value>
        public string OllamaModel { get; set; }

        /// <summary>Optional Ollama generation temperature; null preserves the backend's historical default.</summary>
        /// <value>Current ollama temperature exposed by llm settings.</value>
        public double? OllamaTemperature { get; set; }

        /// <summary>Optional Ollama nucleus sampling probability; null preserves the backend's historical default.</summary>
        /// <value>Current ollama top p exposed by llm settings.</value>
        public double? OllamaTopP { get; set; }

        /// <summary>Obtient ou définit le point de terminaison personnalisé de l’API OpenAI.</summary>
        /// <value>Adresse personnalisée, ou nul pour utiliser la valeur par défaut ou l’environnement.</value>
        public string OpenAiEndpoint { get; set; }

        /// <summary>Obtient ou définit le point de terminaison personnalisé d’Ollama.</summary>
        /// <value>Adresse personnalisée d’Ollama.</value>
        public string OllamaEndpoint { get; set; }

        /// <summary>Obtient ou définit la clé OpenAI protégée avec DPAPI pour l’utilisateur courant.</summary>
        /// <value>Clé chiffrée encodée en Base64.</value>
        public string EncryptedOpenAiKey { get; set; }

        /// <summary>Obtient ou définit le nom affiché des fournisseurs personnalisés.</summary>
        /// <value>Libellé personnalisé des fournisseurs concernés.</value>
        public string CustomProviderName { get; set; }

        /// <summary>Obtient ou définit si Azure OpenAI utilise le jeton Entra fourni par variable d’environnement.</summary>
        /// <value><see langword="true"/> pour utiliser le jeton Entra.</value>
        public bool AzureUseEntraToken { get; set; }

        /// <summary>Obtient ou définit les listes de modèles saisies manuellement, indexées par fournisseur.</summary>
        /// <value>Dictionnaire des listes manuelles par nom de fournisseur.</value>
        public Dictionary<string, string> ManualModelLists { get; set; } = new Dictionary<string, string>();

        /// <summary>Retourne la liste de modèles manuelle du fournisseur ou, à défaut, sa variable d’environnement.</summary>
        /// <param name="provider">Fournisseur dont il faut lire la liste.</param>
        /// <returns>Liste configurée ou valeur de variable d’environnement ; nul si aucune n’existe.</returns>
        public string GetManualModels(LlmProvider provider)
        {
            return ManualModelLists != null && ManualModelLists.TryGetValue(provider.Name, out string value) ? value :
                (provider.ModelVariable == null ? null : Environment.GetEnvironmentVariable(provider.ModelVariable));
        }

        /// <summary>Applique le nom personnalisé enregistré aux fournisseurs personnalisables.</summary>
        public void ApplyProviderLabels() { foreach (var provider in LlmProvider.All) if (provider.IsCustom) provider.CustomDisplayName = CustomProviderName; }

        /// <summary>Obtient ou définit les niveaux de raisonnement, indexés par fournisseur et modèle.</summary>
        /// <value>Dictionnaire indexé par le nom du fournisseur et l’identifiant du modèle.</value>
        public Dictionary<string, string> ReasoningEfforts { get; set; } = new Dictionary<string, string>();

        /// <summary>Obtient ou définit les modèles choisis pour les fournisseurs génériques.</summary>
        /// <value>Dictionnaire des identifiants de modèles.</value>
        public Dictionary<string, string> ProviderModels { get; set; } = new Dictionary<string, string>();

        /// <summary>Obtient ou définit les points de terminaison des fournisseurs génériques.</summary>
        /// <value>Dictionnaire des URL configurées.</value>
        public Dictionary<string, string> ProviderEndpoints { get; set; } = new Dictionary<string, string>();

        /// <summary>Obtient ou définit les clés fournisseurs protégées par utilisateur.</summary>
        /// <value>Dictionnaire de clés DPAPI encodées en Base64.</value>
        public Dictionary<string, string> EncryptedProviderKeys { get; set; } = new Dictionary<string, string>();

        /// <summary>Résout la clé du fournisseur depuis le secret protégé, une variable ou le jeton Azure Entra.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <returns>Clé ou jeton, ou nul si aucune source n’est configurée.</returns>
        public string GetKey(LlmProvider provider)
        {
            if (provider.Name == "OpenAI API") return GetOpenAiKey();
            if (EncryptedProviderKeys != null && EncryptedProviderKeys.TryGetValue(provider.Name, out string cipher) && !string.IsNullOrEmpty(cipher))
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(cipher), null, DataProtectionScope.CurrentUser));
            return provider.IsAzure && AzureUseEntraToken ? Environment.GetEnvironmentVariable("AZURE_OPENAI_ENTRA_TOKEN") :
                provider.KeyVariable == null ? null : Environment.GetEnvironmentVariable(provider.KeyVariable);
        }

        /// <summary>Enregistre la clé du fournisseur avec DPAPI ou retire la clé protégée si l’entrée est vide.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <param name="key">Clé en clair à protéger ; les espaces périphériques sont supprimés.</param>
        public void SetKey(LlmProvider provider, string key)
        {
            if (provider.Name == "OpenAI API") { EncryptedOpenAiKey = null; SetOpenAiKey(key); return; }
            if (EncryptedProviderKeys == null) EncryptedProviderKeys = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(key)) { EncryptedProviderKeys.Remove(provider.Name); return; }
            EncryptedProviderKeys[provider.Name] = Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(key.Trim()), null, DataProtectionScope.CurrentUser));
        }

        /// <summary>Enregistre le point de terminaison propre au fournisseur.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <param name="endpoint">Adresse configurée.</param>
        public void SetEndpoint(LlmProvider provider, string endpoint)
        {
            if (provider.Name == "OpenAI API") OpenAiEndpoint = endpoint;
            else if (provider.IsOllama) OllamaEndpoint = endpoint;
            else { if (ProviderEndpoints == null) ProviderEndpoints = new Dictionary<string, string>(); ProviderEndpoints[provider.Name] = endpoint; }
        }

        /// <summary>Obtient le chemin du fichier settings.json dans le dossier AppData utilisateur.</summary>
        /// <value>Chemin complet de la configuration utilisateur.</value>
        private static string FilePath
        {
            get
            {
                return StoragePathOverride ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "VBAi", "settings.json");
            }
        }

        /// <summary>Charge les paramètres depuis AppData et applique la valeur historique de confirmation aux fichiers anciens.</summary>
        /// <returns>Paramètres désérialisés ou configuration par défaut si le fichier n’existe pas.</returns>
        public static LlmSettings Load()
        {
            string path = Path.GetFullPath(FilePath);
            bool exists = File.Exists(path);
            var loaded = exists ? DecodeSettings(File.ReadAllText(path, Encoding.UTF8)) : new LlmSettings();
            loaded.RememberBaseline(path, exists);
            loaded.ApplyProviderLabels();
            return loaded;
        }

        /// <summary>Applique les noms personnalisés puis enregistre les paramètres JSON en UTF-8 sans BOM.</summary>
        public void Save()
        {
            SaveMerged();
        }

        /// <summary>Déchiffre la clé OpenAI enregistrée ou retourne la variable OPENAI_API_KEY.</summary>
        /// <returns>Clé d’API en clair ou valeur de variable d’environnement.</returns>
        public string GetOpenAiKey()
        {
            if (!string.IsNullOrWhiteSpace(EncryptedOpenAiKey))
            {
                byte[] cipher = Convert.FromBase64String(EncryptedOpenAiKey);
                return Encoding.UTF8.GetString(ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser));
            }
            return Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        }

        /// <summary>Protège et mémorise la clé OpenAI pour l’utilisateur Windows courant.</summary>
        /// <param name="key">Clé en clair ; une valeur nulle ou blanche ne modifie pas le stockage.</param>
        public void SetOpenAiKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            EncryptedOpenAiKey = Convert.ToBase64String(ProtectedData.Protect(
                Encoding.UTF8.GetBytes(key), null, DataProtectionScope.CurrentUser));
        }

        /// <summary>Résout le modèle sélectionné ou sa variable d’environnement, et échoue si aucun n’est configuré.</summary>
        /// <param name="provider">Fournisseur pour lequel obtenir un modèle.</param>
        /// <returns>Identifiant de modèle utilisable.</returns>
        /// <exception cref="InvalidOperationException">Aucun modèle n’est configuré pour ce fournisseur.</exception>
        public string ResolveModel(LlmProvider provider)
        {
            string value = GetSelectedModel(provider);
            if (string.IsNullOrWhiteSpace(value) && provider.ModelVariable != null) value = Environment.GetEnvironmentVariable(provider.ModelVariable);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException(UiText.Get("Configure a model for ") + provider.Name + UiText.Get(" in VBAi settings."));
            return value;
        }

        /// <summary>Retourne le modèle actuellement enregistré pour le fournisseur.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <returns>Identifiant enregistré, ou nul s’il n’existe pas.</returns>
        public string GetSelectedModel(LlmProvider provider)
        {
            if (provider.IsCodex) return CodexModel;
            if (provider.Name == "OpenAI API") return OpenAiModel;
            if (provider.IsOllama) return OllamaModel;
            return ProviderModels != null && ProviderModels.TryGetValue(provider.Name, out string value) ? value : null;
        }

        /// <summary>Enregistre le modèle choisi dans la propriété dédiée ou la table des fournisseurs génériques.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <param name="model">Identifiant du modèle choisi.</param>
        public void SetSelectedModel(LlmProvider provider, string model)
        {
            if (provider.IsCodex) CodexModel = model;
            else if (provider.Name == "OpenAI API") OpenAiModel = model;
            else if (provider.Name == "Ollama") OllamaModel = model;
            else { if (ProviderModels == null) ProviderModels = new Dictionary<string, string>(); ProviderModels[provider.Name] = model; }
        }

        /// <summary>Retourne le niveau de raisonnement enregistré pour le couple fournisseur et modèle.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <param name="model">Modèle concerné.</param>
        /// <returns>Niveau enregistré ou nul.</returns>
        public string GetReasoningEffort(LlmProvider provider, string model)
        {
            return ReasoningEfforts != null && ReasoningEfforts.TryGetValue(provider.Name + ":" + model, out string value)
                ? value : null;
        }

        /// <summary>Enregistre le niveau de raisonnement sous la clé composée du fournisseur et du modèle.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <param name="model">Modèle concerné.</param>
        /// <param name="effort">Niveau à enregistrer.</param>
        public void SetReasoningEffort(LlmProvider provider, string model, string effort)
        {
            if (ReasoningEfforts == null) ReasoningEfforts = new Dictionary<string, string>();
            ReasoningEfforts[provider.Name + ":" + model] = effort;
        }

        /// <summary>Résout le point de terminaison enregistré, la variable d’environnement ou l’adresse par défaut.</summary>
        /// <param name="provider">Fournisseur concerné.</param>
        /// <returns>Adresse effective du point de terminaison.</returns>
        public string ResolveEndpoint(LlmProvider provider)
        {
            string value = null;
            if (provider.Name == "OpenAI API") value = OpenAiEndpoint;
            else if (provider.IsOllama) value = OllamaEndpoint;
            else ProviderEndpoints?.TryGetValue(provider.Name, out value);
            if (string.IsNullOrWhiteSpace(value) && provider.EndpointVariable != null)
                value = Environment.GetEnvironmentVariable(provider.EndpointVariable);
            return string.IsNullOrWhiteSpace(value) ? provider.Endpoint : value;
        }
    }
}
