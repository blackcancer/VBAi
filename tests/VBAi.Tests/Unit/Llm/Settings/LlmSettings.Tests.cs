using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmSettingsCoverageTests
    {
        [TestMethod]
        public void PersistenceUsesUtf8AndMigratesLegacyApprovalWithoutTouchingUserSettings()
        {
            var native = typeof(LlmSettings).GetProperty("FilePath", BindingFlags.NonPublic | BindingFlags.Static);
            StringAssert.EndsWith((string)native.GetValue(null), Path.Combine("VBAi", "settings.json"));
            using (var scope = new LlmBoundaryScope())
            {
                var original = LlmSettings.StoragePathOverride;
                var custom = LlmProvider.All.Single(p => p.IsCustom); var label = custom.CustomDisplayName;
                try
                {
                    LlmSettings.StoragePathOverride = Path.Combine(scope.Root, "nested", "settings.json");
                    Assert.AreEqual("Automatic", LlmSettings.Load().VbeEditApproval);
                    var settings = new LlmSettings { CustomProviderName = "équipe", GitHubAccount = "fixture", VbeEditApproval = "Never", NativeVbeDarkTheme = true };
                    settings.Save();
                    var bytes = File.ReadAllBytes(LlmSettings.StoragePathOverride); Assert.IsFalse(bytes.Take(3).SequenceEqual(new byte[] { 239, 187, 191 }));
                    var loaded = LlmSettings.Load(); Assert.AreEqual("Never", loaded.VbeEditApproval); Assert.AreEqual("équipe", custom.CustomDisplayName); Assert.AreEqual("fixture", loaded.GitHubAccount); Assert.IsTrue(loaded.NativeVbeDarkTheme);
                    foreach (var json in new[] { "{}", "null" }) { File.WriteAllText(LlmSettings.StoragePathOverride, json, Encoding.UTF8); Assert.AreEqual("AskEachTime", LlmSettings.Load().VbeEditApproval); }
                    File.WriteAllText(LlmSettings.StoragePathOverride, "{broken"); Assert.ThrowsException<ArgumentException>(() => LlmSettings.Load());
                }
                finally { LlmSettings.StoragePathOverride = original; custom.CustomDisplayName = label; }
            }
        }
        [DataTestMethod, DataRow("{}"), DataRow("null"),
            DataRow("{\"ProviderName\":\"Ollama\",\"OllamaModel\":\"synthetic-model\",\"UnknownSamplingFixture\":true}")]
        public void LegacySettingsLeaveOllamaSamplingUnspecified(string legacy)
        {
            using (var scope = new LlmBoundaryScope()) {
                string original = LlmSettings.StoragePathOverride;
                try {
                    LlmSettings.StoragePathOverride = Path.Combine(scope.Root, "legacy-sampling.json");
                    Assert.IsNull(new LlmSettings().OllamaTemperature); Assert.IsNull(new LlmSettings().OllamaTopP);
                    Assert.IsNull(LlmSettings.Load().OllamaTemperature); Assert.IsNull(LlmSettings.Load().OllamaTopP);
                    File.WriteAllText(LlmSettings.StoragePathOverride, legacy, Encoding.UTF8);
                    var loaded = LlmSettings.Load();
                    Assert.IsNull(loaded.OllamaTemperature); Assert.IsNull(loaded.OllamaTopP);
                    loaded.Save();
                    var reread = LlmSettings.Load(); Assert.IsNull(reread.OllamaTemperature); Assert.IsNull(reread.OllamaTopP);
                    if (legacy.Contains("UnknownSamplingFixture")) StringAssert.Contains(File.ReadAllText(LlmSettings.StoragePathOverride), "UnknownSamplingFixture");
                } finally { LlmSettings.StoragePathOverride = original; }
            }
        }

        [DataTestMethod, DataRow(null, null), DataRow(0.0, null), DataRow(null, 0.8), DataRow(2.0, 1.0), DataRow(0.7, 0.8)]
        public void OptionalOllamaSamplingRoundTripsAndCanBeClearedToBackendDefaults(double? temperature, double? topP)
        {
            using (var scope = new LlmBoundaryScope()) {
                string original = LlmSettings.StoragePathOverride;
                try {
                    LlmSettings.StoragePathOverride = Path.Combine(scope.Root, "sampling.json");
                    var settings = new LlmSettings { OllamaTemperature = temperature, OllamaTopP = topP, CodexModel = "unrelated" };
                    settings.Save();
                    var loaded = LlmSettings.Load(); Assert.AreEqual(temperature, loaded.OllamaTemperature); Assert.AreEqual(topP, loaded.OllamaTopP);
                    Assert.AreEqual("unrelated", loaded.CodexModel);
                    loaded.OllamaTemperature = null; loaded.OllamaTopP = null; loaded.Save();
                    var cleared = LlmSettings.Load(); Assert.IsNull(cleared.OllamaTemperature); Assert.IsNull(cleared.OllamaTopP);
                    Assert.AreEqual("unrelated", cleared.CodexModel);
                } finally { LlmSettings.StoragePathOverride = original; }
            }
        }

        [TestMethod]
        public void ConcurrentOllamaSamplingEditsMergeIndependentFieldsAndRejectConflictingValues()
        {
            using (var scope = new LlmBoundaryScope()) {
                string original = LlmSettings.StoragePathOverride;
                try {
                    LlmSettings.StoragePathOverride = Path.Combine(scope.Root, "sampling-merge.json");
                    new LlmSettings { OllamaTemperature = 0.7, OllamaTopP = 0.8, CodexModel = "initial" }.Save();
                    var first = LlmSettings.Load(); var second = LlmSettings.Load();
                    first.OllamaTemperature = 0; first.Save();
                    second.OllamaTopP = 0.9; second.CodexModel = "independent"; second.Save();
                    Assert.AreEqual(0.0, second.OllamaTemperature); Assert.AreEqual(0.9, second.OllamaTopP);
                    var merged = LlmSettings.Load(); Assert.AreEqual(0.0, merged.OllamaTemperature); Assert.AreEqual(0.9, merged.OllamaTopP);
                    Assert.AreEqual("independent", merged.CodexModel);
                    var left = LlmSettings.Load(); var right = LlmSettings.Load();
                    left.OllamaTemperature = 1; left.Save();
                    var committed = File.ReadAllBytes(LlmSettings.StoragePathOverride);
                    right.OllamaTemperature = 2;
                    var error = Assert.ThrowsException<IOException>(() => right.Save());
                    StringAssert.Contains(error.Message, "OllamaTemperature");
                    CollectionAssert.AreEqual(committed, File.ReadAllBytes(LlmSettings.StoragePathOverride));
                    Assert.AreEqual(1.0, LlmSettings.Load().OllamaTemperature);
                    left.OllamaTemperature = null; left.Save();
                    Assert.IsNull(LlmSettings.Load().OllamaTemperature); Assert.AreEqual(0.9, LlmSettings.Load().OllamaTopP);
                } finally { LlmSettings.StoragePathOverride = original; }
            }
        }

        [TestMethod]
        public void EveryProviderResolvesModelEndpointManualEffortAndEncryptedKeyPrecedence()
        {
            using (var scope = new LlmBoundaryScope())
                foreach (var provider in LlmProvider.All)
                {
                    var settings = new LlmSettings();
                    Assert.IsNull(settings.GetSelectedModel(provider));
                    Assert.ThrowsException<InvalidOperationException>(() => settings.ResolveModel(provider));
                    if (provider.ModelVariable != null) { Environment.SetEnvironmentVariable(provider.ModelVariable, "environment-model"); Assert.AreEqual("environment-model", settings.ResolveModel(provider)); }
                    settings.SetSelectedModel(provider, "selected"); Assert.AreEqual("selected", settings.GetSelectedModel(provider)); Assert.AreEqual("selected", settings.ResolveModel(provider));
                    Assert.AreEqual(provider.Endpoint, settings.ResolveEndpoint(provider));
                    if (provider.EndpointVariable != null) { Environment.SetEnvironmentVariable(provider.EndpointVariable, "https://environment.invalid"); Assert.AreEqual("https://environment.invalid", settings.ResolveEndpoint(provider)); }
                    settings.SetEndpoint(provider, "https://selected.invalid"); Assert.AreEqual("https://selected.invalid", settings.ResolveEndpoint(provider));
                    Assert.AreEqual(provider.ModelVariable == null ? null : "environment-model", settings.GetManualModels(provider)); settings.ManualModelLists[provider.Name] = "manual"; Assert.AreEqual("manual", settings.GetManualModels(provider)); settings.ManualModelLists = null; Assert.AreEqual(provider.ModelVariable == null ? null : "environment-model", settings.GetManualModels(provider));
                    Assert.IsNull(settings.GetReasoningEffort(provider, "model")); settings.ReasoningEfforts = null; Assert.IsNull(settings.GetReasoningEffort(provider, "model")); settings.SetReasoningEffort(provider, "model", "high"); Assert.AreEqual("high", settings.GetReasoningEffort(provider, "model"));
                    Assert.IsNull(settings.GetKey(provider)); if (provider.KeyVariable != null) { Environment.SetEnvironmentVariable(provider.KeyVariable, "environment-key"); Assert.AreEqual("environment-key", settings.GetKey(provider)); }
                    settings.EncryptedProviderKeys = null; settings.SetKey(provider, " fixture-key "); Assert.AreEqual(provider.Name == "OpenAI API" ? " fixture-key " : "fixture-key", settings.GetKey(provider)); settings.SetKey(provider, " "); Assert.AreEqual(provider.KeyVariable == null ? null : "environment-key", settings.GetKey(provider));
                    settings.EncryptedProviderKeys = new Dictionary<string, string> { [provider.Name] = "" }; Assert.AreEqual(provider.KeyVariable == null ? null : "environment-key", settings.GetKey(provider));
                    settings.ProviderModels = null; if (!provider.IsCodex && provider.Name != "OpenAI API" && !provider.IsOllama) Assert.IsNull(settings.GetSelectedModel(provider)); settings.SetSelectedModel(provider, "recreated"); Assert.AreEqual("recreated", settings.GetSelectedModel(provider));
                    settings.ProviderEndpoints = null; settings.SetEndpoint(provider, "https://recreated.invalid"); Assert.AreEqual("https://recreated.invalid", settings.ResolveEndpoint(provider));
                    if (provider.ModelVariable != null) Environment.SetEnvironmentVariable(provider.ModelVariable, null); if (provider.EndpointVariable != null) Environment.SetEnvironmentVariable(provider.EndpointVariable, null); if (provider.KeyVariable != null) Environment.SetEnvironmentVariable(provider.KeyVariable, null);
                }
        }
        [TestMethod]
        public void NullDictionariesDefaultsAndAzureEntraKeepTheirExplicitFallbackContracts()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var azure = LlmBoundaryScope.Provider("Azure OpenAI"); var settings = new LlmSettings { AzureUseEntraToken = true, EncryptedProviderKeys = null, ProviderEndpoints = null };
                Environment.SetEnvironmentVariable("AZURE_OPENAI_ENTRA_TOKEN", "entra-fixture"); Assert.AreEqual("entra-fixture", settings.GetKey(azure)); Assert.AreEqual(azure.Endpoint, settings.ResolveEndpoint(azure));
                var openAi = LlmBoundaryScope.Provider("OpenAI API"); settings.SetOpenAiKey("original"); settings.SetOpenAiKey(null); Assert.AreEqual("original", settings.GetOpenAiKey()); settings.EncryptedOpenAiKey = " "; Assert.IsNull(settings.GetOpenAiKey());
                settings.EncryptedOpenAiKey = "not base64"; Assert.ThrowsException<FormatException>(() => settings.GetOpenAiKey());
                settings.EncryptedProviderKeys = new Dictionary<string, string> { [azure.Name] = "not base64" }; Assert.ThrowsException<FormatException>(() => settings.GetKey(azure));
            }
        }
    }
}
