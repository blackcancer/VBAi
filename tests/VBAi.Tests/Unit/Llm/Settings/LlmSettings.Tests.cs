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
