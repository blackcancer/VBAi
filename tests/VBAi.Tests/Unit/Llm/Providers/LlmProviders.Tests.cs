using System;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie le catalogue et la résolution des configurations de fournisseurs LLM.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmProvidersTests
    {
        /// <summary>Contrôle les drapeaux, points de terminaison et noms affichés de chaque fournisseur.</summary>
        [TestMethod]
        public void CatalogueFlagsEndpointsAndDisplayNamesMatchProviderContracts()
        {
            foreach (var provider in LlmProvider.All) {
                Assert.IsTrue(provider.Available);
                Assert.IsNotNull(provider.Name);
                Assert.AreEqual(provider.Name == "Codex", provider.IsCodex);
                Assert.AreEqual(provider.Name == "Claude", provider.IsClaude);
                Assert.AreEqual(provider.Name == "Ollama", provider.IsOllama);
                Assert.AreEqual(provider.Name == "GitHub Copilot", provider.IsCopilot);
                Assert.AreEqual(provider.Name == "Azure OpenAI", provider.IsAzure);
                Assert.AreEqual(provider.Name == "Amazon Bedrock", provider.IsBedrock);
                Assert.AreEqual(provider.Name == "Personnalisé (OpenAI)", provider.IsCustom);
                Assert.AreEqual(provider.IsAzure || provider.IsBedrock || provider.IsCustom, provider.ManualModels);
                Assert.AreEqual(!provider.Local && !provider.IsCopilot && !provider.IsCustom, provider.RequiresKey);
                Assert.AreEqual(provider.ModelVariable == null ? null : provider.ModelVariable.Replace("_MODEL", "_ENDPOINT"), provider.EndpointVariable);
                Assert.IsFalse(string.IsNullOrWhiteSpace(provider.ToString()));
                if (provider.IsCustom) {
                    var previous = provider.CustomDisplayName;
                    try {
                        provider.CustomDisplayName = "  ";
                        Assert.AreEqual(UiText.Get("Custom (OpenAI)"), provider.ToString());
                        provider.CustomDisplayName = "Private";
                        Assert.AreEqual("Private (OpenAI compatible)", provider.ToString());
                        Assert.AreEqual("Private", provider.CustomDisplayName);
                    } finally { provider.CustomDisplayName = previous; }
                } else Assert.AreEqual(provider.Name, provider.ToString());
                if (provider.Endpoint != null) Assert.IsTrue(Uri.IsWellFormedUriString(provider.Endpoint, UriKind.Absolute));
                if (provider.KeyVariable != null) Assert.IsFalse(string.IsNullOrWhiteSpace(provider.KeyVariable));
            }
            var ctor = typeof(LlmProvider).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null,
                new[] { typeof(string), typeof(bool), typeof(bool), typeof(string), typeof(string), typeof(string) }, null);
            var unavailable = (LlmProvider)ctor.Invoke(new object[] { "Future", false, false, null, "VBAi_TEST_PROTOCOL_MODEL", null });
            Assert.IsFalse(unavailable.Available);
            Assert.AreEqual("Future" + UiText.Get(" (coming soon)"), unavailable.ToString());
        }

        /// <summary>Retourne le modèle configuré et refuse les valeurs absentes ou composées d’espaces.</summary>
        [TestMethod]
        public void ResolveModelReturnsConfiguredValueAndRejectsMissingOrWhitespace()
        {
            var ctor = typeof(LlmProvider).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null,
                new[] { typeof(string), typeof(bool), typeof(bool), typeof(string), typeof(string), typeof(string) }, null);
            var key = "VBAi_TEST_PROTOCOL_MODEL";
            var provider = (LlmProvider)ctor.Invoke(new object[] { "Test", true, false, null, key, null });
            var previous = Environment.GetEnvironmentVariable(key);
            try {
                foreach (var value in new[] { null, " ", "model-test" }) {
                    Environment.SetEnvironmentVariable(key, value);
                    if (value == "model-test") Assert.AreEqual(value, provider.ResolveModel());
                    else StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => provider.ResolveModel()).Message, key);
                }
            } finally { Environment.SetEnvironmentVariable(key, previous); }
        }
    }
}
