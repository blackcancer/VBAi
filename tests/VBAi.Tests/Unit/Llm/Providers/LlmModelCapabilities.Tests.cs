using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Reflection;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmModelCapabilitiesTests
    {
        private static IDictionary<string, object> Json(object value) => ClaudeProtocol.Object(value);
        private static void Unknown(LlmModelCapabilities value)
        {
            Assert.IsNull(value.ToolCalling); Assert.IsNull(value.Reasoning); Assert.IsNull(value.Vision);
        }

        [TestMethod]
        public void MetadataIsImmutableAndDistinguishesFalseFromUndeclared()
        {
            Unknown(LlmModelCapabilities.Unknown); Unknown(new LlmModelCapabilities());
            var flags = new LlmModelCapabilities(false, true, null);
            Assert.AreEqual((bool?)false, flags.ToolCalling); Assert.AreEqual((bool?)true, flags.Reasoning); Assert.IsNull(flags.Vision);
            foreach (string field in new[] { "ToolCalling", "Reasoning", "Vision" })
                Assert.IsNull(typeof(LlmModelCapabilities).GetProperty(field, BindingFlags.Instance | BindingFlags.NonPublic).GetSetMethod(true));
        }

        [TestMethod]
        public void OpenRouterReadsItsCompleteExplicitListsAndNeverInfersFromModelName()
        {
            var provider = LlmBoundaryScope.Provider("OpenRouter");
            var positive = LlmModelCapabilities.FromCatalogue(provider, Json(new { id = "unknown-model", supported_parameters = new[] { "tools", "reasoning" }, architecture = new { input_modalities = new[] { "text", "image" } } }));
            Assert.AreEqual((bool?)true, positive.ToolCalling); Assert.AreEqual((bool?)true, positive.Reasoning); Assert.AreEqual((bool?)true, positive.Vision);
            var negative = LlmModelCapabilities.FromCatalogue(provider, Json(new { id = "vision-reasoning-tool-model", supported_parameters = new[] { "temperature" }, architecture = new { input_modalities = new[] { "text" } } }));
            Assert.AreEqual((bool?)false, negative.ToolCalling); Assert.AreEqual((bool?)false, negative.Reasoning); Assert.AreEqual((bool?)false, negative.Vision);
            var legacy = LlmModelCapabilities.FromCatalogue(provider, Json(new { supported_parameters = new[] { "include_reasoning" } }));
            Assert.AreEqual((bool?)true, legacy.Reasoning); Assert.IsNull(legacy.Vision);
        }

        [TestMethod]
        public void MissingOrMalformedOpenRouterListsRemainUnknownWhileEmptyExplicitListsAreFalse()
        {
            var provider = LlmBoundaryScope.Provider("OpenRouter");
            foreach (object item in new object[] { new { }, new { supported_parameters = "tools", architecture = new { input_modalities = "image" } },
                new { supported_parameters = new object[] { "tools", 1 }, architecture = new { input_modalities = new object[] { "image", null } } } })
                Unknown(LlmModelCapabilities.FromCatalogue(provider, Json(item)));
            var empty = LlmModelCapabilities.FromCatalogue(provider, Json(new { supported_parameters = new string[0], architecture = new { input_modalities = new string[0] } }));
            Assert.AreEqual((bool?)false, empty.ToolCalling); Assert.AreEqual((bool?)false, empty.Reasoning); Assert.AreEqual((bool?)false, empty.Vision);
        }

        [DataTestMethod]
        [DataRow(true, false), DataRow(false, true)]
        public void MistralCopiesBooleanToolAndVisionDeclarationsButDoesNotInventReasoning(bool tools, bool vision)
        {
            var source = Json(new { capabilities = new { function_calling = tools, vision, reasoning = true } });
            var flags = LlmModelCapabilities.FromCatalogue(LlmBoundaryScope.Provider("Mistral"), source);
            Assert.AreEqual((bool?)tools, flags.ToolCalling); Assert.AreEqual((bool?)vision, flags.Vision); Assert.IsNull(flags.Reasoning);
            ((IDictionary<string, object>)source["capabilities"])["function_calling"] = !tools;
            Assert.AreEqual((bool?)tools, flags.ToolCalling);
            Unknown(LlmModelCapabilities.FromCatalogue(LlmBoundaryScope.Provider("Mistral"), Json(new { capabilities = new { function_calling = "true", vision = 1 } })));
        }

        [DataTestMethod]
        [DataRow(true, false), DataRow(false, true)]
        public void CopilotReadsOnlyDocumentedSdkSupportsFields(bool reasoning, bool vision)
        {
            var source = Json(new { capabilities = new { supports = new { reasoningEffort = reasoning, vision, tools = true, toolCalling = true } } });
            var flags = LlmModelCapabilities.FromCopilot(source);
            Assert.AreEqual((bool?)reasoning, flags.Reasoning); Assert.AreEqual((bool?)vision, flags.Vision); Assert.IsNull(flags.ToolCalling);
            Unknown(LlmModelCapabilities.FromCopilot(Json(new { capabilities = new { supports = new { reasoningEffort = "true", vision = 1 } } })));
            Unknown(LlmModelCapabilities.FromCopilot(null));
        }

        [TestMethod]
        public void CodexReadsEffortsAndModalitiesOnlyWhenActuallyDeclared()
        {
            var full = LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = new[] { new { reasoningEffort = "high" } }, inputModalities = new[] { "text", "image" } }));
            Assert.AreEqual((bool?)true, full.Reasoning); Assert.AreEqual((bool?)true, full.Vision); Assert.IsNull(full.ToolCalling);
            var none = LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = new[] { new { reasoningEffort = "none" } }, inputModalities = new[] { "text" } }));
            Assert.AreEqual((bool?)false, none.Reasoning); Assert.AreEqual((bool?)false, none.Vision);
            Assert.AreEqual((bool?)false, LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = new object[0] })).Reasoning);
            Unknown(LlmModelCapabilities.FromCodex(Json(new { defaultReasoningEffort = "high", model = "vision-tool-reasoning-model" })));
            Unknown(LlmModelCapabilities.FromCodex(null));
        }

        [TestMethod]
        public void MalformedCodexFieldsDoNotRemoveIndependentPositiveDeclarations()
        {
            var flags = LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = new object[] { null, new { }, new { reasoningEffort = "high" } }, inputModalities = new[] { "image" } }));
            Assert.AreEqual((bool?)true, flags.Reasoning); Assert.AreEqual((bool?)true, flags.Vision);
            var malformed = LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = new object[] { null, new { reasoningEffort = " " } }, inputModalities = new[] { "image" } }));
            Assert.IsNull(malformed.Reasoning); Assert.AreEqual((bool?)true, malformed.Vision);
            Unknown(LlmModelCapabilities.FromCodex(Json(new { supportedReasoningEfforts = "high", inputModalities = new object[] { "image", 1 } })));
        }

        [TestMethod]
        public void OtherCataloguesAndManualOptionsStayUnknownDespiteUnrelatedLookalikeFields()
        {
            var source = Json(new { id = "vision-reasoning-tool-model", supported_parameters = new[] { "tools", "reasoning" },
                architecture = new { input_modalities = new[] { "image" } }, capabilities = new { function_calling = true, vision = true } });
            foreach (var provider in LlmProvider.All)
                if (provider.Name != "OpenRouter" && provider.Name != "Mistral") Unknown(LlmModelCapabilities.FromCatalogue(provider, source));
            Unknown(LlmModelCapabilities.FromCatalogue(null, source)); Unknown(LlmModelCapabilities.FromCatalogue(LlmBoundaryScope.Provider("Mistral"), null));
            Assert.AreSame(LlmModelCapabilities.Unknown, new LlmModelOption("vision-reasoning-tool-model", "Manual").Capabilities);
        }
    }
}
