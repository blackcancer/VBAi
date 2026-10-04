using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeStreamOracleTests
    {
        [TestMethod]
        public void PromptEchoCannotProveCompletedRecovery()
        {
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse("Reply with exactly UI_READY_42 and nothing else."));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse("UI_READY_42 plus something else"));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse(null));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsReadyResponse("\r\nUI_READY_42\r\n"));
        }

        [TestMethod]
        public void SuccessfulDiscoveryConditionIsNotObservedTwice()
        {
            int calls = 0, records = 0;
            new OllamaOfficeUi(0, observation => records++).Wait(() => { calls++; return true; }, 1, "synthetic discovery");
            Assert.AreEqual(1, calls, "A successful condition may publish an identity receipt; evaluating twice duplicates it.");
            Assert.AreEqual(1, records);
        }

        [TestMethod]
        public void EchoedComposerAndSettingsCannotProveAssistantStreaming()
        {
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("Write a numbered list beginning immediately with item 1."));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("qwen2.5:7b-instruct | Temperature 0"));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("No text response."));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsNumberedResponse("1. Chair\n2. Table\n3. Pencil\n"));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsNumberedResponse("1) Chair\n2) Table\n3) Pencil\n"));
        }
    }
}
