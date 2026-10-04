using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeStreamOracleTests
    {
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
