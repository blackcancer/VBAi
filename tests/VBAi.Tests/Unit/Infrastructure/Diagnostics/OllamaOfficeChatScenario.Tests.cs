using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeChatScenarioTests
    {
        [TestMethod]
        public void StreamingRequestRetainsTheVisibleNumberedOracleWithinTheVbaContext()
        {
            string request = OllamaOfficeChatScenario.StreamingRequest();
            foreach (string expected in new[] { "VBAi", "1000", "fictional VBA object labels", "1. Object 1", "2. Object 2", "3. Object 3", "do not use tools" })
                StringAssert.Contains(request, expected);
            Assert.IsFalse(request.Contains("unrelated"));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse(request), "An echo of the request cannot qualify streaming.");
            Assert.IsFalse(OllamaOfficeStreamOracle.HasNumberedResponse(new[] { request }),
                "An echo in the transcript cannot qualify streaming.");
        }

        [TestMethod]
        public void RecoveryRequestRetainsExactAcknowledgmentWithoutAcceptingAnEcho()
        {
            string request = OllamaOfficeChatScenario.RecoveryRequest();
            StringAssert.Contains(request, "VBA assistant conversation"); StringAssert.Contains(request, "UI_READY_42");
            StringAssert.Contains(request, "Do not use tools");
            Assert.IsFalse(request.Contains("unrelated")); Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse(request));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsReadyResponse("UI_READY_42"));
        }

        [TestMethod]
        public void MarkerOracleRequiresActualModuleReadInsteadOfSupplyingTheAnswer()
        {
            const string project = @"C:\owned\Disposable.docm", marker = "NATIVE_unprompted";
            string request = OllamaOfficeChatScenario.MarkerRequest(project, "Q028Marker", marker);
            StringAssert.Contains(request, "Project=" + project); StringAssert.Contains(request, "Module=Q028Marker");
            Assert.IsFalse(request.Contains(marker));
        }

        [DataTestMethod][DataRow(null, "Q028Marker", "marker")][DataRow(" ", "Q028Marker", "marker")]
        [DataRow("owned", "Production", "marker")][DataRow("owned", "Q028Marker", null)]
        public void IncompleteOwnedMarkerRequestIsRefused(string project, string module, string marker)
            => Assert.ThrowsException<ArgumentException>(() => OllamaOfficeChatScenario.MarkerRequest(project, module, marker));

        [TestMethod]
        public void PromptCannotLeakTheExpectedMarkerThroughTheProject()
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeChatScenario.MarkerRequest("NATIVE_secret", "Q028Marker", "NATIVE_secret"));

        [TestMethod]
        public void ExplicitReasoningProfileUsesTheActualLocalizedControlChoice()
        {
            Assert.AreEqual(UiText.Get("Reasoning off"), OllamaOfficeChatScenario.ReasoningLabel("none"));
            Assert.AreEqual(UiText.Get("Model default"), OllamaOfficeChatScenario.ReasoningLabel("default"));
            foreach (string unsupported in new[] { null, "", "high", "None" })
                Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeChatScenario.ReasoningLabel(unsupported));
        }
    }
}
