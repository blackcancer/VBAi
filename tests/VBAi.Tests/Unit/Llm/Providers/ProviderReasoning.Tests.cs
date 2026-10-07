namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;

    [TestClass, TestCategory("Unit")]
    public sealed class ProviderReasoningTests
    {
        [TestMethod]
        public void StructuredPublicSectionsTakePriorityWithoutShowingOpaqueFields()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "mirrored", reasoning_details = new object[] {
                new { type = "reasoning.text", index = 0, text = "Public A", signature = "SECRET_SIGNATURE" },
                new { type = "reasoning.summary", index = 1, summary = "Public B" },
                new { type = "reasoning.encrypted", index = 2, data = "SECRET_DATA", text = "SECRET_TEXT" } } }));
            reader.Finish("completed");
            var final = seen.Where(a => a.Status == "completed").ToArray();
            Assert.AreEqual(2, final.Length);
            CollectionAssert.AreEqual(new[] { "Public A", "Public B" }, final.Select(a => a.Detail).ToArray());
            Assert.AreEqual(2, final.Select(a => a.Id).Distinct().Count());
            Assert.IsFalse(seen.Any(a => a.Detail.Contains("SECRET") || a.Detail.Contains("mirrored")));
        }

        [TestMethod]
        public void CompatibilityFieldsAreWhitelistedAndFinalSnapshotReplacesDeltas()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            reader.OpenAi(ClaudeProtocol.Object(new { content = "<think>HIDDEN</think>", signature = "HIDDEN", thinking = "HIDDEN", reasoning_content = "Part ", reasoning = "mirrored" }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "one" }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "Final corrected" }), true);
            reader.Finish("completed"); reader.Finish("failed");
            Assert.AreEqual(1, seen.Select(a => a.Id).Distinct().Count());
            Assert.AreEqual("Final corrected", seen.Last().Detail); Assert.IsFalse(seen.Last().Append);
            Assert.AreEqual("completed", seen.Last().Status); Assert.IsFalse(seen.Any(a => a.Detail.Contains("HIDDEN") || a.Detail.Contains("mirrored")));
        }

        [TestMethod]
        public void DetailIndexKeepsIdentityWhenLaterFragmentsOmitServerId()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_details = new[] { new { type = "reasoning.text", id = "initial-server-id", index = 3, text = "First " } } }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_details = new[] { new { type = "reasoning.text", index = 3, text = "second" } } }));
            reader.Finish("completed");
            Assert.AreEqual(1, seen.Select(a => a.Id).Distinct().Count()); Assert.AreEqual("First second", seen.Last().Detail);
        }

        [TestMethod]
        public void StructuredRepresentationReusesCompatibilityCardAndNeverReplaysMirrors()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "Public summary" }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_details = new[] { new { type = "reasoning.text", index = 0, text = "Public " } } }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_details = new[] { new { type = "reasoning.text", index = 0, text = "summary" } }, reasoning_content = "Public summary" }));
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "Public summary" }));
            reader.Finish("completed");
            Assert.AreEqual(1, seen.Select(a => a.Id).Distinct().Count()); Assert.AreEqual("Public summary", seen.Last().Detail);
        }

        [TestMethod]
        public void EmptyAndOpaqueBlocksDoNotInventReasoningCards()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            reader.OpenAi(ClaudeProtocol.Object(new { reasoning_content = "  ", content = "<think>not public</think>" }));
            reader.ClaudeBlock(0, ClaudeProtocol.Object(new { type = "redacted_thinking", data = "opaque" }), true);
            reader.ClaudeDelta(0, ClaudeProtocol.Object(new { type = "redacted_thinking" }), ClaudeProtocol.Object(new { type = "thinking_delta", thinking = "opaque" }));
            reader.BedrockResponse(ClaudeProtocol.Object(new { output = new { message = new { content = new object[] { new { reasoningContent = new { redactedContent = "opaque", reasoningText = new { text = "opaque", signature = "opaque" } } } } } } }));
            reader.ClaudeDelta(1, ClaudeProtocol.Object(new { type = "thinking", redacted = true }), ClaudeProtocol.Object(new { type = "thinking_delta", thinking = "opaque", reasoning_content = "opaque" }));
            reader.Finish("completed"); Assert.AreEqual(0, seen.Count);
        }

        [DataTestMethod]
        [DataRow(null), DataRow("signature_delta"), DataRow("input_json_delta"), DataRow("text_delta"), DataRow("unknown_delta")]
        public void ClaudeIgnoresOpenAiReasoningFieldsOnOpaqueAndUnknownDeltaTypes(string type)
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            foreach (string block in new[] { "thinking", "text", "tool_use", "unknown" })
                reader.ClaudeDelta(0, ClaudeProtocol.Object(new { type = block }), ClaudeProtocol.Object(new { type,
                    thinking = "unapproved", signature = "opaque", reasoning_content = "unapproved", reasoning = "unapproved", reasoning_text = "unapproved",
                    reasoning_details = new[] { new { type = "reasoning.text", index = 0, text = "unapproved" } } }));
            reader.Finish("completed"); Assert.AreEqual(0, seen.Count);
        }

        [TestMethod]
        public void ClaudeProjectsOnlyNativeThinkingFieldFromMatchingBlockAndDelta()
        {
            var seen = new List<CodexAgentActivity>(); var reader = new ProviderReasoning(seen.Add, "request");
            var delta = ClaudeProtocol.Object(new { type = "thinking_delta", thinking = "Public native text", reasoning_content = "unapproved" });
            foreach (string block in new[] { "text", "tool_use", "unknown" })
                reader.ClaudeDelta(0, ClaudeProtocol.Object(new { type = block }), delta);
            Assert.AreEqual(0, seen.Count);
            reader.ClaudeDelta(1, ClaudeProtocol.Object(new { type = "thinking" }), delta);
            reader.Finish("completed"); Assert.AreEqual("Public native text", seen.Last().Detail);
            Assert.AreEqual(1, seen.Select(a => a.Id).Distinct().Count()); Assert.IsFalse(seen.Any(a => a.Detail.Contains("unapproved")));
        }

        [TestMethod]
        public void InterruptedTerminalPreservesPublicTextAndDifferentRequestsHaveDifferentIds()
        {
            var first = new List<CodexAgentActivity>(); var second = new List<CodexAgentActivity>();
            var a = new ProviderReasoning(first.Add); var b = new ProviderReasoning(second.Add);
            a.OpenAi(ClaudeProtocol.Object(new { reasoning_text = "Visible" })); b.OpenAi(ClaudeProtocol.Object(new { reasoning_text = "Visible" }));
            a.Finish("interrupted"); b.Finish("failed");
            Assert.AreNotEqual(first[0].Id, second[0].Id); Assert.IsTrue(first.Last().Append);
            Assert.AreEqual("", first.Last().Detail); Assert.AreEqual("interrupted", first.Last().Status);
            Assert.AreEqual("failed", second.Last().Status); Assert.AreEqual("Visible", first[0].Detail);
        }
    }
}
