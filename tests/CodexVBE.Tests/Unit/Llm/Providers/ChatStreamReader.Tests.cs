namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class StreamTests
    {
        [TestMethod]
        public async Task OpenAiToolArgumentsAreReassembledByIndexAndTextIsStreamed()
        {
            var progress = new StringBuilder();
            using (var stream = Events(new { choices = new[] { new { delta = new { content = "Bonjour ", tool_calls = new[] { new { index = 1, id = "b", type = "function", function = new { name = "write", arguments = "{\"value\":" } }, new { index = 0, id = "a", type = "function", function = new { name = "read", arguments = "{\"path\":" } } } } } } }, new { choices = new[] { new { delta = new { content = "été", tool_calls = new[] { new { index = 1, function = new { arguments = "2}" } }, new { index = 0, function = new { arguments = "\"Module1\"}" } } } } } } }, new { choices = new[] { new { finish_reason = "tool_calls" } } }, "[DONE]"))
            {
                var result = await ChatStreamReader.ReadAsync(stream, false, s => progress.Append(s), CancellationToken.None);
                Assert.AreEqual("Bonjour été", result["content"]);
                Assert.AreEqual("Bonjour été", progress.ToString());
                var calls = (object[])result["tool_calls"];
                Assert.AreEqual("a", Obj(calls[0])["id"]);
                Assert.AreEqual("b", Obj(calls[1])["id"]);
                Assert.AreEqual("read", Obj(Obj(calls[0])["function"])["name"]);
                Assert.AreEqual("{\"path\":\"Module1\"}", Obj(Obj(calls[0])["function"])["arguments"]);
                Assert.AreEqual("{\"value\":2}", Obj(Obj(calls[1])["function"])["arguments"]);
            }
        }

        [TestMethod]
        public async Task MissingDoneOrInvalidFinishReasonRejectsPartialToolCall()
        {
            using (var stream = Events(new { choices = new[] { new { delta = new { tool_calls = new[] { new { index = 0, id = "partial", function = new { name = "write", arguments = "{" } } } } } } }))
            {
                await Assert.ThrowsExceptionAsync<InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
            }

            using (var stream = Events(new { choices = new[] { new { finish_reason = "content_filter" } } }, "[DONE]"))
            {
                await Assert.ThrowsExceptionAsync<InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
            }
        }

        [TestMethod]
        public async Task ProviderErrorDoesNotLeakResponseBody()
        {
            using (var stream = Events(new { error = new { message = "secret-body" } }, "[DONE]"))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
                Assert.IsFalse(error.Message.Contains("secret-body"));
            }
        }

        [TestMethod]
        public async Task ClaudeToolInputAndThinkingSignatureSurviveStreaming()
        {
            var progress = new StringBuilder();
            using (var stream = Events(new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "Ré" } }, new { type = "content_block_delta", index = 0, delta = new { type = "text_delta", text = "ponse" } }, new { type = "content_block_start", index = 1, content_block = new { type = "thinking", thinking = "", signature = "" } }, new { type = "content_block_delta", index = 1, delta = new { type = "thinking_delta", thinking = "private" } }, new { type = "content_block_delta", index = 1, delta = new { type = "signature_delta", signature = "sig" } }, new { type = "content_block_start", index = 2, content_block = new { type = "tool_use", id = "call-1", name = "read_module", input = new { } } }, new { type = "content_block_delta", index = 2, delta = new { type = "input_json_delta", partial_json = "{\"module\":" } }, new { type = "content_block_delta", index = 2, delta = new { type = "input_json_delta", partial_json = "\"M1\"}" } }, new { type = "message_delta", delta = new { stop_reason = "tool_use" } }, new { type = "message_stop" }))
            {
                var result = await ChatStreamReader.ReadAsync(stream, true, s => progress.Append(s), CancellationToken.None);
                Assert.AreEqual("Réponse", result["content"]);
                Assert.AreEqual("Réponse", progress.ToString());
                var original = (object[])result["_claude_content"];
                Assert.AreEqual("private", Obj(original[1])["thinking"]);
                Assert.AreEqual("sig", Obj(original[1])["signature"]);
                var call = Obj(((object[])result["tool_calls"])[0]);
                Assert.AreEqual("call-1", call["id"]);
                Assert.AreEqual("M1", Obj(Json.DeserializeObject((string)Obj(call["function"])["arguments"]))["module"]);
            }
        }

        [TestMethod]
        public async Task CancellationBeforeReadingIsObserved()
        {
            using (var source = new CancellationTokenSource())
            using (var stream = Events("[DONE]"))
            {
                source.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => ChatStreamReader.ReadAsync(stream, false, null, source.Token));
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    public sealed partial class StreamTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task OpenAiFramingMergesNullScalarsNestedValuesAndMetadata()
        {
            var wire = ": comment\nevent: update\n\ndata: {\"choices\":[{\"delta\":\ndata: {\"role\":\"assistant\",\"type\":\"message\",\"content\":null,\"count\":1,\"nested\":\"before\"}}]}\n\n";
            wire += System.Text.Encoding.UTF8.GetString(Events(
                new { choices = new object[0] },
                new { choices = new[] { new { delta = new { role = "assistant", type = "message", content = "", count = 2, nested = new { value = "first" } } } } },
                new { choices = new[] { new { delta = new { nested = new { value = "second" }, content = "text" } } } },
                new { choices = new[] { new { finish_reason = "stop" } } }, "[DONE]").ToArray());
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(wire))) {
                var result = await ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("assistant", result["role"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("message", result["type"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, result["count"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("firstsecond", Obj(result["nested"])["value"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("text", result["content"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(result.ContainsKey("tool_calls"));
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task StreamStopsRejectInvalidCompletionForBothProtocols()
        {
            foreach (var stop in new[] { null, "length", "tool_calls", "stop", "end_turn", "tool_use", "stop_sequence" }) {
                using (var stream = Events(new { choices = new[] { new { delta = new { content = "ok" }, finish_reason = stop } } }, "[DONE]")) {
                    if (stop == "stop" || stop == "tool_calls") {
                        var result = await ChatStreamReader.ReadAsync(stream, false, value => { }, System.Threading.CancellationToken.None);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("ok", result["content"]);
                    } else await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None));
                }
                using (var stream = Events(new { type = "message_delta", delta = new { stop_reason = stop } }, new { type = "message_stop" })) {
                    if (stop == "end_turn" || stop == "tool_use" || stop == "stop_sequence") {
                        var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", result["content"]);
                    } else await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
                }
            }
            using (var stream = Events(new { type = "message_delta", delta = new { stop_reason = "end_turn" } })) {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
            }
            using (var stream = Events(new { type = "error", message = "private" })) {
                var error = await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.InvalidOperationException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(error.Message.Contains("private"));
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ClaudeIgnoresUnknownDeltaAndHandlesEmptyTextWithoutProgress()
        {
            using (var stream = Events(
                new { type = "message_start" },
                new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "" } },
                new { type = "content_block_delta", index = 0, delta = new { type = "unknown", text = "ignored" } },
                new { type = "content_block_delta", index = 0, delta = new { type = "text_delta", text = "answer" } },
                new { type = "message_delta", delta = new { stop_reason = "end_turn" } }, new { type = "message_stop" })) {
                var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("answer", result["content"]);
            }
            using (var stream = Events(new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "initial" } },
                new { type = "message_delta", delta = new { stop_reason = "stop_sequence" } }, new { type = "message_stop" })) {
                var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("initial", result["content"]);
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task OversizeInputAndCancellationDuringProgressAreRejected()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(new string('x', 10 * 1024 * 1024 + 1) + "\n"))) {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None));
            }
            using (var source = new System.Threading.CancellationTokenSource())
            using (var stream = Events(new { choices = new[] { new { delta = new { content = "first" } } } }, "[DONE]")) {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.OperationCanceledException>(() =>
                    ChatStreamReader.ReadAsync(stream, false, value => source.Cancel(), source.Token));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(stream.CanRead);
            }
        }
    }
}
