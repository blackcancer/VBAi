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

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class StreamTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        private static IDictionary<string, object> Obj(object value)
        {
            return (IDictionary<string, object>)value;
        }

        private static MemoryStream Events(params object[] events)
        {
            var text = string.Join("", events.Select(e => "data: " + (e is string ? e : Json.Serialize(e)) + "\n\n"));
            return new MemoryStream(Encoding.UTF8.GetBytes(text));
        }

        [TestMethod]
        public async Task OpenAiToolArgumentsAreReassembledByIndexAndTextIsStreamed()
        {
            var progress = new StringBuilder();
            using (var stream = Events(
                new { choices = new[] { new { delta = new { content = "Bonjour ", tool_calls = new[] {
                    new { index = 1, id = "b", type = "function", function = new { name = "write", arguments = "{\"value\":" } },
                    new { index = 0, id = "a", type = "function", function = new { name = "read", arguments = "{\"path\":" } }
                } } } } },
                new { choices = new[] { new { delta = new { content = "été", tool_calls = new[] {
                    new { index = 1, function = new { arguments = "2}" } },
                    new { index = 0, function = new { arguments = "\"Module1\"}" } }
                } } } } },
                new { choices = new[] { new { finish_reason = "tool_calls" } } },
                "[DONE]"))
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
            using (var stream = Events(new { choices = new[] { new { delta = new { tool_calls = new[] {
                new { index = 0, id = "partial", function = new { name = "write", arguments = "{" } }
            } } } } }))
            {
                await Assert.ThrowsExceptionAsync<InvalidDataException>(
                    () => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
            }
            using (var stream = Events(new { choices = new[] { new { finish_reason = "content_filter" } } }, "[DONE]"))
            {
                await Assert.ThrowsExceptionAsync<InvalidDataException>(
                    () => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
            }
        }

        [TestMethod]
        public async Task ProviderErrorDoesNotLeakResponseBody()
        {
            using (var stream = Events(new { error = new { message = "secret-body" } }, "[DONE]"))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                    () => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
                Assert.IsFalse(error.Message.Contains("secret-body"));
            }
        }

        [TestMethod]
        public async Task ClaudeToolInputAndThinkingSignatureSurviveStreaming()
        {
            var progress = new StringBuilder();
            using (var stream = Events(
                new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "Ré" } },
                new { type = "content_block_delta", index = 0, delta = new { type = "text_delta", text = "ponse" } },
                new { type = "content_block_start", index = 1, content_block = new { type = "thinking", thinking = "", signature = "" } },
                new { type = "content_block_delta", index = 1, delta = new { type = "thinking_delta", thinking = "private" } },
                new { type = "content_block_delta", index = 1, delta = new { type = "signature_delta", signature = "sig" } },
                new { type = "content_block_start", index = 2, content_block = new { type = "tool_use", id = "call-1", name = "read_module", input = new { } } },
                new { type = "content_block_delta", index = 2, delta = new { type = "input_json_delta", partial_json = "{\"module\":" } },
                new { type = "content_block_delta", index = 2, delta = new { type = "input_json_delta", partial_json = "\"M1\"}" } },
                new { type = "message_delta", delta = new { stop_reason = "tool_use" } },
                new { type = "message_stop" }))
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
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(
                    () => ChatStreamReader.ReadAsync(stream, false, null, source.Token));
            }
        }
    }
}
