namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.IO;
    using System.Security.Cryptography;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using VBAi;

    /// <summary>Vérifie le décodage des flux OpenAI compatibles et Claude.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class StreamTests
    {
        /// <summary>Replays the archived synthetic empty response without attributing its backend cause.</summary>
        [TestMethod]
        public async Task ArchivedOllamaCompleteEmptyStreamPublishesNoTextOrTools()
        {
            // Ollama 0.34.4 / qwen2.5:3b, 2026-10-01; exact synthetic captured SSE bytes.
            // This exercises the reader boundary, not the model's uncaptured pre-parser generation.
            const string archivedEvents =
                "data: {\"id\":\"chatcmpl-833\",\"object\":\"chat.completion.chunk\",\"created\":1790849037,\"model\":\"qwen2.5:3b\",\"system_fingerprint\":\"fp_ollama\",\"choices\":[{\"index\":0,\"delta\":{\"role\":\"assistant\",\"content\":\"\"},\"finish_reason\":null}]}\n\n" +
                "data: {\"id\":\"chatcmpl-833\",\"object\":\"chat.completion.chunk\",\"created\":1790849037,\"model\":\"qwen2.5:3b\",\"system_fingerprint\":\"fp_ollama\",\"choices\":[{\"index\":0,\"delta\":{},\"finish_reason\":\"stop\"}]}\n\n" +
                "data: [DONE]\n\n";
            byte[] bytes = Encoding.UTF8.GetBytes(archivedEvents);
            Assert.AreEqual(433, bytes.Length);
            using (var hash = SHA256.Create())
                Assert.AreEqual("58DA676669B720AFB784C6222425F6375FC6A5A6DD14C6808C792E222DDDA82D",
                    BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", ""));

            var diagnostic = new StreamDiagnostics();
            using (var stream = new MemoryStream(bytes))
            {
                var result = await ChatStreamReader.ReadAsync(stream, false,
                    _ => Assert.Fail("The archived response contains no text fragment."),
                    CancellationToken.None, diagnostic);
                Assert.AreEqual("assistant", result["role"]);
                Assert.AreEqual("", result["content"]);
                Assert.IsFalse(result.ContainsKey("tool_calls"));
                Assert.AreEqual("complete-empty", diagnostic.Outcome);
                Assert.AreEqual(2, diagnostic.JsonChunks);
                Assert.AreEqual(0, diagnostic.TextChunks);
                Assert.AreEqual(0, diagnostic.ToolCallChunks);
                Assert.AreEqual(0, diagnostic.EmptyChoiceChunks);
                Assert.AreEqual(0, diagnostic.MissingDeltaChunks);
                Assert.AreEqual(0, diagnostic.UsageChunks);
                Assert.AreEqual("stop", diagnostic.TerminalReason);
                Assert.IsTrue(diagnostic.EndMarker);
            }
        }

        /// <summary>Recompose les arguments d’outils par index et émet les fragments de texte reçus.</summary>
        /// <returns>Tâche terminée après lecture du flux.</returns>
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

        /// <summary>Refuse un appel d’outil incomplet sans marqueur de fin ou avec une raison d’arrêt invalide.</summary>
        /// <returns>Tâche terminée après les vérifications asynchrones.</returns>
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

        /// <summary>Masque le corps d’erreur du fournisseur lorsqu’un flux signale une erreur.</summary>
        /// <returns>Tâche terminée après lecture et assertion.</returns>
        [TestMethod]
        public async Task ProviderErrorDoesNotLeakResponseBody()
        {
            using (var stream = Events(new { error = new { message = "secret-body" } }, "[DONE]"))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
                Assert.IsFalse(error.Message.Contains("secret-body"));
            }
        }

        /// <summary>Conserve l’entrée d’outil Claude et la signature de réflexion pendant le décodage.</summary>
        /// <returns>Tâche terminée après lecture du flux.</returns>
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

        /// <summary>Observe l’annulation déjà demandée avant de commencer la lecture.</summary>
        /// <returns>Tâche terminée après l’assertion d’annulation.</returns>
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
namespace VBAi.Tests.Unit
{
    /// <summary>Complète les cas limites de framing et de terminaison des flux.</summary>
    public sealed partial class StreamTests
    {
        /// <summary>Fusionne les fragments JSON, valeurs imbriquées et métadonnées sans perdre les scalaires.</summary>
        /// <returns>Tâche terminée après lecture du flux.</returns>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task OpenAiFramingMergesNullScalarsNestedValuesAndMetadata()
        {
            var wire = ": comment\nevent: update\n\ndata: {\"choices\":[{\"delta\":\ndata: {\"role\":\"assistant\",\"type\":\"message\",\"content\":null,\"count\":1,\"nested\":\"before\"}}]}\n\n";
            wire += System.Text.Encoding.UTF8.GetString(Events(
                new { choices = new object[0] },
                new { choices = new[] { new { delta = new { role = "assistant", type = "message", content = "", count = 2, nested = new { value = "first" } } } } },
                new { choices = new[] { new { delta = new { nested = new { value = "second" }, content = "text" } } } },
                new { choices = new[] { new { finish_reason = "stop" } } }, "[DONE]").ToArray());
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(wire)))
            {
                var result = await ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("assistant", result["role"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("message", result["type"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, result["count"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("firstsecond", Obj(result["nested"])["value"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("text", result["content"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(result.ContainsKey("tool_calls"));
            }
        }

        /// <summary>Refuse les raisons de fin invalides pour les deux protocoles et exige leur marqueur final.</summary>
        /// <returns>Tâche terminée après les cas OpenAI compatibles et Claude.</returns>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task StreamStopsRejectInvalidCompletionForBothProtocols()
        {
            foreach (var stop in new[] { null, "length", "tool_calls", "stop", "end_turn", "tool_use", "stop_sequence" })
            {
                using (var stream = Events(new { choices = new[] { new { delta = new { content = "ok" }, finish_reason = stop } } }, "[DONE]"))
                {
                    if (stop == "stop" || stop == "tool_calls")
                    {
                        var result = await ChatStreamReader.ReadAsync(stream, false, value => { }, System.Threading.CancellationToken.None);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("ok", result["content"]);
                    }
                    else await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None));
                }
                using (var stream = Events(new { type = "message_delta", delta = new { stop_reason = stop } }, new { type = "message_stop" }))
                {
                    if (stop == "end_turn" || stop == "tool_use" || stop == "stop_sequence")
                    {
                        var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", result["content"]);
                    }
                    else await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
                }
            }
            using (var stream = Events(new { type = "message_delta", delta = new { stop_reason = "end_turn" } }))
            {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
            }
            using (var stream = Events(new { type = "error", message = "private" }))
            {
                var error = await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.InvalidOperationException>(() => ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(error.Message.Contains("private"));
            }
        }

        /// <summary>Ignore les deltas inconnus et gère le texte initial ou vide sans fragment superflu.</summary>
        /// <returns>Tâche terminée après lecture des flux Claude.</returns>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ClaudeIgnoresUnknownDeltaAndHandlesEmptyTextWithoutProgress()
        {
            using (var stream = Events(
                new { type = "message_start" },
                new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "" } },
                new { type = "content_block_delta", index = 0, delta = new { type = "unknown", text = "ignored" } },
                new { type = "content_block_delta", index = 0, delta = new { type = "text_delta", text = "answer" } },
                new { type = "message_delta", delta = new { stop_reason = "end_turn" } }, new { type = "message_stop" }))
            {
                var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("answer", result["content"]);
            }
            using (var stream = Events(new { type = "content_block_start", index = 0, content_block = new { type = "text", text = "initial" } },
                new { type = "message_delta", delta = new { stop_reason = "stop_sequence" } }, new { type = "message_stop" }))
            {
                var result = await ChatStreamReader.ReadAsync(stream, true, null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("initial", result["content"]);
            }
        }

        /// <summary>Refuse une entrée trop volumineuse et une annulation déclenchée pendant la progression.</summary>
        /// <returns>Tâche terminée après vérification de la fermeture du flux.</returns>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task OversizeInputAndCancellationDuringProgressAreRejected()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(new string('x', 10 * 1024 * 1024 + 1) + "\n")))
            {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None));
            }
            using (var source = new System.Threading.CancellationTokenSource())
            using (var stream = Events(new { choices = new[] { new { delta = new { content = "first" } } } }, "[DONE]"))
            {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.OperationCanceledException>(() =>
                    ChatStreamReader.ReadAsync(stream, false, value => source.Cancel(), source.Token));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(stream.CanRead);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using VBAi;

    public sealed partial class StreamTests
    {
        [TestMethod]
        public async Task UnterminatedLinesAndBlankFramesStopAtTheWireByteBudget()
        {
            const int limit = 10 * 1024 * 1024;
            foreach (byte value in new[] { (byte)'x', (byte)'\n' })
                using (var stream = new GeneratedSseStream(value, limit + 8192))
                {
                    await Assert.ThrowsExceptionAsync<InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
                    Assert.IsTrue(stream.BytesRead <= limit + 1, "The reader must stop at the limit before consuming an oversized line or ignored blank frames.");
                    Assert.IsTrue(stream.Disposed);
                }
        }

        [TestMethod]
        public async Task Utf8ResponseAtExactWireLimitCompletesAndExtraByteIsRejected()
        {
            const int limit = 10 * 1024 * 1024;
            byte[] finish = Events(new { choices = new[] { new { delta = new { content = "été漢字" }, finish_reason = "stop" } } }, "[DONE]").ToArray();
            foreach (int extra in new[] { 0, 1 })
            {
                // An ignored SSE comment pads the response without enlarging its JSON payload.
                byte[] wire = new byte[limit + extra];
                wire[0] = (byte)':';
                for (int i = 1; i < wire.Length - finish.Length - 1; i++) wire[i] = (byte)'x';
                wire[wire.Length - finish.Length - 1] = (byte)'\n';
                Buffer.BlockCopy(finish, 0, wire, wire.Length - finish.Length, finish.Length);
                using (var stream = new MemoryStream(wire))
                {
                    if (extra != 0)
                        await Assert.ThrowsExceptionAsync<InvalidDataException>(() => ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None));
                    else
                        Assert.AreEqual("été漢字", (await ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None))["content"]);
                }
            }
        }

        [TestMethod]
        public async Task ManyFragmentsRemainPlainStringsAndPreserveRepeatedMetadata()
        {
            var frames = new List<object>();
            var expected = new StringBuilder();
            for (int i = 0; i < 4000; i++)
            {
                string part = i % 2 == 0 ? "é" : "漢";
                expected.Append(part);
                frames.Add(new
                {
                    choices = new[] { new { delta = new { role = "assistant", type = "message", content = part,
                    tool_calls = new[] { new { index = 0, type = "function", function = new { arguments = part } } } } } }
                });
            }
            frames.Add(new { choices = new[] { new { finish_reason = "tool_calls" } } }); frames.Add("[DONE]");
            using (var stream = Events(frames.ToArray()))
            {
                var result = await ChatStreamReader.ReadAsync(stream, false, null, CancellationToken.None);
                Assert.AreEqual(expected.ToString(), result["content"]);
                Assert.AreEqual("assistant", result["role"]); Assert.AreEqual("message", result["type"]);
                var call = Obj(((object[])result["tool_calls"])[0]);
                Assert.AreEqual("function", call["type"]);
                Assert.AreEqual(expected.ToString(), Obj(call["function"])["arguments"]);
                Assert.IsTrue(Json.Serialize(result).Contains("arguments"));
            }
        }

        private sealed class GeneratedSseStream : Stream
        {
            private readonly byte value;
            private readonly int length;
            internal int BytesRead;
            internal bool Disposed;
            internal GeneratedSseStream(byte value, int length) { this.value = value; this.length = length; }
            public override bool CanRead => !Disposed;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => length;
            public override long Position { get => BytesRead; set => throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count)
            {
                int read = Math.Min(count, length - BytesRead);
                for (int i = 0; i < read; i++) buffer[offset + i] = value;
                BytesRead += read; return read;
            }
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            { cancellationToken.ThrowIfCancellationRequested(); return Task.FromResult(Read(buffer, offset, count)); }
            protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
            public override void Flush() => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class StreamTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task StreamDiagnosticsDistinguishTerminalEmptyStatisticsToolsAndProtocolFailure()
        {
            var diagnostic = new VBAi.StreamDiagnostics();
            using (var stream = Events(new { choices = new object[0], usage = new { completion_tokens = 0 } },
                new { choices = new[] { new { delta = new { content = "" }, finish_reason = "stop" } } }, "[DONE]"))
            {
                var result = await VBAi.ChatStreamReader.ReadAsync(stream, false, _ => Microsoft.VisualStudio.TestTools.UnitTesting.Assert.Fail("Empty text must not publish fragments."), System.Threading.CancellationToken.None, diagnostic);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", result["content"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("complete-empty", diagnostic.Outcome);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, diagnostic.JsonChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, diagnostic.EmptyChoiceChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, diagnostic.UsageChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, diagnostic.TextChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("stop", diagnostic.TerminalReason);
                var snapshot = diagnostic.Snapshot();
                diagnostic.SetTerminalReason("secret-reason");
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("stop", snapshot.TerminalReason);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(diagnostic.EndMarker);
            }
            diagnostic = new VBAi.StreamDiagnostics();
            using (var stream = Events(new { choices = new[] { new { delta = new { tool_calls = new[] { new { index = 0, id = "secret-id", function = new { name = "secret-tool", arguments = "secret-arguments" } } } }, finish_reason = "tool_calls" } } }, "[DONE]"))
            {
                var result = await VBAi.ChatStreamReader.ReadAsync(stream, false, _ => Microsoft.VisualStudio.TestTools.UnitTesting.Assert.Fail("Tool-only response must not publish text."), System.Threading.CancellationToken.None, diagnostic);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, ((object[])result["tool_calls"]).Length);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("complete-tools", diagnostic.Outcome);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, diagnostic.ToolCallChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, diagnostic.TextChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(diagnostic).Contains("secret"));
            }
            diagnostic = new VBAi.StreamDiagnostics();
            using (var stream = Events(new { choices = new[] { new { message = new { content = "secret-body" }, finish_reason = "secret-reason" } } }, "[DONE]"))
            {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.InvalidDataException>(() => VBAi.ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None, diagnostic));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("protocol-error", diagnostic.Outcome);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("unknown", diagnostic.TerminalReason);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, diagnostic.MissingDeltaChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(diagnostic).Contains("secret"));
            }
            diagnostic = new VBAi.StreamDiagnostics();
            using (var stream = Events(new { choices = new[] { new { delta = new { content = "secret-text" }, finish_reason = "stop" } } }, "[DONE]"))
            {
                var result = await VBAi.ChatStreamReader.ReadAsync(stream, false, null, System.Threading.CancellationToken.None, diagnostic);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("secret-text", result["content"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("complete-text", diagnostic.Outcome);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, diagnostic.TextChunks);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(diagnostic).Contains("secret"));
            }
        }
    }
}
