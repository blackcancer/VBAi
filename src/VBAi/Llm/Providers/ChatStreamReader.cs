using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Retains fixed-size protocol metadata, never provider content or tool arguments.</summary>
    internal sealed class StreamDiagnostics
    {

        /// <summary>Gets or sets the json chunks.</summary>
        /// <value>Current json chunks exposed by stream diagnostics.</value>
        public int JsonChunks { get; internal set; }

        /// <summary>Gets or sets the empty choice chunks.</summary>
        /// <value>Current empty choice chunks exposed by stream diagnostics.</value>
        public int EmptyChoiceChunks { get; internal set; }

        /// <summary>Gets or sets the usage chunks.</summary>
        /// <value>Current usage chunks exposed by stream diagnostics.</value>
        public int UsageChunks { get; internal set; }

        /// <summary>Gets or sets the missing delta chunks.</summary>
        /// <value>Current missing delta chunks exposed by stream diagnostics.</value>
        public int MissingDeltaChunks { get; internal set; }

        /// <summary>Gets or sets the text chunks.</summary>
        /// <value>Current text chunks exposed by stream diagnostics.</value>
        public int TextChunks { get; internal set; }

        /// <summary>Gets or sets the tool call chunks.</summary>
        /// <value>Current tool call chunks exposed by stream diagnostics.</value>
        public int ToolCallChunks { get; internal set; }

        /// <summary>Gets or sets the end marker.</summary>
        /// <value>Current end marker exposed by stream diagnostics.</value>
        public bool EndMarker { get; internal set; }

        /// <summary>Gets or sets the terminal reason.</summary>
        /// <value>Current terminal reason exposed by stream diagnostics.</value>
        public string TerminalReason { get; private set; } = "missing";

        /// <summary>Gets or sets the outcome.</summary>
        /// <value>Current outcome exposed by stream diagnostics.</value>
        public string Outcome { get; internal set; } = "reading";

        /// <summary>Freezes the current scalar metadata for a point-in-time observation.</summary>
        /// <returns>stream diagnostics produced by the operation for snapshot on stream diagnostics.</returns>
        internal StreamDiagnostics Snapshot() { return (StreamDiagnostics)MemberwiseClone(); }

        /// <summary>Stores only recognized provider stop reasons; other values become <c>unknown</c>.</summary>
        /// <param name="reason">OpenAI- or Claude-style terminal reason received from the stream, or null when absent.</param>
        internal void SetTerminalReason(string reason)
        {
            switch (reason)
            {
                case null: TerminalReason = "missing"; break;
                case "stop": case "tool_calls": case "length": case "content_filter":
                case "end_turn": case "tool_use": case "stop_sequence": TerminalReason = reason; break;
                default: TerminalReason = "unknown"; break;
            }
        }
    }

    /// <summary>Lit les événements SSE des fournisseurs et assemble une réponse complète avant son utilisation.</summary>
    internal static class ChatStreamReader
    {

        /// <summary>Convertit un événement décodé en dictionnaire de propriétés.</summary>
        /// <param name="value">Valeur de l’événement à convertir.</param>
        /// <returns>Dictionnaire de propriétés, ou null si la valeur est incompatible.</returns>
        private static IDictionary<string, object> Obj(object value) { return ClaudeProtocol.Object(value); }

        /// <summary>Lit une propriété textuelle d’un événement.</summary>
        /// <param name="value">Dictionnaire de l’événement.</param>
        /// <param name="key">Nom de la propriété.</param>
        /// <returns>Texte de la propriété, ou null si elle est absente.</returns>
        private static string Text(IDictionary<string, object> value, string key) { return ClaudeProtocol.Text(value, key); }

        /// <summary>Lit le flux SSE, assemble les fragments et refuse les réponses incomplètes avant de retourner le message.</summary>
        /// <param name="stream">Flux de réponse du fournisseur.</param>
        /// <param name="claude">Vrai lorsque le flux suit le protocole Claude.</param>
        /// <param name="progress">Callback appelé pour chaque fragment textuel reçu, éventuellement null.</param>
        /// <param name="token">Jeton qui annule la lecture et ferme le flux.</param>
        /// <param name="diagnostics">Optional bounded metadata receiver, excluding provider content.</param>
        /// <returns>Message assistant assemblé avec ses appels d’outils complets.</returns>
        public static async Task<IDictionary<string, object>> ReadAsync(Stream stream, bool claude, Action<string> progress, CancellationToken token, StreamDiagnostics diagnostics = null)
        {
            diagnostics = diagnostics ?? new StreamDiagnostics();
            try
            {
                var result = await ReadCoreAsync(stream, claude, progress, token, diagnostics);
                diagnostics.Outcome = diagnostics.ToolCallChunks > 0 ? "complete-tools" :
                    string.IsNullOrWhiteSpace(Convert.ToString(result["content"])) ? "complete-empty" : "complete-text";
                return result;
            }
            catch (Exception error)
            {
                diagnostics.Outcome = token.IsCancellationRequested ? "cancelled" :
                    error is IOException && !(error is InvalidDataException) ? "transport-error" : "protocol-error";
                throw;
            }
        }

        /// <summary>Parses bounded SSE events, assembles text/tool fragments, and rejects incomplete or filtered turns.</summary>
        /// <param name="stream">Provider response stream wrapped in a byte-limiting stream.</param>
        /// <param name="claude">Selects Claude event blocks when true; otherwise parses OpenAI-style choices/deltas.</param>
        /// <param name="progress">Optional callback for text fragments only; tool arguments are never sent through it.</param>
        /// <param name="token">Cancellation token; cancellation disposes the stream and prevents partial tool execution.</param>
        /// <param name="diagnostics">Scalar-only counters updated as events are parsed.</param>
        /// <returns>Complete assistant message with assembled tool calls only after a valid terminal marker/reason.</returns>
        private static async Task<IDictionary<string, object>> ReadCoreAsync(Stream stream, bool claude, Action<string> progress, CancellationToken token, StreamDiagnostics diagnostics)
        {
            token.ThrowIfCancellationRequested();
            var json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
            var message = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = "" };
            var calls = new SortedDictionary<int, IDictionary<string, object>>();
            var blocks = new SortedDictionary<int, IDictionary<string, object>>();
            var inputs = new Dictionary<int, StringBuilder>();
            var data = new StringBuilder(); string stop = null; bool ended = false;
            using (token.Register(() => stream.Dispose()))
            using (var reader = new StreamReader(new BoundedStream(stream), Encoding.UTF8)) {
                while (!ended) {
                    token.ThrowIfCancellationRequested();
                    string line = await reader.ReadLineAsync();
                    if (line == null) break;
                    if (line.StartsWith("data:", StringComparison.Ordinal)) { if (data.Length > 0) data.Append('\n'); data.Append(line.Substring(5).TrimStart(' ')); continue; }
                    if (line.Length != 0 || data.Length == 0) continue;
                    string payload = data.ToString(); data.Clear();
                    if (payload == "[DONE]") { ended = true; diagnostics.EndMarker = true; break; }
                    diagnostics.JsonChunks++;
                    var root = Obj(json.DeserializeObject(payload));
                    if (root.ContainsKey("error") || Text(root, "type") == "error") throw new InvalidOperationException(UiText.Get("The provider interrupted the response with an error."));
                    if (!claude) {
                        var choices = ClaudeProtocol.Array(root, "choices"); if (choices.Length == 0) { diagnostics.EmptyChoiceChunks++; if (root.ContainsKey("usage")) diagnostics.UsageChunks++; continue; }
                        var choice = Obj(choices[0]);
                        if (Text(choice, "finish_reason") != null) { stop = Text(choice, "finish_reason"); diagnostics.SetTerminalReason(stop); }
                        if (!choice.ContainsKey("delta")) { diagnostics.MissingDeltaChunks++; continue; }
                        var delta = Obj(choice["delta"]);
                        if (ClaudeProtocol.Array(delta, "tool_calls").Length > 0) diagnostics.ToolCallChunks++;
                        foreach (var rawCall in ClaudeProtocol.Array(delta, "tool_calls")) {
                            var call = Obj(rawCall); int index = Convert.ToInt32(call["index"]); call.Remove("index");
                            IDictionary<string, object> target;
                            if (!calls.TryGetValue(index, out target)) calls[index] = target = new Dictionary<string, object>();
                            Merge(target, call);
                        }
                        delta.Remove("tool_calls"); string text = Text(delta, "content"); Merge(message, delta);
                        if (!string.IsNullOrEmpty(text)) { diagnostics.TextChunks++; progress?.Invoke(text); }
                    } else {
                        string type = Text(root, "type");
                        if (type == "content_block_start") {
                            int index = Convert.ToInt32(root["index"]); blocks[index] = Obj(root["content_block"]);
                            if (Text(blocks[index], "type") == "tool_use") diagnostics.ToolCallChunks++;
                            if (Text(blocks[index], "type") == "text") { string initial = Text(blocks[index], "text"); if (!string.IsNullOrEmpty(initial)) { diagnostics.TextChunks++; progress?.Invoke(initial); } }
                        }
                        if (type == "content_block_delta") {
                            int index = Convert.ToInt32(root["index"]); var block = blocks[index]; var delta = Obj(root["delta"]); string kind = Text(delta, "type");
                            if (kind == "input_json_delta") { if (!inputs.TryGetValue(index, out var input)) inputs[index] = input = new StringBuilder(); input.Append(Text(delta, "partial_json")); }
                            else if (kind == "text_delta") { string text = Text(delta, "text"); Append(block, "text", text); if (!string.IsNullOrEmpty(text)) diagnostics.TextChunks++; progress?.Invoke(text); }
                            else if (kind == "thinking_delta") Append(block, "thinking", Text(delta, "thinking"));
                            else if (kind == "signature_delta") Append(block, "signature", Text(delta, "signature"));
                        }
                        if (type == "message_delta") { stop = Text(Obj(root["delta"]), "stop_reason"); diagnostics.SetTerminalReason(stop); }
                        if (type == "message_stop") { ended = true; diagnostics.EndMarker = true; }
                    }
                }
            }
            token.ThrowIfCancellationRequested();
            diagnostics.SetTerminalReason(stop);
            bool validStop = claude ? stop == "end_turn" || stop == "tool_use" || stop == "stop_sequence" : stop == "stop" || stop == "tool_calls";
            if (!ended || !validStop) throw new InvalidDataException(UiText.Get("Response interrupted, truncated or filtered; no partial tool call was executed."));
            if (claude) {
                foreach (var pair in inputs) blocks[pair.Key]["input"] = json.DeserializeObject(pair.Value.ToString());
                foreach (var block in blocks.Values) FinalizeText(block);
                return ClaudeProtocol.Response(new Dictionary<string, object> { ["content"] = blocks.Values.Cast<object>().ToArray(), ["stop_reason"] = stop });
            }
            FinalizeText(message);
            foreach (var call in calls.Values) FinalizeText(call);
            if (calls.Count > 0) message["tool_calls"] = calls.Values.Cast<object>().ToArray();
            return message;
        }

        /// <summary>Reads a non-streamed provider body under the same byte and cancellation bounds.</summary>
        /// <param name="stream">Provider body stream wrapped in the same byte limit as streaming responses.</param>
        /// <param name="token">Token used to cancel the operation.</param>
        /// <returns>Complete response body text, or cancellation/transport/size failure.</returns>
        internal static async Task<string> ReadBodyAsync(Stream stream, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            using (token.Register(() => stream.Dispose()))
            using (var reader = new StreamReader(new BoundedStream(stream), Encoding.UTF8))
            {
                string body = await reader.ReadToEndAsync().ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                return body;
            }
        }

        /// <summary>Fusionne un fragment dans le message en concaténant les chaînes et en fusionnant récursivement les objets.</summary>
        /// <param name="target">Dictionnaire qui reçoit les valeurs fusionnées.</param>
        /// <param name="delta">Dictionnaire contenant le nouveau fragment.</param>
        private static void Merge(IDictionary<string, object> target, IDictionary<string, object> delta)
        {
            foreach (var pair in delta) {
                if (pair.Value == null) continue;
                object previous; target.TryGetValue(pair.Key, out previous);
                var child = pair.Value as IDictionary<string, object>;
                if (child != null) { var into = previous as IDictionary<string, object> ?? new Dictionary<string, object>(); Merge(into, child); target[pair.Key] = into; }
                else if (pair.Value is string text && pair.Key != "role" && pair.Key != "type") Append(target, pair.Key, text);
                else target[pair.Key] = pair.Value;
            }
        }

        /// <summary>Accumulates token fragments without copying the complete prefix on every delta.</summary>
        /// <param name="target">Message or content-block dictionary receiving the accumulated fragment.</param>
        /// <param name="key">Protocol field name whose prior text should be extended.</param>
        /// <param name="text">New fragment appended without reconstructing the accumulated prefix.</param>
        private static void Append(IDictionary<string, object> target, string key, string text)
        {
            target.TryGetValue(key, out var previous);
            var buffer = previous as StringBuilder;
            if (buffer == null) target[key] = buffer = new StringBuilder(previous as string ?? "");
            buffer.Append(text);
        }

        /// <summary>Publishes ordinary strings only after the provider has completed the message.</summary>
        /// <param name="value">Message tree whose buffered StringBuilder values are converted to ordinary strings in place.</param>
        private static void FinalizeText(IDictionary<string, object> value)
        {
            foreach (var key in value.Keys.ToArray())
            {
                if (value[key] is StringBuilder buffer) value[key] = buffer.ToString();
                else if (value[key] is IDictionary<string, object> child) FinalizeText(child);
            }
        }

        /// <summary>Bounds network bytes before StreamReader can accumulate an unterminated line.</summary>
        private sealed class BoundedStream : Stream
        {

            /// <summary>Maximum provider response bytes accepted before an InvalidDataException is raised.</summary>
            private const int Limit = 10 * 1024 * 1024;

            /// <summary>Provider response stream wrapped by the byte-counting boundary.</summary>
            private readonly Stream inner;

            /// <summary>Total bytes delivered so far, including one extra byte used to detect limit overflow.</summary>
            private int received;

            /// <summary>Wraps a response stream and starts its byte count at zero.</summary>
            /// <param name="inner">Underlying provider stream; disposal of this wrapper also disposes it.</param>
            internal BoundedStream(Stream inner) { this.inner = inner; }

            /// <summary>Gets the can read.</summary>
            /// <value>Current can read exposed by bounded stream.</value>
            public override bool CanRead => inner.CanRead;

            /// <summary>Gets the can seek.</summary>
            /// <value>Current can seek exposed by bounded stream.</value>
            public override bool CanSeek => false;

            /// <summary>Gets the can write.</summary>
            /// <value>Current can write exposed by bounded stream.</value>
            public override bool CanWrite => false;

            /// <summary>Gets the length.</summary>
            /// <value>Current length exposed by bounded stream.</value>
            public override long Length => throw new NotSupportedException();

            /// <summary>Gets or sets the position.</summary>
            /// <value>Current position exposed by bounded stream.</value>
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

            /// <summary>Reads no more than the remaining byte budget plus one overflow-detection byte.</summary>
            /// <param name="buffer">Destination buffer.</param>
            /// <param name="offset">Buffer offset where bytes are written.</param>
            /// <param name="count">Requested maximum bytes for this read.</param>
            /// <returns>Bytes read, unless the cumulative response exceeds the configured limit.</returns>
            public override int Read(byte[] buffer, int offset, int count)
            {
                return Count(inner.Read(buffer, offset, Math.Min(count, Limit - received + 1)));
            }

            /// <summary>Asynchronously reads within the remaining byte budget plus one overflow-detection byte.</summary>
            /// <param name="buffer">Destination buffer.</param>
            /// <param name="offset">Buffer offset where bytes are written.</param>
            /// <param name="count">Requested maximum bytes for this read.</param>
            /// <param name="cancellationToken">Token forwarded to the underlying stream read.</param>
            /// <returns>Bytes read, unless the cumulative response exceeds the configured limit.</returns>
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                return Count(await inner.ReadAsync(buffer, offset, Math.Min(count, Limit - received + 1), cancellationToken).ConfigureAwait(false));
            }

            /// <summary>Adds bytes from one underlying read and rejects cumulative response overflow.</summary>
            /// <param name="count">Number of bytes returned by the underlying stream.</param>
            /// <returns>The same count when the total remains within the 10 MiB limit.</returns>
            private int Count(int count)
            {
                received += count;
                if (received > Limit) throw new InvalidDataException(UiText.Get("Response too large."));
                return count;
            }

            /// <summary>Disposes the wrapped provider stream when disposing this wrapper.</summary>
            /// <param name="disposing">True when called by Dispose rather than finalization.</param>
            protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }

            /// <summary>Throws because the provider response wrapper is read-only.</summary>
            public override void Flush() => throw new NotSupportedException();

            /// <summary>Throws because provider response streams are consumed forward-only.</summary>
            /// <param name="offset">Unused requested seek offset.</param>
            /// <param name="origin">Unused seek origin.</param>
            /// <returns>No value; this method always throws NotSupportedException.</returns>
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            /// <summary>Throws because the provider response wrapper cannot be resized.</summary>
            /// <param name="value">Unused requested length.</param>
            public override void SetLength(long value) => throw new NotSupportedException();

            /// <summary>Throws because provider response streams are read-only.</summary>
            /// <param name="buffer">Unused source buffer.</param>
            /// <param name="offset">Unused source offset.</param>
            /// <param name="count">Unused byte count.</param>
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
