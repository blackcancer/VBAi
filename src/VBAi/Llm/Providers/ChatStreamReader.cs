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
        /// <returns>Message assistant assemblé avec ses appels d’outils complets.</returns>
        public static async Task<IDictionary<string, object>> ReadAsync(Stream stream, bool claude, Action<string> progress, CancellationToken token)
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
                    if (payload == "[DONE]") { ended = true; break; }
                    var root = Obj(json.DeserializeObject(payload));
                    if (root.ContainsKey("error") || Text(root, "type") == "error") throw new InvalidOperationException(UiText.Get("The provider interrupted the response with an error."));
                    if (!claude) {
                        var choices = ClaudeProtocol.Array(root, "choices"); if (choices.Length == 0) continue;
                        var choice = Obj(choices[0]);
                        if (Text(choice, "finish_reason") != null) stop = Text(choice, "finish_reason");
                        if (!choice.ContainsKey("delta")) continue;
                        var delta = Obj(choice["delta"]);
                        foreach (var rawCall in ClaudeProtocol.Array(delta, "tool_calls")) {
                            var call = Obj(rawCall); int index = Convert.ToInt32(call["index"]); call.Remove("index");
                            IDictionary<string, object> target;
                            if (!calls.TryGetValue(index, out target)) calls[index] = target = new Dictionary<string, object>();
                            Merge(target, call);
                        }
                        delta.Remove("tool_calls"); string text = Text(delta, "content"); Merge(message, delta);
                        if (!string.IsNullOrEmpty(text)) progress?.Invoke(text);
                    } else {
                        string type = Text(root, "type");
                        if (type == "content_block_start") {
                            int index = Convert.ToInt32(root["index"]); blocks[index] = Obj(root["content_block"]);
                            if (Text(blocks[index], "type") == "text") { string initial = Text(blocks[index], "text"); if (!string.IsNullOrEmpty(initial)) progress?.Invoke(initial); }
                        }
                        if (type == "content_block_delta") {
                            int index = Convert.ToInt32(root["index"]); var block = blocks[index]; var delta = Obj(root["delta"]); string kind = Text(delta, "type");
                            if (kind == "input_json_delta") { if (!inputs.TryGetValue(index, out var input)) inputs[index] = input = new StringBuilder(); input.Append(Text(delta, "partial_json")); }
                            else if (kind == "text_delta") { string text = Text(delta, "text"); Append(block, "text", text); progress?.Invoke(text); }
                            else if (kind == "thinking_delta") Append(block, "thinking", Text(delta, "thinking"));
                            else if (kind == "signature_delta") Append(block, "signature", Text(delta, "signature"));
                        }
                        if (type == "message_delta") stop = Text(Obj(root["delta"]), "stop_reason");
                        if (type == "message_stop") ended = true;
                    }
                }
            }
            token.ThrowIfCancellationRequested();
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
        private static void Append(IDictionary<string, object> target, string key, string text)
        {
            target.TryGetValue(key, out var previous);
            var buffer = previous as StringBuilder;
            if (buffer == null) target[key] = buffer = new StringBuilder(previous as string ?? "");
            buffer.Append(text);
        }

        /// <summary>Publishes ordinary strings only after the provider has completed the message.</summary>
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
            private const int Limit = 10 * 1024 * 1024;
            private readonly Stream inner;
            private int received;
            internal BoundedStream(Stream inner) { this.inner = inner; }
            public override bool CanRead => inner.CanRead;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count)
            {
                return Count(inner.Read(buffer, offset, Math.Min(count, Limit - received + 1)));
            }
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                return Count(await inner.ReadAsync(buffer, offset, Math.Min(count, Limit - received + 1), cancellationToken).ConfigureAwait(false));
            }
            private int Count(int count)
            {
                received += count;
                if (received > Limit) throw new InvalidDataException(UiText.Get("Response too large."));
                return count;
            }
            protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }
            public override void Flush() => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
