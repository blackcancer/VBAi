using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal static class ChatStreamReader
    {
        private static IDictionary<string, object> Obj(object value) { return ClaudeProtocol.Object(value); }
        private static string Text(IDictionary<string, object> value, string key) { return ClaudeProtocol.Text(value, key); }

        public static async Task<IDictionary<string, object>> ReadAsync(Stream stream, bool claude, Action<string> progress, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
            var message = new Dictionary<string, object> { ["role"] = "assistant", ["content"] = "" };
            var calls = new SortedDictionary<int, IDictionary<string, object>>();
            var blocks = new SortedDictionary<int, IDictionary<string, object>>();
            var inputs = new Dictionary<int, string>();
            var data = new StringBuilder(); int total = 0; string stop = null; bool ended = false;
            using (token.Register(() => stream.Dispose()))
            using (var reader = new StreamReader(stream, Encoding.UTF8)) {
                while (!ended) {
                    token.ThrowIfCancellationRequested();
                    string line = await reader.ReadLineAsync();
                    if (line == null) break;
                    total += line.Length;
                    if (total > 10 * 1024 * 1024) throw new InvalidDataException(UiText.Get("Response too large."));
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
                            if (kind == "input_json_delta") { string previous; inputs.TryGetValue(index, out previous); inputs[index] = previous + Text(delta, "partial_json"); }
                            else if (kind == "text_delta") { string text = Text(delta, "text"); block["text"] = Text(block, "text") + text; progress?.Invoke(text); }
                            else if (kind == "thinking_delta") block["thinking"] = Text(block, "thinking") + Text(delta, "thinking");
                            else if (kind == "signature_delta") block["signature"] = Text(block, "signature") + Text(delta, "signature");
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
                foreach (var pair in inputs) blocks[pair.Key]["input"] = json.DeserializeObject(pair.Value);
                return ClaudeProtocol.Response(new Dictionary<string, object> { ["content"] = blocks.Values.Cast<object>().ToArray(), ["stop_reason"] = stop });
            }
            if (calls.Count > 0) message["tool_calls"] = calls.Values.Cast<object>().ToArray();
            return message;
        }

        private static void Merge(IDictionary<string, object> target, IDictionary<string, object> delta)
        {
            foreach (var pair in delta) {
                if (pair.Value == null) continue;
                object previous; target.TryGetValue(pair.Key, out previous);
                var child = pair.Value as IDictionary<string, object>;
                if (child != null) { var into = previous as IDictionary<string, object> ?? new Dictionary<string, object>(); Merge(into, child); target[pair.Key] = into; }
                else if (pair.Value is string && previous is string && pair.Key != "role" && pair.Key != "type") target[pair.Key] = (string)previous + (string)pair.Value;
                else target[pair.Key] = pair.Value;
            }
        }
    }
}
