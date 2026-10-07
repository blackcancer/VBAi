using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace VBAi
{
    /// <summary>Projects only explicitly published provider reasoning text into request-scoped activity cards.</summary>
    internal sealed class ProviderReasoning
    {
        private sealed class Section
        {
            internal string Id;
            internal readonly StringBuilder Text = new StringBuilder();
            internal bool Published;
        }

        private readonly Action<CodexAgentActivity> publish;
        private readonly string prefix;
        private readonly Dictionary<string, Section> sections = new Dictionary<string, Section>(StringComparer.Ordinal);
        private readonly List<Section> order = new List<Section>();
        private readonly HashSet<string> structuredSources = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> structuredSections = new HashSet<string>(StringComparer.Ordinal);
        private bool terminal;

        internal ProviderReasoning(Action<CodexAgentActivity> publish, string prefix = null)
        {
            this.publish = publish;
            this.prefix = prefix ?? Guid.NewGuid().ToString("N");
        }

        /// <summary>Recognizes documented public reasoning details before their mirrored compatibility fields.</summary>
        internal void OpenAi(IDictionary<string, object> value, bool snapshot = false, string source = "openai")
        {
            if (value == null) return;
            bool structured = false;
            int ordinal = 0;
            foreach (object raw in ClaudeProtocol.Array(value, "reasoning_details"))
            {
                var detail = raw as IDictionary<string, object>;
                if (detail == null) { ordinal++; continue; }
                string type = ClaudeProtocol.Text(detail, "type");
                string text = type == "reasoning.text" ? PublicString(detail, "text") :
                    type == "reasoning.summary" ? PublicString(detail, "summary") : null;
                if (!string.IsNullOrEmpty(text))
                {
                    structured = true;
                    // A stream may omit its server ID after the initial fragment.
                    // The documented index (or response-array position) keeps the section stable.
                    string identity = detail.TryGetValue("index", out var index) && index is int ? Convert.ToString(index, CultureInfo.InvariantCulture) : ordinal.ToString(CultureInfo.InvariantCulture);
                    string sectionKey = source + ":section:" + identity;
                    bool switchedRepresentation = structuredSections.Add(sectionKey) && sections.ContainsKey(sectionKey);
                    // Structured public details replace an earlier compatibility mirror
                    // for this declared index; they never create a duplicate card.
                    Put(sectionKey, text, snapshot || switchedRepresentation);
                }
                ordinal++;
            }
            if (structured) { structuredSources.Add(source); return; }
            if (structuredSources.Contains(source)) return;
            foreach (string field in new[] { "reasoning_content", "reasoning", "reasoning_text" })
            {
                string text = PublicString(value, field);
                if (!string.IsNullOrEmpty(text)) { Put(source + ":section:0", text, snapshot); break; }
            }
        }

        internal void ClaudeBlock(int index, IDictionary<string, object> block, bool snapshot)
        {
            if (block == null || ClaudeProtocol.Text(block, "type") != "thinking" ||
                (block.TryGetValue("redacted", out var redacted) && Equals(redacted, true))) return;
            Put("claude:block:" + index.ToString(CultureInfo.InvariantCulture), PublicString(block, "thinking"), snapshot);
        }

        internal void ClaudeDelta(int index, IDictionary<string, object> block, IDictionary<string, object> delta)
        {
            if (block == null || delta == null || ClaudeProtocol.Text(block, "type") == "redacted_thinking" ||
                (block.TryGetValue("redacted", out var redacted) && Equals(redacted, true))) return;
            if (ClaudeProtocol.Text(block, "type") == "thinking" && ClaudeProtocol.Text(delta, "type") == "thinking_delta")
                Put("claude:block:" + index.ToString(CultureInfo.InvariantCulture), PublicString(delta, "thinking"), false);
        }

        internal void ClaudeResponse(IDictionary<string, object> response)
        {
            int index = 0;
            foreach (object raw in ClaudeProtocol.Array(response, "content")) ClaudeBlock(index++, ClaudeProtocol.Object(raw), true);
        }

        internal void BedrockResponse(IDictionary<string, object> response)
        {
            var output = response != null && response.TryGetValue("output", out var rawOutput) ? ClaudeProtocol.Object(rawOutput) : null;
            var message = output != null && output.TryGetValue("message", out var rawMessage) ? ClaudeProtocol.Object(rawMessage) : null;
            int index = 0;
            foreach (object raw in ClaudeProtocol.Array(message, "content"))
            {
                var block = ClaudeProtocol.Object(raw);
                var reasoning = block != null && block.TryGetValue("reasoningContent", out var rawReasoning) ? ClaudeProtocol.Object(rawReasoning) : null;
                if (reasoning != null && !reasoning.ContainsKey("redactedContent") && reasoning.TryGetValue("reasoningText", out var rawText) && ClaudeProtocol.Object(rawText) is IDictionary<string, object> text)
                    Put("bedrock:block:" + index.ToString(CultureInfo.InvariantCulture), PublicString(text, "text"), true);
                index++;
            }
        }

        private static string PublicString(IDictionary<string, object> value, string key) =>
            value != null && value.TryGetValue(key, out var raw) ? raw as string : null;

        private void Put(string key, string text, bool snapshot)
        {
            if (terminal || publish == null || text == null) return;
            if (!sections.TryGetValue(key, out var section))
            {
                if (text.Length == 0) return;
                section = new Section { Id = prefix + ":reasoning:" + order.Count.ToString(CultureInfo.InvariantCulture) };
                sections.Add(key, section); order.Add(section);
            }
            bool wasPublished = section.Published;
            if (snapshot) section.Text.Clear();
            section.Text.Append(text);
            if (!wasPublished && string.IsNullOrWhiteSpace(section.Text.ToString())) return;
            section.Published = true;
            publish(new CodexAgentActivity { Id = section.Id, Kind = "reasoning", Title = UiText.Get("Reasoning"),
                Detail = snapshot || !wasPublished ? section.Text.ToString() : text, Status = "inProgress", Append = !snapshot });
        }

        /// <summary>Completes only after a valid whole response; failure markers retain already published text without repeating it.</summary>
        internal void Finish(string status)
        {
            if (terminal) return;
            terminal = true;
            foreach (var section in order.Where(s => s.Published))
                publish?.Invoke(new CodexAgentActivity { Id = section.Id, Kind = "reasoning", Title = UiText.Get("Reasoning"),
                    Detail = status == "completed" ? section.Text.ToString() : "", Status = status, Append = status != "completed" });
        }
    }
}
