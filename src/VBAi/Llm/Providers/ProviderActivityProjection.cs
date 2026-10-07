using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Projects public provider events into the existing activity/session shape.</summary>
    internal static class ProviderActivityProjection
    {
        /// <summary>Creates a tool timeline entry from complete arguments without exposing code or secrets.</summary>
        /// <param name="id">Stable local call identity.</param>
        /// <param name="name">Declared tool or catalogue gateway name.</param>
        /// <param name="arguments">Complete JSON argument object.</param>
        /// <param name="status">Observed tool lifecycle state.</param>
        /// <param name="error">Public result diagnostic, when supplied.</param>
        /// <returns>Activity retaining the technical identifier in its selectable details.</returns>
        internal static CodexAgentActivity Tool(string id, string name, string arguments, string status, string error = null)
        {
            var values = ReadObject(arguments);
            string declared = name;
            if (name == "invoke_tool" && values != null && values.Count == 2 &&
                values.TryGetValue("ToolName", out var target) && target is string tool &&
                values.TryGetValue("ArgumentsJson", out var inner) && inner is string nested &&
                !LlmVbeTools.IsCatalogTool(tool))
            {
                var parsed = ReadObject(nested);
                if (parsed != null) { name = tool; values = parsed; }
            }
            var activity = CodexAgentActivity.FromItem(new Dictionary<string, object>
            {
                ["id"] = id, ["type"] = "dynamicToolCall", ["tool"] = name, ["arguments"] = values, ["status"] = status
            }, status != "inProgress");
            if (declared != name) activity.Detail = declared + " → " + activity.Detail;
            if (!string.IsNullOrWhiteSpace(error)) activity.Detail = CodexAgentActivity.Limit(activity.Detail + "\n" + error);
            return activity;
        }

        /// <summary>Scopes a transport-local identity without modifying an event retained by its provider.</summary>
        /// <param name="activity">Original transport event.</param>
        /// <param name="scope">Unique client/turn prefix bound by the UI subscriber.</param>
        /// <returns>Detached activity or null for an absent identity.</returns>
        internal static CodexAgentActivity WithScope(CodexAgentActivity activity, string scope)
        {
            if (activity == null || string.IsNullOrEmpty(activity.Id)) return null;
            return new CodexAgentActivity
            {
                Id = scope + ":" + activity.Id, Kind = activity.Kind, Title = activity.Title,
                Detail = CodexAgentActivity.Limit(activity.Detail), Status = activity.Status,
                DurationMs = activity.DurationMs, Append = activity.Append
            };
        }

        /// <summary>Uses the actual VBAi result receipt rather than a transport's delivery-success label.</summary>
        /// <param name="output">Serialized tool result.</param>
        /// <returns>Completed, failed, or interrupted when native execution remains uncertain.</returns>
        internal static string ToolOutcome(string output)
        {
            var result = ReadObject(output);
            if (result == null) return "failed";
            if (result.TryGetValue("Data", out var raw) && raw is IDictionary<string, object> data &&
                data.TryGetValue("Uncertain", out var uncertain) && uncertain is bool unknown && unknown) return "interrupted";
            return result.TryGetValue("Ok", out var ok) && ok is bool success && success ? "completed" : "failed";
        }

        /// <summary>Reads only a public error/recovery reason, never the returned VBA or arbitrary result data.</summary>
        /// <param name="output">Serialized tool result.</param>
        /// <returns>Public diagnostic or null.</returns>
        internal static string ToolDiagnostic(string output)
        {
            var result = ReadObject(output);
            if (result == null) return null;
            if (result.TryGetValue("Error", out var error) && error is string text && !string.IsNullOrWhiteSpace(text)) return CodexAgentActivity.Limit(text);
            if (result.TryGetValue("Data", out var raw) && raw is IDictionary<string, object> data &&
                data.TryGetValue("Uncertain", out var uncertain) && uncertain is bool unknown && unknown &&
                data.TryGetValue("Reason", out var reason) && reason is string diagnostic) return CodexAgentActivity.Limit(diagnostic);
            return null;
        }

        /// <summary>Reads a bounded object for presentation without weakening the dispatch parser's own validation.</summary>
        /// <param name="json">Serialized metadata/result.</param>
        /// <returns>Dictionary or null when metadata cannot be projected.</returns>
        private static IDictionary<string, object> ReadObject(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try { return new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.DeserializeObject(json) as IDictionary<string, object>; }
            catch (ArgumentException) { return null; }
            catch (InvalidOperationException) { return null; }
        }
    }
}
