using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Validates the complete shape before approval or native dispatch.</summary>
        private static void ValidateVbeOptionArguments(IDictionary<string, object> values)
        {
            if (values.ContainsKey("OptionChanges"))
            {
                if (new[] { "Pane", "Property", "Value", "Query" }.Any(values.ContainsKey))
                    throw new ArgumentException("OptionChanges cannot be combined with single-option fields.");
                ValidateVbeOptionChanges(values["OptionChanges"]);
            }
            else
                foreach (string field in new[] { "Pane", "Property", "Value" })
                    if (!values.TryGetValue(field, out var value) || value == null ||
                        (field != "Value" && (!(value is string text) || string.IsNullOrWhiteSpace(text))))
                        throw new ArgumentException(field + " is required for a single-option edit.");
        }

        /// <summary>Rejects malformed or excessive nested edits without invoking a native mutation.</summary>
        private static void ValidateVbeOptionChanges(object value)
        {
            if (!(value is object[] changes) || changes.Length < 1 || changes.Length > 64)
                throw new ArgumentException("OptionChanges must contain one to sixty-four edits.");
            foreach (object change in changes)
            {
                if (!(change is IDictionary<string, object> fields) ||
                    fields.Keys.Any(key => !new[] { "Pane", "Property", "Value", "Query" }.Contains(key)))
                    throw new ArgumentException("Each option change must contain only Pane, Property, Value and optional Query.");
                foreach (string name in new[] { "Pane", "Property" })
                    if (!fields.TryGetValue(name, out var label) || !(label is string text) ||
                        string.IsNullOrWhiteSpace(text) || text.Length > 1024 || text.Any(char.IsControl))
                        throw new ArgumentException(name + " must be a nonempty single-line native label.");
                if (fields.TryGetValue("Query", out var query) && (!(query is string category) ||
                    string.IsNullOrWhiteSpace(category) || category.Length > 1024 || category.Any(char.IsControl)))
                    throw new ArgumentException("Query must be a nonempty single-line native category.");
                if (!fields.TryGetValue("Value", out var scalar) ||
                    (!(scalar is string) && !(scalar is bool) && !(scalar is int) &&
                     !(scalar is long) && !(scalar is double) && !(scalar is decimal)))
                    throw new ArgumentException("Value must be a string, number or boolean.");
            }
        }

        /// <summary>Describes a bounded collection of native edits sharing one complete revision and one commit.</summary>
        private static object VbeOptionChangesSchema() => new
        {
            type = "array", minItems = 1, maxItems = 64,
            items = new
            {
                type = "object", additionalProperties = false, required = new[] { "Pane", "Property", "Value" },
                properties = new Dictionary<string, object>
                {
                    ["Pane"] = new { type = "string" }, ["Property"] = new { type = "string" },
                    ["Query"] = new { type = "string" },
                    ["Value"] = new { anyOf = new object[] { new { type = "string" }, new { type = "number" }, new { type = "boolean" } } }
                }
            }
        };
    }
}
