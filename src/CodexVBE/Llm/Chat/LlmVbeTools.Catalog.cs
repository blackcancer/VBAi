using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed partial class LlmVbeTools
    {
        private readonly HashSet<string> loadedFamilies = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<string, int> familyPriority = new Dictionary<string, int>(StringComparer.Ordinal);
        private int nextFamilyPriority;
        private static readonly string[] Families = { "code", "forms", "debug", "git", "environment" };
        private static readonly HashSet<string> CoreTools = new HashSet<string>(StringComparer.Ordinal)
        { "status", "list_projects", "list_modules", "read_module", "monaco_open", "monaco_read", "discover_tools", "invoke_tool" };
        private static object[] CatalogDefinitions => new[] {
            Definition("discover_tools", "Discover complete tool schemas for one family: code, forms, debug, git or environment; Family=all lists every available schema. Discussion/Plan expose inspections only. HTTP also loads discovered families for following model rounds. All writes remain subject to Agent mode and project/privacy policies.", new[] { "Family" }, "Family"),
            Definition("invoke_tool", "Invoke an exact tool discovered with discover_tools. ToolName is its name and ArgumentsJson is its JSON argument object encoded as a string. All normal mode, project, privacy and approval guards apply. Never recursively invoke a catalogue gateway.", new[] { "ToolName", "ArgumentsJson" }, "ToolName", "ArgumentsJson")
        };
        internal static bool IsCatalogTool(string name) => name == "discover_tools" || name == "invoke_tool";
        internal static string ToolFamily(string name)
        {
            if (name.StartsWith("git_", StringComparison.Ordinal)) return "git";
            if (name.IndexOf("form", StringComparison.Ordinal) >= 0 || name.IndexOf("designer", StringComparison.Ordinal) >= 0) return "forms";
            if (name.IndexOf("debug", StringComparison.Ordinal) >= 0 || name.IndexOf("watch", StringComparison.Ordinal) >= 0 ||
                name.IndexOf("breakpoint", StringComparison.Ordinal) >= 0 || name.StartsWith("run_", StringComparison.Ordinal) ||
                name.StartsWith("step_", StringComparison.Ordinal) || name == "compile_project" || name == "immediate_execute" || name == "invoke_command" || name == "list_commands") return "debug";
            if (name.StartsWith("monaco_", StringComparison.Ordinal) || name.IndexOf("code", StringComparison.Ordinal) >= 0 ||
                name.IndexOf("module", StringComparison.Ordinal) >= 0 || name.IndexOf("procedure", StringComparison.Ordinal) >= 0 ||
                name.IndexOf("rename", StringComparison.Ordinal) >= 0 || name == "replace_lines" || name == "project_symbols") return "code";
            return "environment";
        }
        internal object[] CatalogForProvider(bool gatewayOnly = false)
        {
            return Definitions.Where(raw =>
            {
                string name = (string)((dynamic)raw).function.name;
                return (Mode == ChatMode.Agent || ReadOnlyTools.Contains(name)) &&
                    (CoreTools.Contains(name) || (!gatewayOnly && loadedFamilies.Contains(ToolFamily(name))));
            }).OrderByDescending(raw => CoreTools.Contains((string)((dynamic)raw).function.name))
                .ThenByDescending(raw => familyPriority.TryGetValue(ToolFamily((string)((dynamic)raw).function.name), out var priority) ? priority : 0)
                .Take(64).ToArray();
        }
        internal void ResetCatalog() { loadedFamilies.Clear(); familyPriority.Clear(); nextFamilyPriority = 0; }
        internal async Task<string> InvokeCatalogAsync(string name, string arguments)
        {
            try
            {
                var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                if (name == "discover_tools")
                {
                    if (values.Count != 1 || !values.TryGetValue("Family", out var value) || !(value is string)) throw new ArgumentException("Family is required.");
                    string family = (string)value;
                    if (family != "all" && !Families.Contains(family)) throw new ArgumentException("Unknown tool family. Use code, forms, debug, git, environment or all.");
                    var selected = family == "all" ? Families : new[] { family };
                    foreach (string item in selected) { loadedFamilies.Add(item); familyPriority[item] = ++nextFamilyPriority; }
                    var definitions = Definitions.Where(raw => {
                        string tool = (string)((dynamic)raw).function.name;
                        return !IsCatalogTool(tool) && selected.Contains(ToolFamily(tool)) && (Mode == ChatMode.Agent || ReadOnlyTools.Contains(tool));
                    }).ToArray();
                    return json.Serialize(Response.Success(new { Mode = Mode.ToString(), Families = selected, Tools = definitions,
                        EditingAvailable = Mode == ChatMode.Agent, Instruction = "Use invoke_tool or the loaded exact tool. Runtime project and privacy guards still apply." }));
                }
                if (values.Count != 2 || !values.TryGetValue("ToolName", out var target) || !(target is string) ||
                    !values.TryGetValue("ArgumentsJson", out var args) || !(args is string)) throw new ArgumentException("ToolName and ArgumentsJson are required strings.");
                if (IsCatalogTool((string)target)) throw new ArgumentException("Recursive catalogue invocation is not allowed.");
                return await InvokeAsync((string)target, (string)args);
            }
            catch (Exception error) { return json.Serialize(Response.Failure(error.Message)); }
        }
    }
}
