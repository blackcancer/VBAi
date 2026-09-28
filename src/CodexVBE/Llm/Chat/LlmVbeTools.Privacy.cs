using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;

namespace CodexVBE
{
    internal sealed partial class LlmVbeTools
    {
        private readonly HashSet<string> readProjectGrants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool sharedContextReadAllowed;

        // These commands do not return source, project inventories or native context.
        // New tools without a Project argument fail closed until classified here.
        private static readonly HashSet<string> IndependentTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "status", "list_projects", "discover_tools", "invoke_tool", "read_user_file", "inspect_code_file",
            "certificate_trust", "verify_vba_signature_file", "list_signing_certificates",
            "list_form_control_types", "read_debug_options", "read_vbe_options", "set_vbe_option",
            "list_commands", "list_addins", "list_toolbars", "toolbar_controls", "open_debug_pane"
        };
        // These have a Project argument but native results/effects also span the VBE.
        private static readonly HashSet<string> SharedProjectTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "debug_global", "immediate_execute", "run_procedure", "procedure_run_status",
            "run_procedure_values", "procedure_values_status", "edit_watch", "remove_watch",
            "read_project_signature_dialog", "close_standalone_project"
        };

        internal void SetReadAccess(IEnumerable<string> projects, bool sharedContext)
        {
            readProjectGrants.Clear();
            foreach (string project in projects ?? Enumerable.Empty<string>())
                if (!string.IsNullOrWhiteSpace(project)) readProjectGrants.Add(project);
            sharedContextReadAllowed = sharedContext;
        }

        internal void RequireProjectRead(string project)
        {
            if (string.IsNullOrEmpty(BoundProject)) return; // Unbound internal callers retain their contract.
            if (!SameProject(project, BoundProject) && !readProjectGrants.Contains(project ?? "") && !IsAuthorizedAlias(project))
                throw new InvalidOperationException(UiText.Get("Read access to another project is not authorized for this conversation."));
        }

        // Reference tokens use project names, while saved chat scopes use absolute paths.
        // Resolve that alias only against one live project; duplicate names fail closed.
        private bool IsAuthorizedAlias(string selector)
        {
            if (string.IsNullOrWhiteSpace(selector)) return false;
            try
            {
                var inventory = Execute(new Request { Command = "list_projects" });
                if (!inventory.Ok) return false;
                var rows = (json.DeserializeObject(json.Serialize(inventory.Data)) as object[] ?? new object[0])
                    .OfType<IDictionary<string, object>>().ToArray();
                var targets = rows.Where(row => SameProject(Convert.ToString(row["Name"]), selector) ||
                    SameProject(Convert.ToString(row["FileName"]), selector)).ToArray();
                if (targets.Length != 1) return false;
                var target = targets[0];
                return new[] { BoundProject }.Concat(readProjectGrants).Any(grant =>
                    SameProject(Convert.ToString(target["FileName"]), grant) ||
                    (SameProject(Convert.ToString(target["Name"]), grant) && rows.Count(row => SameProject(Convert.ToString(row["Name"]), grant)) == 1));
            }
            catch { return false; }
        }

        private static bool SameProject(string first, string second) =>
            string.Equals(first, second, StringComparison.OrdinalIgnoreCase);

        private void GuardProjectPrivacy(string name, string arguments)
        {
            if (string.IsNullOrEmpty(BoundProject)) return;
            var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
            object project;
            if (values != null && values.TryGetValue("Project", out project))
            {
                // Shape validation remains responsible for malformed selectors. No native
                // branch accepts a missing, non-string or blank required Project argument.
                if (!(project is string) || string.IsNullOrWhiteSpace((string)project)) return;
                string selector = Convert.ToString(project);
                if (ReadOnlyTools.Contains(name) && name != "compile_project") RequireProjectRead(selector);
                else if (!SameProject(selector, BoundProject))
                    throw new InvalidOperationException("Cette action vise un autre projet que celui de la conversation.");
                if (!SharedProjectTools.Contains(name)) return;
            }
            else if (IndependentTools.Contains(name)) return;
            // Git tools resolve their repository from the bound macro, never from active VBE state.
            else if (name.StartsWith("git_", StringComparison.Ordinal)) return;
            if (!sharedContextReadAllowed)
                throw new InvalidOperationException(UiText.Get("This tool uses shared VBE context. Authorize shared context in Project access before using it."));
        }

        private IDictionary<string, object> Fields(object data) =>
            json.DeserializeObject(json.Serialize(data)) as IDictionary<string, object>;

        private object FilterProjects(object data)
        {
            if (string.IsNullOrEmpty(BoundProject)) return data;
            var projects = json.DeserializeObject(json.Serialize(data)) as object[] ?? new object[0];
            // A name selector is allowed only when unique. Never treat several same-name
            // projects as the bound project, including projects granted by name.
            var rows = projects.OfType<IDictionary<string, object>>().ToArray();
            return rows.Where(row => new[] { BoundProject }.Concat(readProjectGrants).Any(selector =>
            {
                object path, projectName;
                if (row.TryGetValue("FileName", out path) && !string.IsNullOrEmpty(Convert.ToString(path)) && SameProject(Convert.ToString(path), selector)) return true;
                return row.TryGetValue("Name", out projectName) && SameProject(Convert.ToString(projectName), selector)
                    && rows.Count(other => SameProject(Convert.ToString(other["Name"]), selector)) == 1;
            })).ToArray();
        }

        private Response FilterProjectResponse(string name, Response response, string requestedProject)
        {
            if (!response.Ok || string.IsNullOrEmpty(BoundProject)) return response;
            if (name == "list_projects") return Response.Success(FilterProjects(response.Data));
            if (name != "debug_state") return response;
            var fields = Fields(response.Data);
            object selectedName, selectedPath;
            fields.TryGetValue("SelectedProject", out selectedName);
            fields.TryGetValue("SelectedProjectPath", out selectedPath);
            string selector = string.IsNullOrEmpty(Convert.ToString(selectedPath)) ? Convert.ToString(selectedName) : Convert.ToString(selectedPath);
            // Context belongs to the requested project, not merely any readable project.
            object requestedName;
            fields.TryGetValue("Project", out requestedName);
            bool same = !string.IsNullOrEmpty(selector) && (SameProject(selector, requestedProject) ||
                (SameProject(Convert.ToString(selectedName), requestedProject) && SameProject(Convert.ToString(selectedName), Convert.ToString(requestedName))));
            if (!same)
                foreach (string field in new[] { "SelectedProject", "SelectedProjectPath", "ActiveModule", "Selection" }) fields[field] = null;
            return Response.Success(fields);
        }

        private object ScopedLiveSnapshot()
        {
            Response projects;
            try { projects = Execute(new Request { Command = "list_projects" }); }
            catch { projects = Response.Failure("Project inventory unavailable."); }
            using (var process = Process.GetCurrentProcess())
                return new { HostProcess = process.ProcessName, HostProcessId = process.Id,
                    VbeConnected = true, HostAnsiCodePage = Encoding.Default.CodePage,
                    CodeFileEncodingPolicy = LlmVbeContext.EncodingInstructions,
                    Projects = projects.Ok ? FilterProjects(projects.Data) : null,
                    ProjectsError = projects.Ok ? null : projects.Error,
                    AvailableAccess = "Only the bound project and explicitly authorized read projects; shared VBE context requires separate user authorization." };
        }
    }
}
