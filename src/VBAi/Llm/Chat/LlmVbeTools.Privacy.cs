using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;

namespace VBAi
{

    /// <summary>Owns the llm vbe tools state and operations.</summary>
    internal sealed partial class LlmVbeTools
    {

        /// <summary>Maintains the read project grants state for llm vbe tools.</summary>
        private readonly HashSet<string> readProjectGrants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Maintains the shared context read allowed state for llm vbe tools.</summary>
        private bool sharedContextReadAllowed;

        // These commands do not return source, project inventories or native context.
        // New tools without a Project argument fail closed until classified here.
        /// <summary>Maintains the independent tools state for llm vbe tools.</summary>
        private static readonly HashSet<string> IndependentTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "status", "list_projects", "discover_tools", "invoke_tool", "read_user_file", "inspect_code_file",
            "certificate_trust", "verify_vba_signature_file", "list_signing_certificates",
            "list_form_control_types", "read_debug_options", "read_vbe_options", "set_vbe_option",
            "list_commands", "list_addins", "list_toolbars", "toolbar_controls", "open_debug_pane"
        };
        // These have a Project argument but native results/effects also span the VBE.
        /// <summary>Maintains the shared project tools state for llm vbe tools.</summary>
        private static readonly HashSet<string> SharedProjectTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "debug_global", "immediate_execute", "read_immediate", "inspect_local_scalars", "run_procedure", "procedure_run_status",
            "run_procedure_values", "procedure_values_status", "edit_watch", "remove_watch",
            "run_vba_tests", "stop_vba_tests",
            "read_project_signature_dialog", "close_standalone_project", "publish_solidworks_macro"
        };

        /// <summary>Sets the project grants and shared context policy for tool reads.</summary>
        /// <param name="projects">i enumerable&lt;string&gt; that supplies the projects for this operation.</param>
        /// <param name="sharedContext">Indicates whether shared context is enabled.</param>
        internal void SetReadAccess(IEnumerable<string> projects, bool sharedContext)
        {
            readProjectGrants.Clear();
            foreach (string project in projects ?? Enumerable.Empty<string>())
                if (!string.IsNullOrWhiteSpace(project)) readProjectGrants.Add(project);
            sharedContextReadAllowed = sharedContext;
        }

        /// <summary>Rejects a project read when the current session has not granted access.</summary>
        /// <param name="project">Text that supplies the project value. Use the format required by the calling operation.</param>
        internal void RequireProjectRead(string project)
        {
            if (string.IsNullOrEmpty(BoundProject)) return; // Unbound internal callers retain their contract.
            if (!SameProject(project, BoundProject) && !readProjectGrants.Contains(project ?? "") && !IsAuthorizedAlias(project))
                throw new InvalidOperationException(UiText.Get("Read access to another project is not authorized for this conversation."));
        }

        // Reference tokens use project names, while saved chat scopes use absolute paths.
        // Resolve that alias only against one live project; duplicate names fail closed.
        /// <summary>Checks whether a requested project alias resolves to an authorized project.</summary>
        /// <param name="selector">Text that supplies the selector value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is authorized alias on llm vbe tools.</returns>
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

        /// <summary>Compares project identities using the names accepted by the live VBE session.</summary>
        /// <param name="first">Text that supplies the first value. Use the format required by the calling operation.</param>
        /// <param name="second">Text that supplies the second value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for same project on llm vbe tools.</returns>
        private static bool SameProject(string first, string second) =>
            string.Equals(first, second, StringComparison.OrdinalIgnoreCase);

        /// <summary>Guards a tool request against the current project access policy.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="arguments">Text that supplies the arguments value. Use the format required by the calling operation.</param>
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

        /// <summary>Reads project fields only after the configured privacy checks pass.</summary>
        /// <param name="data">object that supplies the data for this operation.</param>
        /// <returns>i dictionary&lt;string, object&gt; produced by the operation for fields on llm vbe tools.</returns>
        private IDictionary<string, object> Fields(object data) =>
            json.DeserializeObject(json.Serialize(data)) as IDictionary<string, object>;

        /// <summary>Removes projects that are outside the session&apos;s explicit read grants.</summary>
        /// <param name="data">object that supplies the data for this operation.</param>
        /// <returns>object produced by the operation for filter projects on llm vbe tools.</returns>
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

        /// <summary>Removes ungranted project data from a tool response before it reaches the model.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="response">response that supplies the response for this operation.</param>
        /// <param name="requestedProject">Text that supplies the requested project value. Use the format required by the calling operation.</param>
        /// <returns>response produced by the operation for filter project response on llm vbe tools.</returns>
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

        /// <summary>Captures live project state within the session&apos;s authorized project scope.</summary>
        /// <returns>object produced by the operation for scoped live snapshot on llm vbe tools.</returns>
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
