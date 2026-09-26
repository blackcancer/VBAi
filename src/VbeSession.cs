using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed class VbeSession
    {
        private readonly dynamic vbe;
        private readonly VbeDebug debugger;
        private readonly VbeForms forms;
        private readonly VbeProjectComponents components;
        private readonly VbeEditorWindows editorWindows;
        private readonly VbeCodeNavigation codeNavigation;

        public VbeSession(object vbe) { this.vbe = vbe; debugger = new VbeDebug(vbe);
            forms = new VbeForms(vbe); components = new VbeProjectComponents(vbe, forms);
            editorWindows = new VbeEditorWindows(vbe); codeNavigation = new VbeCodeNavigation(vbe); }

        public Response Execute(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Command))
                return Response.Failure("A command is required.");

            switch (request.Command)
            {
                case "status":
                    return Response.Success(new { Version = "0.1.0", Connected = true });
                case "list_projects":
                    return Response.Success(ListProjects());
                case "list_modules":
                    return Response.Success(ListModules(request.Project));
                case "vbe_windows":
                    return Response.Success(editorWindows.Windows());
                case "code_panes":
                    return Response.Success(editorWindows.CodePanes());
                case "list_procedures":
                    return Response.Success(codeNavigation.Procedures(request.Project, request.Module));
                case "find_code":
                    return Response.Success(codeNavigation.Find(request));
                case "select_procedure":
                    return Response.Success(codeNavigation.SelectProcedure(request, debugger));
                case "project_properties":
                    return Response.Success(components.ProjectProperties(request.Project));
                case "component_properties":
                    return Response.Success(components.ComponentProperties(request.Project, request.Module));
                case "component_property_value":
                    return Response.Success(components.ComponentPropertyValue(request.Project, request.Module, request.Property));
                case "component_probe":
                    return Response.Success(components.ComponentProbe(request.Project, request.Module, request.Action, request.Query));
                case "set_project_property":
                    return Response.Success(components.SetProjectProperty(request));
                case "set_component_property":
                    return Response.Success(components.SetComponentProperty(request));
                case "rename_project":
                    return Response.Failure("Project rename is disabled: it correlated with an Excel process crash during validation.");
                case "rename_component":
                    return Response.Success(components.RenameComponent(request));
                case "remove_component":
                    return Response.Success(components.RemoveComponent(request));
                case "import_component":
                    return Response.Success(components.ImportComponent(request));
                case "export_component":
                    return Response.Success(components.ExportComponent(request));
                case "list_references":
                    return Response.Success(ListReferences(request.Project));
                case "add_reference_guid":
                    return Response.Success(AddReferenceGuid(request));
                case "add_reference_file":
                    return Response.Success(AddReferenceFile(request));
                case "remove_reference":
                    return Response.Success(RemoveReference(request));
                case "read_module":
                    return Response.Success(ReadModule(request.Project, request.Module));
                case "create_module":
                    return Response.Success(CreateComponent(request, 1));
                case "create_class":
                    return Response.Success(CreateComponent(request, 2));
                case "replace_lines":
                    return ReplaceLines(request);
                case "debug_state":
                    return Response.Success(debugger.State(request.Project));
                case "list_commands":
                    return Response.Success(debugger.ListCommands(request.Query));
                case "select_code":
                    return Response.Success(debugger.SelectCode(request));
                case "invoke_debug":
                    return Response.Success(debugger.InvokeCommand(request));
                case "list_forms":
                    return Response.Success(forms.List(request.Project));
                case "list_form_control_types":
                    return Response.Success(forms.ControlTypes());
                case "form_state":
                    return Response.Success(forms.State(request.Project, request.Form));
                case "form_tree":
                    return Response.Success(forms.Tree(request.Project, request.Form));
                case "form_parent_probe":
                    return Response.Success(forms.ParentProbe(request.Project, request.Form));
                case "form_properties":
                    return Response.Success(forms.Properties(request.Project, request.Form));
                case "set_form_property":
                    return Response.Success(forms.SetProperty(request));
                case "set_form_picture":
                    return Response.Success(forms.SetPicture(request));
                case "form_control_properties":
                    return Response.Success(forms.ControlProperties(request.Project, request.Form, request.Control));
                case "create_form":
                    return Response.Success(forms.Create(request));
                case "open_form":
                    return Response.Success(forms.Open(request.Project, request.Form));
                case "add_form_control":
                    return Response.Success(forms.AddControl(request));
                case "add_nested_form_control":
                    return Response.Success(forms.AddNestedControl(request));
                case "set_form_node_property":
                    return Response.Success(forms.SetNodeProperty(request));
                case "set_form_node_picture":
                    return Response.Success(forms.SetNodePicture(request));
                case "set_form_control_geometry":
                    return Response.Success(forms.SetControlGeometry(request));
                case "rename_form_control":
                    return Response.Success(forms.RenameControl(request));
                case "set_form_control_caption":
                    return Response.Success(forms.SetControlCaption(request));
                case "set_form_control_font":
                    return Response.Success(forms.SetControlFont(request));
                default:
                    return Response.Failure("Unknown command: " + request.Command);
            }
        }

        private object ListProjects()
        {
            var result = new List<object>();
            foreach (dynamic project in vbe.VBProjects)
            {
                string fileName = null;
                try { fileName = (string)project.FileName; }
                catch (Exception) { } // An unsaved host document has no accessible path.
                result.Add(new { Name = (string)project.Name, FileName = fileName, Mode = (int)project.Mode });
            }
            return result;
        }

        private object ListModules(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                result.Add(new { Name = (string)component.Name, Type = (int)component.Type,
                    Lines = (int)component.CodeModule.CountOfLines });
            return result;
        }

        private object ReadModule(string projectName, string moduleName)
        {
            dynamic module = GetModule(projectName, moduleName);
            string code = GetCode(module);
            return new { Project = projectName, Module = moduleName, Code = code, Sha256 = Hash(code) };
        }

        private sealed class ReferenceInfo
        {
            public string Name { get; set; }
            public string Guid { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
            public bool BuiltIn { get; set; }
            public string FullPath { get; set; }
        }

        private object ListReferences(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = ReadReferences(project);
            return new { Project = projectName, Version = ReferencesVersion(result), References = result };
        }

        private static List<ReferenceInfo> ReadReferences(dynamic project)
        {
            var result = new List<ReferenceInfo>();
            foreach (dynamic reference in project.References)
            {
                bool broken = (bool)reference.IsBroken;
                string name = null;
                string fullPath = null;
                if (!broken)
                {
                    try { name = (string)reference.Name; } catch { }
                    try { fullPath = (string)reference.FullPath; } catch { }
                }
                result.Add(new ReferenceInfo { Name = name, Guid = (string)reference.GUID,
                    Major = (int)reference.Major, Minor = (int)reference.Minor,
                    IsBroken = broken, BuiltIn = (bool)reference.BuiltIn, FullPath = fullPath });
            }
            return result;
        }

        private static string ReferencesVersion(List<ReferenceInfo> references)
        {
            var text = new StringBuilder();
            foreach (var reference in references)
            {
                text.Append(reference.Guid).Append('|').Append(reference.Major).Append('|')
                    .Append(reference.Minor).Append('|').Append(reference.IsBroken).Append('|')
                    .Append(reference.BuiltIn).Append('|')
                    .Append(reference.Name).Append('|').Append(reference.FullPath).Append('\n');
            }
            return Hash(text.ToString());
        }

        private dynamic CheckedReferenceProject(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedReferencesVersion))
                throw new ArgumentException("ExpectedReferencesVersion is required from list_references.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            if (!string.Equals(ReferencesVersion(ReadReferences(project)), request.ExpectedReferencesVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Project references changed since they were read.");
            return project;
        }

        private object AddReferenceGuid(Request request)
        {
            System.Guid parsed;
            if (!System.Guid.TryParse(request.Guid, out parsed) || request.Major < 0 || request.Minor < 0)
                throw new ArgumentException("Guid, nonnegative Major and Minor are required.");
            dynamic project = CheckedReferenceProject(request);
            if (((List<ReferenceInfo>)ReadReferences(project)).Any(item => string.Equals(item.Guid, parsed.ToString("B"),
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("A reference with this GUID is already selected.");
            project.References.AddFromGuid(parsed.ToString("B"), request.Major, request.Minor);
            return ListReferences(request.Project);
        }

        private object AddReferenceFile(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Path) ||
                !Regex.IsMatch(request.Path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("A fully qualified reference file path is required.");
            string path = Path.GetFullPath(request.Path);
            if (!File.Exists(path)) throw new FileNotFoundException("Reference file not found.", path);
            dynamic project = CheckedReferenceProject(request);
            project.References.AddFromFile(path);
            return ListReferences(request.Project);
        }

        private object RemoveReference(Request request)
        {
            System.Guid parsed;
            if (!System.Guid.TryParse(request.Guid, out parsed) || request.Major < 0 || request.Minor < 0)
                throw new ArgumentException("Guid, nonnegative Major and Minor are required.");
            dynamic project = CheckedReferenceProject(request);
            dynamic target = null;
            foreach (dynamic reference in project.References)
                if (string.Equals((string)reference.GUID, parsed.ToString("B"), StringComparison.OrdinalIgnoreCase) &&
                    (int)reference.Major == request.Major && (int)reference.Minor == request.Minor)
                { target = reference; break; }
            if (target == null) throw new InvalidOperationException("The exact reference was not found.");
            if ((bool)target.BuiltIn)
                throw new InvalidOperationException("The VBE marks this reference as built in and non-removable.");
            project.References.Remove(target);
            return ListReferences(request.Project);
        }

        private object CreateComponent(Request request, int componentType)
        {
            if (string.IsNullOrWhiteSpace(request.Module) ||
                !Regex.IsMatch(request.Module, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Module must start with a letter and contain at most 40 letters, digits or underscores.");
            if (request.ExpectedMode != 2)
                throw new ArgumentException("ExpectedMode must be 2 (design mode), obtained from list_projects.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != request.ExpectedMode)
                throw new InvalidOperationException("The project is no longer in design mode.");
            foreach (dynamic existing in project.VBComponents)
                if (string.Equals((string)existing.Name, request.Module, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");

            dynamic component = project.VBComponents.Add(componentType);
            try { component.Name = request.Module; }
            catch
            {
                // Only the newly created component is rolled back when its requested name is rejected.
                try { project.VBComponents.Remove(component); } catch { }
                throw;
            }
            string actualName = (string)component.Name;
            int actualType = (int)component.Type;
            if (!string.Equals(actualName, request.Module, StringComparison.Ordinal) || actualType != componentType)
                throw new InvalidOperationException("The VBE did not create the requested component identity.");
            dynamic module = component.CodeModule;
            string code = GetCode(module);
            return new { Project = request.Project, Module = actualName, Type = actualType,
                Lines = (int)module.CountOfLines, Code = code, Sha256 = Hash(code) };
        }

        private Response ReplaceLines(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256))
                return Response.Failure("ExpectedSha256 is required for edits.");
            if (request.StartLine < 1 || request.Count < 0 || request.Text == null)
                return Response.Failure("Invalid line range or replacement text.");

            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2) // vbext_vm_Design
                return Response.Failure("The project must be in design mode before editing.");
            dynamic module = GetModule(request.Project, request.Module);
            string before = GetCode(module);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                return Response.Failure("The module changed since it was read.");
            int lineCount = (int)module.CountOfLines;
            if (request.StartLine > lineCount + 1 || request.Count > lineCount - request.StartLine + 1)
                return Response.Failure("The requested line range is outside the module.");

            if (request.Count > 0) module.DeleteLines(request.StartLine, request.Count);
            if (request.Text.Length > 0) module.InsertLines(request.StartLine, request.Text);
            string after = GetCode(module);
            return Response.Success(new { Sha256 = Hash(after), Lines = (int)module.CountOfLines });
        }

        private dynamic GetProject(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Project is required.");
            var matches = new List<dynamic>();
            foreach (dynamic project in vbe.VBProjects)
                if (string.Equals((string)project.Name, name, StringComparison.OrdinalIgnoreCase)) matches.Add(project);
            if (matches.Count != 1) throw new InvalidOperationException("Project name is absent or ambiguous: " + name);
            return matches[0];
        }

        private dynamic GetModule(string projectName, string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module is required.");
            dynamic project = GetProject(projectName);
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                    return component.CodeModule;
            throw new InvalidOperationException("Module not found: " + moduleName);
        }

        private static string GetCode(dynamic module)
        {
            int count = (int)module.CountOfLines;
            return count == 0 ? string.Empty : (string)module.Lines[1, count];
        }

        private static string Hash(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code))).Replace("-", "").ToLowerInvariant();
        }
    }
}
