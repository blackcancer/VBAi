using System;
using System.Collections.Generic;
using System.Linq;
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

        public VbeSession(object vbe) { this.vbe = vbe; debugger = new VbeDebug(vbe); forms = new VbeForms(vbe); }

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
                case "list_references":
                    return Response.Success(ListReferences(request.Project));
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
                case "form_state":
                    return Response.Success(forms.State(request.Project, request.Form));
                case "form_tree":
                    return Response.Success(forms.Tree(request.Project, request.Form));
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

        private object ListReferences(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = new List<object>();
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
                result.Add(new { Name = name, Guid = (string)reference.GUID,
                    Major = (int)reference.Major, Minor = (int)reference.Minor,
                    IsBroken = broken, FullPath = fullPath });
            }
            return new { Project = projectName, References = result };
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
