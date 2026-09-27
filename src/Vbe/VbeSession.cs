using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
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
        private readonly VbeReferenceTypes referenceTypes;

        public VbeSession(object vbe) { this.vbe = vbe; debugger = new VbeDebug(vbe);
            forms = new VbeForms(vbe); components = new VbeProjectComponents(vbe, forms);
            editorWindows = new VbeEditorWindows(vbe); codeNavigation = new VbeCodeNavigation(vbe, forms);
            referenceTypes = new VbeReferenceTypes(vbe); }

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
                case "vbe_environment":
                    return Response.Success(editorWindows.Environment());
                case "list_addins":
                    return Response.Success(editorWindows.AddIns());
                case "focus_vbe_window":
                    return Response.Success(editorWindows.FocusWindow(request.WindowCaption, request.WindowType));
                case "show_vbe_window":
                    return Response.Success(editorWindows.ShowWindow(request.WindowCaption, request.WindowType));
                case "window_linkage":
                    return Response.Success(editorWindows.WindowLinkage(request.WindowCaption, request.WindowType));
                case "close_vbe_window":
                    return Response.Success(editorWindows.CloseWindow(request.WindowCaption, request.WindowType));
                case "code_panes":
                    return Response.Success(editorWindows.CodePanes());
                case "open_object_browser":
                    return Response.Success(debugger.OpenObjectBrowser(editorWindows));
                case "list_procedures":
                    return Response.Success(codeNavigation.Procedures(request.Project, request.Module));
                case "find_code":
                    return Response.Success(codeNavigation.Find(request));
                case "select_procedure":
                    return Response.Success(codeNavigation.SelectProcedure(request, debugger));
                case "create_event_procedure":
                    return Response.Success(codeNavigation.CreateEventProcedure(request));
                case "create_procedure":
                    return Response.Success(codeNavigation.CreateProcedure(request));
                case "replace_procedure":
                    return Response.Success(codeNavigation.ReplaceProcedure(request));
                case "remove_procedure":
                    return Response.Success(codeNavigation.RemoveProcedure(request));
                case "insert_code_file":
                    return Response.Success(codeNavigation.InsertCodeFile(request));
                case "inspect_code_file":
                    return Response.Success(codeNavigation.InspectCodeFile(request.Path));
                case "project_properties":
                    return Response.Success(components.ProjectProperties(request.Project));
                case "project_persistence_status":
                    return Response.Success(components.PersistenceStatus(request.Project));
                case "save_host_document":
                    return Response.Success(components.SaveHostDocument(request));
                case "save_host_document_as":
                    return Response.Success(components.SaveHostDocumentAs(request));
                case "project_signature_status":
                    return Response.Success(components.SignatureStatus(request.Project));
                case "list_signing_certificates":
                    return Response.Success(ListSigningCertificates());
                case "read_project_signature_dialog":
                    return Response.Success(debugger.QueueSignatureDialog(request));
                case "sign_project":
                    return Response.Success(BeginSignProject(request));
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
                case "set_class_instancing":
                    return Response.Success(components.SetClassInstancing(request));
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
                case "list_reference_types":
                    return Response.Success(referenceTypes.ListTypes(request));
                case "list_type_members":
                    return Response.Success(referenceTypes.ListMembers(request));
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
                case "run_sub":
                    return Response.Success(debugger.RunSub(request));
                case "compile_project":
                    return Response.Success(debugger.CompileProject(request));
                case "open_debug_pane":
                    return Response.Success(debugger.OpenDebugPane(request.Action, editorWindows));
                case "add_watch":
                    return Response.Success(debugger.QueueAddWatchDialog(request));
                case "edit_watch":
                    return Response.Success(debugger.QueueEditWatchDialog(request));
                case "quick_watch":
                    return Response.Success(debugger.QueueQuickWatchDialog(request));
                case "read_debug_options":
                    return Response.Success(debugger.QueueDebugOptionsDialog());
                case "read_vbe_options":
                    return Response.Success(debugger.QueueDebugOptionsDialog());
                case "remove_watch":
                    return Response.Success(debugger.RemoveSelectedWatch(request));
                case "debug_global":
                    return Response.Success(debugger.ExecuteGlobalDebugCommand(request));
                case "list_commands":
                    return Response.Success(debugger.ListCommands(request.Query, request.Offset, request.Limit));
                case "select_code":
                    return Response.Success(debugger.SelectCode(request));
                case "select_code_range":
                    return Response.Success(debugger.SelectCodeRange(request));
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
                case "form_list_items":
                    return Response.Success(forms.ListItems(request));
                case "set_form_list_initializer":
                    return Response.Success(forms.SetListInitializer(request));
                case "probe_append_form_list_item":
                    return Response.Success(forms.AppendListItem(request));
                case "add_form_list_item":
                    return Response.Success(forms.AddListItem(request));
                case "remove_form_list_item":
                    return Response.Success(forms.RemoveListItem(request));
                case "form_event_catalog":
                    return Response.Success(forms.EventCatalog(request.Project, request.Form, request.ControlPath));
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
                case "z_order_form_control":
                    return Response.Success(forms.ZOrderControl(request));
                case "form_property_accessors":
                    return Response.Success(forms.PropertyAccessors(request));
                case "duplicate_form_label":
                    return Response.Success(forms.DuplicateLabel(request));
                case "duplicate_form_textbox":
                    return Response.Success(forms.DuplicateTextBox(request));
                case "duplicate_form_checkbox":
                    return Response.Success(forms.DuplicateCheckBox(request));
                case "duplicate_form_togglebutton":
                    return Response.Success(forms.DuplicateToggleButton(request));
                case "duplicate_form_commandbutton":
                    return Response.Success(forms.DuplicateCommandButton(request));
                case "duplicate_form_combobox":
                    return Response.Success(forms.DuplicateComboBox(request));
                case "duplicate_empty_form_frame":
                    return Response.Success(forms.DuplicateEmptyFrame(request));
                case "frame_copy_plan":
                    return Response.Success(forms.FrameCopyPlan(request));
                case "duplicate_form_frame_labels":
                    return Response.Success(forms.DuplicateFrameWithLabels(request));
                case "frame_simple_copy_plan":
                    return Response.Success(forms.FrameSimpleCopyPlan(request));
                case "duplicate_form_frame_simple_children":
                    return Response.Success(forms.DuplicateFrameWithSimpleChildren(request));
                case "frame_profile_copy_plan":
                    return Response.Success(forms.FrameProfileCopyPlan(request));
                case "duplicate_form_frame_profiled":
                    return Response.Success(forms.DuplicateFrameProfiled(request));
                case "duplicate_form_optionbutton":
                    return Response.Success(forms.DuplicateOptionButton(request));
                case "remove_form_control":
                    return Response.Success(forms.RemoveControl(request));
                case "add_form_page":
                    return Response.Success(forms.AddPageOrTab(request, "Pages"));
                case "add_form_tab":
                    return Response.Success(forms.AddPageOrTab(request, "Tabs"));
                case "remove_form_page_tab":
                    return Response.Success(forms.RemovePageOrTab(request));
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

        private object BeginSignProject(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.ExpectedProjectVersion) ||
                string.IsNullOrWhiteSpace(request.CertificateThumbprint))
                throw new ArgumentException("Project, ExpectedProjectVersion and CertificateThumbprint are required.");
            dynamic state = components.ProjectProperties(request.Project);
            if ((int)state.Mode != 2 || request.ExpectedMode != 2)
                throw new InvalidOperationException("The project must be in design mode (ExpectedMode=2).");
            if (!string.Equals((string)state.Version, request.ExpectedProjectVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since it was read.");
            dynamic liveProject = GetProject(request.Project);
            if (!(bool)liveProject.Saved)
                throw new InvalidOperationException("Save the VBA project before signing it.");
            dynamic signatureStatus = components.SignatureStatus(request.Project);
            if ((bool)signatureStatus.Available && (bool)signatureStatus.Signed)
                throw new InvalidOperationException("The host reports an existing VBA signature; this command only adds the first signature.");
            bool unsignedVerified = (bool)signatureStatus.Available && !(bool)signatureStatus.Signed;
            if (string.Equals(System.Diagnostics.Process.GetCurrentProcess().ProcessName, "EXCEL",
                StringComparison.OrdinalIgnoreCase))
            {
                string projectPath = null;
                try { projectPath = (string)liveProject.FileName; }
                catch { }
                if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath) ||
                    !File.Exists(projectPath) || !unsignedVerified)
                    throw new InvalidOperationException("Save the macro-enabled Excel workbook before adding its first VBA signature.");
                string extension = Path.GetExtension(projectPath);
                if (!new[] { ".xlsm", ".xlam", ".xlsb", ".xltm", ".xls", ".xla", ".xlt" }
                    .Contains(extension, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The saved Excel format does not support a VBA project signature.");
                bool hasVbaContent = false;
                foreach (dynamic component in liveProject.VBComponents)
                    if ((int)component.Type != 100 || (int)component.CodeModule.CountOfLines > 0)
                    { hasVbaContent = true; break; }
                if (!hasVbaContent)
                    throw new InvalidOperationException("The Excel workbook has no VBA content to sign; add code and save it first.");
            }
            string thumbprint = request.CertificateThumbprint.Replace(" ", "").ToUpperInvariant();
            if (!Regex.IsMatch(thumbprint, "^[0-9A-F]{40}$"))
                throw new ArgumentException("CertificateThumbprint must be a SHA-1 certificate thumbprint.");
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var matches = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (matches.Count != 1) throw new InvalidOperationException("The selected certificate is absent or ambiguous.");
                var certificate = matches[0];
                if (!certificate.HasPrivateKey || DateTime.Now < certificate.NotBefore || DateTime.Now > certificate.NotAfter)
                    throw new InvalidOperationException("The certificate needs a usable private key and current validity.");
                bool codeSigning = certificate.Extensions.OfType<X509EnhancedKeyUsageExtension>()
                    .SelectMany(extension => extension.EnhancedKeyUsages.Cast<System.Security.Cryptography.Oid>())
                    .Any(oid => oid.Value == "1.3.6.1.5.5.7.3.3");
                if (!codeSigning) throw new InvalidOperationException("The certificate is not intended for code signing.");
                string displayName = certificate.GetNameInfo(X509NameType.SimpleName, false);
                if (string.IsNullOrWhiteSpace(displayName))
                    throw new InvalidOperationException("The certificate display name is empty.");
                var matchingThumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in store.Certificates.Cast<X509Certificate2>())
                    if (string.Equals(item.GetNameInfo(X509NameType.SimpleName, false), displayName,
                        StringComparison.OrdinalIgnoreCase)) matchingThumbprints.Add(item.Thumbprint);
                using (var machineStore = new X509Store(StoreName.My, StoreLocation.LocalMachine))
                {
                    machineStore.Open(OpenFlags.ReadOnly);
                    foreach (var item in machineStore.Certificates.Cast<X509Certificate2>())
                        if (string.Equals(item.GetNameInfo(X509NameType.SimpleName, false), displayName,
                            StringComparison.OrdinalIgnoreCase)) matchingThumbprints.Add(item.Thumbprint);
                }
                if (matchingThumbprints.Count != 1 || !matchingThumbprints.Contains(thumbprint))
                    throw new InvalidOperationException("The certificate display name is ambiguous across personal certificate stores.");
                object scheduled = debugger.QueueSignatureDialog(request);
                return new { Scheduled = true, Project = request.Project,
                    CertificateThumbprint = thumbprint, CertificateName = displayName,
                    UnsignedVerified = unsignedVerified,
                    NativeCommand = scheduled };
            }
        }

        private object ListSigningCertificates()
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                return store.Certificates.Cast<X509Certificate2>()
                    .Where(certificate => certificate.HasPrivateKey &&
                        certificate.Extensions.OfType<X509EnhancedKeyUsageExtension>()
                            .SelectMany(extension => extension.EnhancedKeyUsages.Cast<System.Security.Cryptography.Oid>())
                            .Any(oid => oid.Value == "1.3.6.1.5.5.7.3.3"))
                    .Select(certificate => new {
                        certificate.Thumbprint,
                        Name = certificate.GetNameInfo(X509NameType.SimpleName, false),
                        certificate.Subject, certificate.Issuer,
                        NotBefore = certificate.NotBefore.ToString("o"),
                        NotAfter = certificate.NotAfter.ToString("o"),
                        EligibleNow = DateTime.Now >= certificate.NotBefore && DateTime.Now <= certificate.NotAfter
                    }).ToArray();
            }
        }

        internal VbaGitProject GitProject(string projectName, string hostPath)
        {
            return new VbaGitProject(() => (object)GetProject(projectName), hostPath);
        }

        internal string GitScope(string projectName)
        {
            dynamic project = GetProject(projectName);
            string path = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) throw new InvalidOperationException("Enregistrez le document avant d’utiliser Git.");
            return Path.GetFullPath(path);
        }

        internal object PersistProjectSignature(string projectName)
        {
            return components.PersistExcelSignature(projectName);
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
            return VbeProjectResolver.Resolve(vbe, name);
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
