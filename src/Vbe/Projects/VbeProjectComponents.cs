using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class VbeProjectComponents
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        private readonly dynamic vbe;
        private readonly VbeForms forms;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };

        public VbeProjectComponents(object vbe, VbeForms forms)
        {
            this.vbe = vbe;
            this.forms = forms;
        }

        public object ProjectProperties(string projectName)
        {
            dynamic project = GetProject(projectName);
            var properties = ReadProperties((object)project);
            var components = new List<object>();
            foreach (dynamic component in project.VBComponents)
                components.Add(new { Name = (string)component.Name, Type = (int)component.Type });
            var references = new List<object>();
            foreach (dynamic reference in project.References)
                references.Add(new { Guid = (string)reference.GUID,
                    Major = (int)reference.Major, Minor = (int)reference.Minor,
                    IsBroken = (bool)reference.IsBroken, BuiltIn = (bool)reference.BuiltIn });
            string version = Hash(json.Serialize(new { Mode = (int)project.Mode,
                Properties = properties, Components = components, References = references }));
            return new { Project = (string)project.Name, Mode = (int)project.Mode,
                Version = version, Properties = properties, Components = components,
                References = references };
        }

        public object ComponentProperties(string projectName, string componentName)
        {
            dynamic component = GetComponent(GetProject(projectName), componentName);
            return ComponentSnapshot(projectName, component);
        }

        public object SignatureStatus(string projectName)
        {
            dynamic project = GetProject(projectName);
            if (!string.Equals(Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Host", Reason = "This host does not expose Excel.Workbook.VBASigned." };
            try
            {
                dynamic excel = Marshal.GetActiveObject("Excel.Application");
                uint excelProcessId;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out excelProcessId);
                if (excelProcessId != (uint)Process.GetCurrentProcess().Id)
                    return new { Project = projectName, Available = false, Signed = (bool?)null,
                        Source = "Excel.Workbook.VBASigned", Reason = "The registered Excel instance is not this VBE host." };
                string projectPath = null;
                try { projectPath = (string)project.FileName; }
                catch { /* An unsaved workbook may have no project path. */ }
                bool singleUnsavedProject = string.IsNullOrWhiteSpace(projectPath) &&
                    (int)excel.Workbooks.Count == 1 && (int)vbe.VBProjects.Count == 1;
                foreach (dynamic workbook in excel.Workbooks)
                {
                    bool matches = singleUnsavedProject ||
                        (!string.IsNullOrWhiteSpace(projectPath) &&
                         string.Equals(Path.GetFullPath((string)workbook.FullName),
                             Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase));
                    if (!matches) continue;
                    return new { Project = projectName, Available = true, Signed = (bool?)((bool)workbook.VBASigned),
                        Source = "Excel.Workbook.VBASigned", Reason = (string)null };
                }
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Excel.Workbook.VBASigned", Reason = "No workbook matches the selected VBE project." };
            }
            catch (Exception ex)
            {
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Excel.Workbook.VBASigned", Reason = ex.Message };
            }
        }

        public object PersistenceStatus(string projectName)
        {
            dynamic project = GetProject(projectName);
            bool projectSaved = (bool)project.Saved;
            if (!string.Equals(Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = false, HostPath = (string)null, HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null, HostHasPath = (bool?)null,
                    Reason = "The host is not Excel; its document save state is unavailable." };
            try
            {
                dynamic workbook = MatchExcelWorkbook(project, true);
                string path = (string)workbook.Path;
                bool hasPath = !string.IsNullOrWhiteSpace(path);
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = true, HostPath = hasPath ? (string)workbook.FullName : null,
                    HostSaved = (bool?)((bool)workbook.Saved),
                    HostReadOnly = (bool?)((bool)workbook.ReadOnly),
                    HostHasPath = (bool?)hasPath, Reason = (string)null };
            }
            catch (Exception ex)
            {
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = false, HostPath = (string)null, HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null, HostHasPath = (bool?)null, Reason = ex.Message };
            }
        }

        public object SaveHostDocument(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedHostPath) ||
                !Path.IsPathRooted(request.ExpectedHostPath))
                throw new ArgumentException("ExpectedHostPath must be the absolute path read from project_persistence_status.");
            if (!string.Equals(Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This host has no supported document Save API in CodexVBE.");
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            string projectPath = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath) ||
                !string.Equals(Path.GetFullPath(projectPath), Path.GetFullPath(request.ExpectedHostPath),
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected VBA project's workbook path changed since it was read.");
            dynamic workbook = MatchExcelWorkbook(project, false);
            if ((bool)workbook.ReadOnly) throw new InvalidOperationException("The workbook is read-only and cannot be saved.");
            bool beforeHostSaved = (bool)workbook.Saved;
            bool beforeProjectSaved = (bool)project.Saved;
            workbook.Save();
            bool hostSaved = (bool)workbook.Saved;
            bool projectSaved = (bool)project.Saved;
            if (!hostSaved || !projectSaved)
                throw new InvalidOperationException("Excel did not mark the workbook and VBA project as saved; a BeforeSave handler may have cancelled the save.");
            return new { Project = request.Project, HostPath = projectPath,
                SaveInvoked = true, HostSavedBefore = beforeHostSaved,
                ProjectSavedBefore = beforeProjectSaved, HostSaved = hostSaved,
                ProjectSaved = projectSaved,
                Verification = "ExcelWorkbookSaveAndSavedReadback",
                Limit = "The host reported Saved=true. Reopen the file to verify that a specific code edit persisted on disk." };
        }

        public object SaveHostDocumentAs(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Path) ||
                string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("Path and ExpectedProjectVersion are required.");
            if (!string.Equals(Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This host has no supported SaveAs API in CodexVBE.");
            string path = RequireAbsolutePath(request.Path);
            if (!string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The first Excel SaveAs supports only a macro-enabled .xlsm workbook.");
            if (File.Exists(path)) throw new IOException("SaveAs destination already exists: " + path);
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                throw new DirectoryNotFoundException("The SaveAs destination directory does not exist.");
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            dynamic workbook = MatchExcelWorkbook(project, true);
            if ((bool)workbook.ReadOnly)
                throw new InvalidOperationException("The workbook is read-only and cannot be saved.");
            string oldProjectPath = null;
            try { oldProjectPath = (string)project.FileName; }
            catch { } // VBProject.FileName can fail before the first save.
            if (!string.IsNullOrWhiteSpace((string)workbook.Path) ||
                !string.IsNullOrWhiteSpace(oldProjectPath))
                throw new InvalidOperationException("This command is only for the first save of an unsaved Excel VBA project.");
            workbook.SaveAs(path, 52); // xlOpenXMLWorkbookMacroEnabled
            string actual = Path.GetFullPath((string)workbook.FullName);
            string projectPath = Path.GetFullPath((string)project.FileName);
            if (!string.Equals(actual, path, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(projectPath, path, StringComparison.OrdinalIgnoreCase) ||
                !File.Exists(path) || !(bool)workbook.Saved || !(bool)project.Saved)
                throw new InvalidOperationException("Excel SaveAs returned without matching saved workbook and project paths.");
            return new { Project = request.Project, HostPath = actual, ProjectPath = projectPath,
                SaveAsInvoked = true, Bytes = new FileInfo(path).Length,
                HostSaved = true, ProjectSaved = true,
                Verification = "ExcelWorkbookAndProjectPathReadback",
                Limit = "Reopen the .xlsm to verify persistence of a specific VBA edit." };
        }

        private dynamic MatchExcelWorkbook(dynamic project, bool allowUnsaved)
        {
            dynamic excel = Marshal.GetActiveObject("Excel.Application");
            uint excelProcessId;
            GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out excelProcessId);
            if (excelProcessId != (uint)Process.GetCurrentProcess().Id)
                throw new InvalidOperationException("The registered Excel instance is not this VBE host.");
            string projectPath = null;
            try { projectPath = (string)project.FileName; }
            catch { }
            if (string.IsNullOrWhiteSpace(projectPath) && allowUnsaved &&
                (int)excel.Workbooks.Count == 1 && (int)vbe.VBProjects.Count == 1)
                return excel.Workbooks.Item(1);
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath))
                throw new InvalidOperationException("The project has no saved workbook path.");
            dynamic match = null;
            foreach (dynamic workbook in excel.Workbooks)
            {
                if (!string.Equals(Path.GetFullPath((string)workbook.FullName),
                    Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null) throw new InvalidOperationException("Multiple workbooks match the selected VBA project.");
                match = workbook;
            }
            if (match == null) throw new InvalidOperationException("No workbook matches the selected VBA project.");
            return match;
        }

        public object PersistExcelSignature(string projectName)
        {
            if (!string.Equals(Process.GetCurrentProcess().ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                return new { Available = false, Saved = false,
                    Reason = "The host is not Excel; save the host document with its native command." };
            dynamic project = GetProject(projectName);
            string projectPath = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath))
                throw new InvalidOperationException("The Excel VBA project has no saved workbook path.");
            dynamic excel = Marshal.GetActiveObject("Excel.Application");
            uint excelProcessId;
            GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out excelProcessId);
            if (excelProcessId != (uint)Process.GetCurrentProcess().Id)
                throw new InvalidOperationException("The registered Excel instance is not this VBE host.");
            dynamic match = null;
            foreach (dynamic workbook in excel.Workbooks)
            {
                if (!string.Equals(Path.GetFullPath((string)workbook.FullName),
                    Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null) throw new InvalidOperationException("Multiple workbooks match the selected VBA project.");
                match = workbook;
            }
            if (match == null) throw new InvalidOperationException("No workbook matches the selected VBA project.");
            if ((bool)match.ReadOnly) throw new InvalidOperationException("The signed workbook is read-only and cannot be saved.");
            if (!(bool)match.VBASigned) throw new InvalidOperationException("Excel does not report a signed VBA project before saving.");
            match.Save();
            if (!(bool)match.VBASigned) throw new InvalidOperationException("Excel no longer reports the VBA project as signed after saving.");
            return new { Available = true, Saved = true, Path = projectPath,
                Signed = true, Verification = "ExcelWorkbookSaveAndVBASignedReadback",
                Limit = "A reopening check is needed to prove the signature persisted on disk." };
        }

        public object ComponentPropertyValue(string projectName, string componentName, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property is required.");
            return ComponentProbe(projectName, componentName, "designer_property_value", propertyName);
        }

        // Bridge-only diagnostic: each stage is deliberately separate so a COM hang is attributable.
        public object ComponentProbe(string projectName, string componentName, string stage, string propertyName)
        {
            dynamic component = GetComponent(GetProject(projectName), componentName);
            switch (stage)
            {
                case "identity":
                    return new { Name = (string)component.Name, Type = (int)component.Type };
                case "descriptor_names":
                    return TypeDescriptor.GetProperties((object)component)
                        .Cast<PropertyDescriptor>().Select(item => new { item.Name,
                            Type = item.PropertyType?.FullName, item.IsReadOnly }).ToArray();
                case "designer_property_names":
                    var names = new List<string>();
                    foreach (dynamic property in component.Properties) names.Add((string)property.Name);
                    return names;
                case "code_count":
                    return new { Lines = (int)component.CodeModule.CountOfLines };
                case "code_sha":
                    dynamic module = component.CodeModule;
                    int lines = (int)module.CountOfLines;
                    return new { Lines = lines,
                        Sha256 = Hash(lines == 0 ? "" : (string)module.Lines(1, lines)) };
                case "descriptor_value":
                    PropertyDescriptor descriptor = TypeDescriptor.GetProperties((object)component).Find(propertyName, true);
                    if (descriptor == null) throw new ArgumentException("Descriptor not found: " + propertyName);
                    object value = descriptor.GetValue((object)component);
                    return new { Name = descriptor.Name, Type = descriptor.PropertyType?.FullName,
                        Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar",
                        Value = value != null && Marshal.IsComObject(value) ? null : Scalar(value) };
                case "designer_property_value":
                    if ((int)component.Type == 100 &&
                        string.Equals(propertyName, "MailEnvelope", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("MailEnvelope blocks the Excel document-component COM inspector and is not read automatically.");
                    dynamic selectedProperty = component.Properties.Item(propertyName);
                    object designerValue = selectedProperty.Value;
                    return new { Name = (string)selectedProperty.Name,
                        Kind = designerValue != null && Marshal.IsComObject(designerValue) ? "object" : "scalar",
                        Value = designerValue != null && Marshal.IsComObject(designerValue) ? null : Scalar(designerValue) };
                default:
                    throw new ArgumentException("Unsupported component probe Action.");
            }
        }

        private object ComponentSnapshot(string projectName, dynamic component)
        {
            string name = (string)component.Name;
            int type = (int)component.Type;
            var properties = ReadProperties((object)component);
            var designerProperties = new List<VbePropertyInfo>();
            var hostProperties = new List<VbePropertyInfo>();
            foreach (dynamic property in component.Properties)
            {
                var info = new VbePropertyInfo { Name = (string)property.Name };
                if (type == 100)
                {
                    info.Kind = "host property";
                    info.Display = "Value not read in bulk; use component_property_value for a named property.";
                    if (string.Equals(info.Name, "MailEnvelope", StringComparison.OrdinalIgnoreCase))
                        info.Error = "Getter blocks Excel document-component inspection.";
                    hostProperties.Add(info);
                    continue;
                }
                try
                {
                    object value = property.Value;
                    info.Type = value?.GetType().FullName;
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = Scalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                designerProperties.Add(info);
            }
            dynamic codeModule = component.CodeModule;
            int lineCount = (int)codeModule.CountOfLines;
            string code = lineCount == 0 ? "" : (string)codeModule.Lines(1, lineCount);
            string codeHash = Hash(code);
            string formVersion = null;
            if (type == 3)
            {
                dynamic state = forms.State(projectName, name);
                formVersion = (string)state.Version;
            }
            string version = Hash(json.Serialize(new { Name = name, Type = type,
                Properties = properties, DesignerProperties = designerProperties,
                HostProperties = hostProperties,
                CodeSha256 = codeHash, FormVersion = formVersion }));
            return new { Project = projectName, Component = name, Type = type,
                Version = version, CodeSha256 = codeHash, CodeLines = lineCount,
                FormVersion = formVersion, Properties = properties,
                DesignerProperties = designerProperties, HostProperties = hostProperties };
        }

        public object SetProjectProperty(Request request)
        {
            if (string.Equals(request.Property, "Name", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Project rename is disabled: it correlated with an Excel process crash during validation.");
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            SetScalar((object)project, request.Property, request.Value);
            return ProjectProperties((string)project.Name);
        }

        public object SetComponentProperty(Request request)
        {
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if (string.Equals(request.Property, "Name", StringComparison.OrdinalIgnoreCase))
                ValidateIdentifier(request.Value as string);
            SetScalar((object)component, request.Property, request.Value);
            return ComponentSnapshot(request.Project, component);
        }

        public object SetClassInstancing(Request request)
        {
            if (!(request.Value is int) && !(request.Value is long) && !(request.Value is double) &&
                !(request.Value is decimal))
                throw new ArgumentException("Value must be the number 1 (Private) or 2 (PublicNotCreatable).");
            int requested;
            try { requested = Convert.ToInt32(request.Value, CultureInfo.InvariantCulture); }
            catch (Exception ex) { throw new ArgumentException("Value must be 1 or 2.", ex); }
            if ((requested != 1 && requested != 2) ||
                Convert.ToDecimal(request.Value, CultureInfo.InvariantCulture) != requested)
                throw new ArgumentException("Value must be 1 (Private) or 2 (PublicNotCreatable).");
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            if ((int)component.Type != 2) throw new InvalidOperationException("The component must be a class module.");
            AssertComponentVersion(request, component);
            dynamic property = component.Properties.Item("Instancing");
            int before = Convert.ToInt32(property.Value, CultureInfo.InvariantCulture);
            if (before != requested) property.Value = requested;
            int actual = Convert.ToInt32(property.Value, CultureInfo.InvariantCulture);
            if (actual != requested) throw new InvalidOperationException("The VBE did not retain class Instancing.");
            return new { Project = request.Project, Class = request.Module, Before = before,
                Instancing = actual, Meaning = actual == 1 ? "Private" : "PublicNotCreatable",
                Component = ComponentSnapshot(request.Project, component) };
        }

        public object RenameComponent(Request request)
        {
            ValidateIdentifier(request.NewName);
            dynamic project = GetDesignProject(request.Project);
            foreach (dynamic existing in project.VBComponents)
                if (!string.Equals((string)existing.Name, request.Module, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            component.Name = request.NewName;
            if (!string.Equals((string)component.Name, request.NewName, StringComparison.Ordinal))
                throw new InvalidOperationException("The VBE did not retain the requested component name.");
            return ComponentSnapshot(request.Project, component);
        }

        public object RemoveComponent(Request request)
        {
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if ((int)component.Type == 100)
                throw new InvalidOperationException("Host document modules cannot be removed from their project.");
            project.VBComponents.Remove(component);
            return ProjectProperties(request.Project);
        }

        public object ImportComponent(Request request)
        {
            string path = RequireExistingPath(request.Path);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            var before = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic existing in project.VBComponents) before.Add((string)existing.Name);

            Exception importError = null;
            try { project.VBComponents.Import(path); }
            catch (Exception ex) { importError = ex; }

            // A COM error can occur after the component was added. Never call Import
            // again merely because its result or the immediate readback failed.
            var added = new List<string>();
            try
            {
                foreach (dynamic existing in project.VBComponents)
                {
                    string name = (string)existing.Name;
                    if (!before.Contains(name)) added.Add(name);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Import outcome is uncertain after COM error. " +
                    "Read list_modules before retrying; automatic re-import is disabled. " + ex.Message, ex);
            }
            if (added.Count != 1)
            {
                string cause = importError == null ? "" : " COM error: " + importError.Message;
                throw new InvalidOperationException("Import outcome is uncertain (new components: " +
                    added.Count + "). Read list_modules before retrying; automatic re-import is disabled." + cause);
            }

            string importedName = added[0];
            object importedState = TryImmediateRead(() => ComponentSnapshot(request.Project,
                GetComponent(project, importedName)), out string componentError);
            object projectState = TryImmediateRead(() => ProjectProperties(request.Project),
                out string projectError);
            bool verified = importedState != null && projectState != null;
            return new { Applied = true, Verified = verified, VerificationPending = !verified,
                ImportedName = importedName,
                Imported = importedState, Project = projectState,
                ImportError = importError?.Message, ComponentReadbackError = componentError,
                ProjectReadbackError = projectError,
                NextRead = verified ? null : "Call component_properties and project_properties in a separate request before another mutation." };
        }

        private static object TryImmediateRead(Func<object> read, out string error)
        {
            error = null;
            try { return read(); }
            catch (Exception ex) { error = ex.Message; return null; }
        }

        public object ExportComponent(Request request)
        {
            string path = RequireAbsolutePath(request.Path);
            if (File.Exists(path)) throw new IOException("Export destination already exists: " + path);
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if ((int)component.Type == 3 &&
                string.Equals(Path.GetExtension(path), ".frm", StringComparison.OrdinalIgnoreCase) &&
                File.Exists(Path.ChangeExtension(path, ".frx")))
                throw new IOException("The UserForm FRX companion destination already exists.");
            component.Export(path);
            if (!File.Exists(path)) throw new IOException("The VBE did not create the export file.");
            return new { Project = request.Project, Component = request.Module,
                Path = path, Bytes = new FileInfo(path).Length,
                ComponentState = ComponentSnapshot(request.Project, component) };
        }

        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        private dynamic GetDesignProject(string name)
        {
            dynamic project = GetProject(name);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The project must be in design mode.");
            return project;
        }

        private static dynamic GetComponent(dynamic project, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Component is required.");
            dynamic match = null;
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (match != null) throw new InvalidOperationException("Component name is ambiguous.");
                    match = component;
                }
            if (match == null) throw new InvalidOperationException("Component not found: " + name);
            return match;
        }

        private void AssertProjectVersion(Request request, dynamic project)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("ExpectedProjectVersion is required.");
            dynamic state = ProjectProperties((string)project.Name);
            if (!string.Equals((string)state.Version, request.ExpectedProjectVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since it was read.");
        }

        private void AssertComponentVersion(Request request, dynamic component)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedComponentVersion))
                throw new ArgumentException("ExpectedComponentVersion is required.");
            dynamic state = ComponentSnapshot(request.Project, component);
            if (!string.Equals((string)state.Version, request.ExpectedComponentVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The component changed since it was read.");
        }

        private static void SetScalar(object target, string name, object value)
        {
            if (string.IsNullOrWhiteSpace(name) || value == null)
                throw new ArgumentException("Property and non-null Value are required.");
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(target).Find(name, true);
            if (descriptor == null) throw new InvalidOperationException("Property is not exposed: " + name);
            if (descriptor.IsReadOnly) throw new InvalidOperationException("Property is read-only: " + name);
            Type type = descriptor.PropertyType;
            if (type == typeof(object)) type = descriptor.GetValue(target)?.GetType() ?? value.GetType();
            object converted;
            if (type == typeof(string))
            {
                if (!(value is string)) throw new ArgumentException("A string is required.");
                converted = value;
            }
            else if (type.IsEnum)
                converted = value is string ? Enum.Parse(type, (string)value, true) :
                    Enum.ToObject(type, Convert.ToInt32(value, CultureInfo.InvariantCulture));
            else if (type.IsPrimitive || type == typeof(decimal))
                converted = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
            else throw new InvalidOperationException("Property is not an editable scalar: " + name);
            descriptor.SetValue(target, converted);
            object actual = descriptor.GetValue(target);
            if (!Equals(actual, converted) &&
                !string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture),
                    Convert.ToString(converted, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The VBE did not retain property " + name + ".");
        }

        private static List<VbePropertyInfo> ReadProperties(object target)
        {
            var result = new List<VbePropertyInfo>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(target))
            {
                var info = new VbePropertyInfo { Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName, ReadOnly = descriptor.IsReadOnly };
                if (!IsSafeScalarType(descriptor.PropertyType))
                {
                    info.Kind = "object";
                    info.Display = "Object getter not invoked; use a dedicated VBIDE inspection command.";
                    result.Add(info);
                    continue;
                }
                try
                {
                    object value = descriptor.GetValue(target);
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = Scalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                result.Add(info);
            }
            return result;
        }

        private static bool IsSafeScalarType(Type type)
        {
            return type != null && (type.IsPrimitive || type.IsEnum ||
                type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime));
        }

        private static object Scalar(object value)
        {
            if (value == null || value is string || value is bool || value is byte ||
                value is short || value is int || value is long || value is float ||
                value is double || value is decimal) return value;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string RequireAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) ||
                !Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("A fully qualified file path is required.");
            return Path.GetFullPath(path);
        }

        private static string RequireExistingPath(string path)
        {
            string fullPath = RequireAbsolutePath(path);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Import file not found.", fullPath);
            return fullPath;
        }

        private static void ValidateIdentifier(string name)
        {
            if (string.IsNullOrWhiteSpace(name) ||
                !Regex.IsMatch(name, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Name must start with a letter and contain at most 40 letters, digits or underscores.");
        }

        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
