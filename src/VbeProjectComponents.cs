using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class VbeProjectComponents
    {
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
            foreach (dynamic property in component.Properties)
            {
                var info = new VbePropertyInfo { Name = (string)property.Name };
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
                CodeSha256 = codeHash, FormVersion = formVersion }));
            return new { Project = projectName, Component = name, Type = type,
                Version = version, CodeSha256 = codeHash, CodeLines = lineCount,
                FormVersion = formVersion, Properties = properties,
                DesignerProperties = designerProperties };
        }

        public object SetProjectProperty(Request request)
        {
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            if (string.Equals(request.Property, "Name", StringComparison.OrdinalIgnoreCase))
                ValidateIdentifier(request.Value as string);
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

        public object RenameProject(Request request)
        {
            ValidateIdentifier(request.NewName);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            foreach (dynamic existing in vbe.VBProjects)
                if (!string.Equals((string)existing.Name, request.Project, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A project with this name already exists.");
            project.Name = request.NewName;
            if (!string.Equals((string)project.Name, request.NewName, StringComparison.Ordinal))
                throw new InvalidOperationException("The VBE did not retain the requested project name.");
            return ProjectProperties(request.NewName);
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
            dynamic component = project.VBComponents.Import(path);
            return new { Imported = ComponentSnapshot(request.Project, component),
                Project = ProjectProperties(request.Project) };
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
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Project is required.");
            dynamic match = null;
            foreach (dynamic project in vbe.VBProjects)
                if (string.Equals((string)project.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (match != null) throw new InvalidOperationException("Project name is ambiguous.");
                    match = project;
                }
            if (match == null) throw new InvalidOperationException("Project not found: " + name);
            return match;
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
