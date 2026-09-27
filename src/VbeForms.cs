using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        private readonly dynamic vbe;
        private static readonly HashSet<string> BuiltInControls = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Forms.CheckBox.1", "Forms.ComboBox.1", "Forms.CommandButton.1", "Forms.Frame.1",
            "Forms.Image.1", "Forms.Label.1", "Forms.ListBox.1", "Forms.MultiPage.1",
            "Forms.OptionButton.1", "Forms.ScrollBar.1", "Forms.SpinButton.1",
            "Forms.TabStrip.1", "Forms.TextBox.1", "Forms.ToggleButton.1"
        };

        public VbeForms(object vbe) { this.vbe = vbe; }

        public object ControlTypes() { return VbeControlCatalog.List(BuiltInControls); }

        public object List(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                if ((int)component.Type == 3)
                    result.Add(new { Name = (string)component.Name, DesignerOpen = (bool)component.HasOpenDesigner });
            return result;
        }

        public object State(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            return Snapshot(projectName, form);
        }

        public object Properties(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            return DescribeProperties(form);
        }

        public object Tree(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            List<object> nodes;
            List<VbePropertyInfo> properties;
            int nodeCount;
            string version = TreeVersion(form, out nodes, out properties, out nodeCount);
            return new { Project = projectName, Form = formName, FormVersion = version, TreeVersion = version,
                NodeCount = nodeCount, Properties = properties, Controls = nodes };
        }

        public object EventCatalog(string projectName, string formName, string controlPath)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            dynamic tree = Tree(projectName, formName);
            bool userForm = string.IsNullOrWhiteSpace(controlPath) ||
                string.Equals(controlPath, "UserForm", StringComparison.OrdinalIgnoreCase);
            if (!userForm && !TreeContainsPath((IEnumerable)tree.Controls, controlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = userForm ? (object)form.Designer : ResolveTreeItem(form.Designer, controlPath);
            object catalog = VbeComEvents.Read(target);
            return new { Project = projectName, Form = formName,
                ControlPath = userForm ? "UserForm" : controlPath,
                ObjectName = userForm ? "UserForm" : (string)((dynamic)target).Name,
                Type = TypeDescriptor.GetClassName(target),
                TreeVersion = (string)tree.TreeVersion, Catalog = catalog };
        }

        private static string TreeVersion(dynamic form, out List<object> nodes,
            out List<VbePropertyInfo> properties, out int nodeCount)
        {
            nodeCount = 0;
            object designer = form.Designer;
            nodes = ReadChildControls(((dynamic)designer).Controls, designer, (string)form.Name,
                "Controls", 0, ref nodeCount);
            properties = DescribeProperties(form);
            string json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }
                .Serialize(new { Form = (string)form.Name, Properties = properties, Controls = nodes });
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(json)))
                    .Replace("-", "").ToLowerInvariant();
        }

        public object ParentProbe(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            object designer = form.Designer;
            var rows = new List<object>();
            foreach (dynamic control in form.Designer.Controls)
            {
                object parent = control.Parent;
                rows.Add(new { Control = (string)control.Name,
                    ParentName = SafeComName(parent), ParentType = TypeDescriptor.GetClassName(parent),
                    DesignerName = SafeComName(designer), DesignerType = TypeDescriptor.GetClassName(designer),
                    SameDesigner = SameComIdentity(parent, designer),
                    SameComponent = SameComIdentity(parent, (object)form) });
            }
            return new { Project = projectName, Form = formName, Rows = rows };
        }

        private static string SafeComName(object item)
        {
            try { return (string)((dynamic)item).Name; }
            catch { return null; }
        }

        private static List<object> ReadChildControls(dynamic collection, object owner, string formName,
            string path, int depth, ref int nodeCount)
        {
            var result = new List<object>();
            foreach (dynamic control in collection)
                if (SameContainer((object)control.Parent, owner, formName))
                    result.Add(ReadTreeNode(control, "Control", formName, path, depth, ref nodeCount));
            return result;
        }

        private static bool SameContainer(object actualParent, object expectedOwner, string formName)
        {
            if (SameComIdentity(actualParent, expectedOwner)) return true;
            string actualPath = ContainerIdentity(actualParent, formName);
            string expectedPath = ContainerIdentity(expectedOwner, formName);
            return actualPath != null && string.Equals(actualPath, expectedPath, StringComparison.OrdinalIgnoreCase);
        }

        private static string ContainerIdentity(object item, string formName)
        {
            if (item == null) return null;
            var path = new List<string>();
            for (int depth = 0; depth < 16 && item != null; depth++)
            {
                string type = TypeDescriptor.GetClassName(item);
                string name = SafeComName(item);
                if (string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrEmpty(name)) name = formName;
                if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(name)) return null;
                path.Add(type + ":" + name);
                if (string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase))
                    return string.Join("/", path);
                try { item = ((dynamic)item).Parent; }
                catch { return null; }
            }
            return null;
        }

        private static bool SameComIdentity(object left, object right)
        {
            if (left == null || right == null) return false;
            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right))
                return ReferenceEquals(left, right);
            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            finally
            {
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            }
        }

        private static object ReadTreeNode(object item, string kind, string formName,
            string parentPath, int depth, ref int nodeCount)
        {
            if (++nodeCount > 512 || depth > 16)
                throw new InvalidOperationException("The UserForm control hierarchy exceeds the inspection limit.");
            dynamic current = item;
            string name = (string)current.Name;
            string path = parentPath + "/" + name;
            var children = new List<object>();
            PropertyDescriptorCollection descriptors = TypeDescriptor.GetProperties(item);
            if (kind == "Control")
            {
                PropertyDescriptor controls = descriptors.Find("Controls", true);
                if (controls != null)
                {
                    object nested = controls.GetValue(item);
                    if (nested != null)
                        children.AddRange(ReadChildControls(nested, item, formName,
                            path + "/Controls", depth + 1, ref nodeCount));
                }
                PropertyDescriptor pages = descriptors.Find("Pages", true);
                if (pages != null)
                    foreach (dynamic page in (dynamic)pages.GetValue(item))
                        children.Add(ReadTreeNode(page, "Page", formName, path + "/Pages", depth + 1, ref nodeCount));
                PropertyDescriptor tabs = descriptors.Find("Tabs", true);
                if (tabs != null)
                    foreach (dynamic tab in (dynamic)tabs.GetValue(item))
                        children.Add(ReadTreeNode(tab, "Tab", formName, path + "/Tabs", depth + 1, ref nodeCount));
            }
            else if (kind == "Page")
            {
                PropertyDescriptor controls = descriptors.Find("Controls", true);
                if (controls != null)
                    children.AddRange(ReadChildControls(controls.GetValue(item), item, formName,
                        path + "/Controls", depth + 1, ref nodeCount));
            }
            return new { Path = path, Name = name, Kind = kind,
                Type = TypeDescriptor.GetClassName(item), Properties = ReadObjectProperties(item), Children = children };
        }

        private static List<VbePropertyInfo> ReadObjectProperties(object item)
        {
            var result = new List<VbePropertyInfo>();
            string targetType = TypeDescriptor.GetClassName(item);
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(item))
            {
                var info = new VbePropertyInfo { Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName, ReadOnly = descriptor.IsReadOnly,
                    AllowedValues = EnumChoices(descriptor.PropertyType),
                    SetterStatus = DesignerSetterStatus(targetType, descriptor) };
                try
                {
                    object value = descriptor.GetValue(item);
                    info.Kind = value != null && (Marshal.IsComObject(value) || value is Font || value is Image)
                        ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = NormalizeScalar(value);
                    else if (value is Image) info.Digest = ImageDigest((Image)value);
                    else if (value is Font)
                    {
                        try { info.Members = DescribeObjectMembers((object)((dynamic)item).Font); }
                        catch (Exception ex) { info.Error = ex.Message; }
                    }
                }
                catch (Exception ex) { info.Error = ex.Message; }
                result.Add(info);
            }
            return result;
        }

        public object SetProperty(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Property) || request.Value == null)
                throw new ArgumentException("Property and a non-null Value are required.");
            dynamic project = GetDesignProject(request.Project);
            dynamic form = GetForm(project, request.Form);
            AssertVersion(request, form);
            string[] path = request.Property.Split('.');
            if (path.Length < 1 || path.Length > 2 || path.Any(part => string.IsNullOrWhiteSpace(part)))
                throw new ArgumentException("Use a form property name or one member path such as Font.Name.");
            string propertyName = path[0];
            if (path.Length == 1 && string.Equals(propertyName, "Name", StringComparison.OrdinalIgnoreCase))
            {
                string newName = request.Value as string;
                if (string.IsNullOrWhiteSpace(newName) || !Regex.IsMatch(newName, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                    throw new ArgumentException("Name must start with a letter and contain at most 40 letters, digits or underscores.");
                foreach (dynamic component in project.VBComponents)
                    if (!string.Equals((string)component.Name, (string)form.Name, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals((string)component.Name, newName, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("A component with this name already exists.");
                form.Name = newName;
            }
            else
            {
                dynamic property = null;
                foreach (dynamic candidate in form.Properties)
                    if (string.Equals((string)candidate.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    { property = candidate; break; }
                if (property == null) throw new InvalidOperationException("Unknown UserForm property: " + propertyName);
                if ((int)property.NumIndices != 0)
                    throw new InvalidOperationException("Indexed properties need an indexed editor; this path cannot be written as a scalar.");
                if (path.Length == 1)
                {
                    PropertyDescriptor descriptor = TypeDescriptor.GetProperties((object)form.Designer).Find(propertyName, true);
                    if (descriptor != null && descriptor.IsReadOnly)
                        throw new InvalidOperationException("This UserForm property is read-only: " + propertyName);
                    object previous = null;
                    try { previous = property.Value; }
                    catch (Exception ex) { throw new InvalidOperationException("Cannot read the current value of " + propertyName + ": " + ex.Message); }
                    if (previous != null && Marshal.IsComObject(previous))
                        throw new InvalidOperationException("Object property: use a member path such as Font.Name or a dedicated object command.");
                    // VBIDE exposes Tag as an untyped null Variant until it is assigned.
                    Type targetType = previous?.GetType() ?? descriptor?.PropertyType;
                    if (targetType == null && string.Equals(propertyName, "Tag", StringComparison.OrdinalIgnoreCase))
                        targetType = typeof(string);
                    if (targetType == null || targetType == typeof(object))
                        throw new InvalidOperationException("Cannot determine a safe scalar type for " + propertyName);
                    property.Value = ConvertScalar(request.Value, targetType);
                }
                else
                {
                    if (!string.Equals(propertyName, "Font", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("This object property needs a dedicated editor; scalar member writes are unavailable.");
                    SetFormFontMember(form.Designer, path[1], request.Value);
                }
            }
            string currentName = (string)form.Name;
            var properties = DescribeProperties(form);
            return new { Project = request.Project, Form = currentName, Property = request.Property,
                Properties = properties, State = Snapshot(request.Project, form) };
        }

        public object SetPicture(Request request)
        {
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            object picture = OlePictureLoader.Load(request.Path);
            dynamic designer = form.Designer;
            designer.Picture = picture;
            object installed = designer.Picture;
            string expected = OlePictureLoader.Fingerprint(picture);
            string actual = OlePictureLoader.Fingerprint(installed);
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                throw new InvalidOperationException("The UserForm did not retain the requested OLE picture.");
            return new { Project = request.Project, Form = request.Form,
                Picture = actual, Properties = DescribeProperties(form), State = Snapshot(request.Project, form) };
        }

        private static object ConvertScalar(object value, Type targetType)
        {
            if (targetType == null || targetType == typeof(object) || targetType.IsArray ||
                (!targetType.IsPrimitive && targetType != typeof(string) && targetType != typeof(decimal) && !targetType.IsEnum))
                throw new InvalidOperationException("This property is not a supported scalar type.");
            if (targetType.IsEnum)
            {
                if (value is string) return Enum.Parse(targetType, (string)value, true);
                return Enum.ToObject(targetType, Convert.ToInt32(value, CultureInfo.InvariantCulture));
            }
            if (targetType == typeof(bool) && value is string)
            {
                bool parsed;
                if (!bool.TryParse((string)value, out parsed))
                    throw new ArgumentException("Boolean value must be true or false.");
                return parsed;
            }
            if (targetType == typeof(string))
            {
                if (!(value is string)) throw new ArgumentException("This property requires a string.");
                return value;
            }
            if (targetType == typeof(float) || targetType == typeof(double) || targetType == typeof(decimal))
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(number) || double.IsInfinity(number))
                    throw new ArgumentException("Numeric value must be finite.");
            }
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }

        private static void SetFormFontMember(dynamic designer, string member, object value)
        {
            dynamic font = designer.Font;
            if (string.Equals(member, "Name", StringComparison.OrdinalIgnoreCase))
            {
                string name = value as string;
                if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Font.Name requires a nonempty string.");
                font.Name = name;
                if (!string.Equals((string)font.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The VBE did not retain Font.Name.");
            }
            else if (string.Equals(member, "Size", StringComparison.OrdinalIgnoreCase))
            {
                double size = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0 || size > 200)
                    throw new ArgumentException("Font.Size must be greater than 0 and at most 200.");
                font.Size = size;
                if (Math.Abs(Convert.ToDouble(font.Size, CultureInfo.InvariantCulture) - size) > 0.01)
                    throw new InvalidOperationException("The VBE did not retain Font.Size.");
            }
            else
            {
                bool enabled = (bool)ConvertScalar(value, typeof(bool));
                if (string.Equals(member, "Bold", StringComparison.OrdinalIgnoreCase))
                { font.Bold = enabled; if ((bool)font.Bold != enabled) throw new InvalidOperationException("The VBE did not retain Font.Bold."); }
                else if (string.Equals(member, "Italic", StringComparison.OrdinalIgnoreCase))
                { font.Italic = enabled; if ((bool)font.Italic != enabled) throw new InvalidOperationException("The VBE did not retain Font.Italic."); }
                else if (string.Equals(member, "Underline", StringComparison.OrdinalIgnoreCase))
                { font.Underline = enabled; if ((bool)font.Underline != enabled) throw new InvalidOperationException("The VBE did not retain Font.Underline."); }
                else if (string.Equals(member, "Strikethrough", StringComparison.OrdinalIgnoreCase))
                { font.Strikethrough = enabled; if ((bool)font.Strikethrough != enabled) throw new InvalidOperationException("The VBE did not retain Font.Strikethrough."); }
                else throw new InvalidOperationException("Unsupported Font member: " + member);
            }
        }

        private static List<VbePropertyInfo> DescribeProperties(dynamic form)
        {
            var result = new List<VbePropertyInfo>();
            object designer = form.Designer;
            PropertyDescriptorCollection descriptors = TypeDescriptor.GetProperties(designer);
            foreach (dynamic property in form.Properties)
            {
                string name = (string)property.Name;
                PropertyDescriptor descriptor = descriptors.Find(name, true);
                var info = new VbePropertyInfo { Name = name, Type = descriptor?.PropertyType?.FullName,
                    ReadOnly = descriptor == null ? (bool?)null : descriptor.IsReadOnly,
                    AllowedValues = EnumChoices(descriptor?.PropertyType),
                    SetterStatus = descriptor == null ? "Unknown" : DesignerSetterStatus("UserForm", descriptor) };
                if (string.Equals(name, "Picture", StringComparison.OrdinalIgnoreCase))
                {
                    info.Kind = "object";
                    info.Type = "stdole.IPictureDisp";
                    try
                    {
                        object installed = ((dynamic)designer).Picture;
                        info.Digest = OlePictureLoader.Fingerprint(installed);
                        info.Display = installed == null ? "(empty)" : info.Digest;
                    }
                    catch (Exception ex) { info.Error = ex.Message; }
                    result.Add(info);
                    continue;
                }
                try { info.NumIndices = (int)property.NumIndices; }
                catch (Exception ex) { info.Error = ex.Message; }
                object raw = null;
                try { raw = property.Value; }
                catch (Exception ex) { info.Error = ex.Message; }
                object managed = null;
                if (descriptor != null)
                {
                    try { managed = descriptor.GetValue(designer); }
                    catch (Exception ex) { if (info.Error == null) info.Error = ex.Message; }
                }
                object inspected = managed ?? raw;
                if (inspected != null && info.Type == null) info.Type = inspected.GetType().FullName;
                if (info.NumIndices > 0) info.Kind = "indexed";
                else if (inspected is Image)
                {
                    info.Kind = "object";
                    info.Display = inspected.GetType().Name;
                    info.Digest = ImageDigest((Image)inspected);
                    info.Members = DescribeObjectMembers(inspected);
                }
                else if (inspected != null &&
                    (Marshal.IsComObject(inspected) || inspected is Font || inspected is System.Collections.IEnumerable && !(inspected is string)))
                {
                    info.Kind = "object";
                    info.Display = inspected.GetType().Name;
                    info.Members = DescribeObjectMembers(inspected);
                }
                else
                {
                    info.Kind = "scalar";
                    info.Value = NormalizeScalar(raw ?? managed);
                    if (managed is Color) info.Display = ((Color)managed).Name;
                }
                result.Add(info);
            }
            return result;
        }

        private static List<VbePropertyInfo> DescribeObjectMembers(object source)
        {
            var members = new List<VbePropertyInfo>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(source))
            {
                if (members.Count >= 64) break;
                var info = new VbePropertyInfo { Name = descriptor.Name, Type = descriptor.PropertyType?.FullName,
                    ReadOnly = descriptor.IsReadOnly, AllowedValues = EnumChoices(descriptor.PropertyType) };
                try
                {
                    object value = descriptor.GetValue(source);
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = NormalizeScalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                members.Add(info);
            }
            return members;
        }

        private static string[] EnumChoices(Type type)
        {
            if (type == null || !type.IsEnum) return null;
            try { return Enum.GetNames(type); }
            catch { return null; }
        }

        private static string DesignerSetterStatus(string targetType, PropertyDescriptor descriptor)
        {
            if (descriptor.IsReadOnly) return "DescriptorReadOnly";
            if (string.Equals(descriptor.Name, "_Font_Reserved", StringComparison.OrdinalIgnoreCase))
                return "GetterUnavailable";
            if (string.Equals(targetType, "Label", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "Cancel", StringComparison.OrdinalIgnoreCase))
                return "BlockedNativeSetterFailure";
            if (string.Equals(targetType, "ToggleButton", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "Value", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "TextBox", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "ScrollBars", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "ColumnCount", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "SpinButton", StringComparison.OrdinalIgnoreCase) &&
                new[] { "Min", "Max", "Value", "Delay", "SmallChange" }
                    .Any(name => string.Equals(descriptor.Name, name, StringComparison.OrdinalIgnoreCase)))
                return "BlockedAfterHostCrash";
            return "DescriptorCandidateUnverified";
        }

        private static object NormalizeScalar(object value)
        {
            if (value == null) return null;
            if (value is string || value is bool || value is byte || value is sbyte ||
                value is short || value is ushort || value is int || value is uint ||
                value is long || value is ulong || value is float || value is double || value is decimal)
                return value;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string ImageDigest(Image image)
        {
            try
            {
                using (var memory = new MemoryStream())
                using (var sha = SHA256.Create())
                {
                    image.Save(memory, ImageFormat.Png);
                    return BitConverter.ToString(sha.ComputeHash(memory.ToArray())).Replace("-", "").ToLowerInvariant();
                }
            }
            catch { return null; }
        }

        public object ControlProperties(string projectName, string formName, string controlName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            object control = GetControl(form.Designer, controlName);
            string targetType = TypeDescriptor.GetClassName(control);
            var result = new List<object>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(control))
            {
                object value = null;
                string error = null;
                try
                {
                    object raw = descriptor.GetValue(control);
                    if (raw != null && !Marshal.IsComObject(raw))
                        value = Convert.ToString(raw, CultureInfo.InvariantCulture);
                }
                catch (Exception ex) { error = ex.Message; }
                result.Add(new { Name = descriptor.Name,
                    Type = descriptor.PropertyType == null ? null : descriptor.PropertyType.FullName,
                    ReadOnly = descriptor.IsReadOnly,
                    SetterStatus = DesignerSetterStatus(targetType, descriptor),
                    AllowedValues = EnumChoices(descriptor.PropertyType),
                    Value = value, Error = error });
            }
            return result;
        }

        public object Create(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Form) ||
                !Regex.IsMatch(request.Form, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Form must start with a letter and contain at most 40 letters, digits or underscores.");
            dynamic project = GetDesignProject(request.Project);
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, request.Form, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");
            dynamic form = project.VBComponents.Add(3); // vbext_ct_MSForm
            form.Name = request.Form;
            form.DesignerWindow().Visible = true;
            return Snapshot(request.Project, form);
        }

        public object Open(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            form.DesignerWindow().Visible = true;
            return Snapshot(projectName, form);
        }

        public object AddControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            if (!VbeControlCatalog.IsCandidate(request.ControlType ?? string.Empty, BuiltInControls))
                throw new ArgumentException("ControlType must be a native MSForms or installed x64 CATID_Control ProgID.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic designer = form.Designer;
            foreach (dynamic existing in designer.Controls)
                if (string.Equals((string)existing.Name, request.Control, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with this name already exists.");
            dynamic control = designer.Controls.Add(request.ControlType, request.Control, true);
            control.Left = request.Left;
            control.Top = request.Top;
            control.Width = request.Width;
            control.Height = request.Height;
            if (request.Caption != null) control.Caption = request.Caption;
            form.DesignerWindow().Visible = true;
            return Snapshot(request.Project, form);
        }

        public object SetControlGeometry(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            control.Left = request.Left;
            control.Top = request.Top;
            control.Width = request.Width;
            control.Height = request.Height;
            form.DesignerWindow().Visible = true;
            return Snapshot(request.Project, form);
        }

        public object RenameControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateName(request.NewName, "NewName");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic designer = form.Designer;
            foreach (dynamic existing in designer.Controls)
                if (string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with the new name already exists.");
            dynamic control = GetControl(designer, request.Control);
            control.Name = request.NewName;
            return Snapshot(request.Project, form);
        }

        public object SetControlCaption(Request request)
        {
            if (request.Caption == null) throw new ArgumentException("Caption is required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            control.Caption = request.Caption;
            return Snapshot(request.Project, form);
        }

        public object SetControlFont(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.FontName) || !IsFinite(request.FontSize) ||
                request.FontSize <= 0 || request.FontSize > 200)
                throw new ArgumentException("FontName and a FontSize between 0 and 200 are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            dynamic font = control.Font;
            font.Name = request.FontName;
            font.Size = request.FontSize;
            font.Bold = request.FontBold;
            return Snapshot(request.Project, form);
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

        private dynamic GetDesignProject(string name)
        {
            dynamic project = GetProject(name);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The project must be in design mode.");
            return project;
        }

        private static dynamic GetForm(dynamic project, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Form is required.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if ((int)component.Type != 3) throw new InvalidOperationException("The component is not a UserForm.");
                    return component;
                }
            throw new InvalidOperationException("UserForm not found: " + name);
        }

        private static dynamic GetControl(dynamic designer, string name)
        {
            foreach (dynamic control in designer.Controls)
                if (string.Equals((string)control.Name, name, StringComparison.OrdinalIgnoreCase)) return control;
            throw new InvalidOperationException("Control not found: " + name);
        }

        private static void ValidateName(string name, string label)
        {
            if (string.IsNullOrWhiteSpace(name) || !Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                throw new ArgumentException(label + " must be a VBA identifier.");
        }

        private static void ValidateGeometry(Request request)
        {
            if (!IsFinite(request.Left) || !IsFinite(request.Top) || !IsFinite(request.Width) || !IsFinite(request.Height) ||
                request.Left < 0 || request.Top < 0 || request.Width <= 0 || request.Height <= 0 ||
                request.Left > 32767 || request.Top > 32767 || request.Width > 32767 || request.Height > 32767)
                throw new ArgumentException("Control geometry must use finite nonnegative point coordinates and positive sizes.");
        }

        private static bool IsFinite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }

        private static void AssertVersion(Request request, dynamic form)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedFormVersion))
                throw new ArgumentException("ExpectedFormVersion is required for form edits.");
            string current = Version(form);
            if (!string.Equals(current, request.ExpectedFormVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form changed since it was read.");
        }

        private static object Snapshot(string projectName, dynamic form)
        {
            dynamic designer = form.Designer;
            var controls = new List<object>();
            foreach (dynamic control in designer.Controls)
            {
                if (!SameContainer((object)control.Parent, (object)designer, (string)form.Name)) continue;
                string caption = null;
                try { caption = (string)control.Caption; } catch { }
                string fontName = null;
                double? fontSize = null;
                bool? fontBold = null;
                try
                {
                    dynamic font = control.Font;
                    fontName = (string)font.Name;
                    fontSize = Convert.ToDouble(font.Size, CultureInfo.InvariantCulture);
                    fontBold = (bool)font.Bold;
                }
                catch { }
                controls.Add(new { Name = (string)control.Name, Caption = caption,
                    Left = (double)control.Left, Top = (double)control.Top,
                    Width = (double)control.Width, Height = (double)control.Height,
                    FontName = fontName, FontSize = fontSize, FontBold = fontBold });
            }
            return new { Project = projectName, Form = (string)form.Name,
                Caption = (string)form.Properties.Item("Caption").Value,
                Width = Convert.ToDouble(form.Properties.Item("Width").Value, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(form.Properties.Item("Height").Value, CultureInfo.InvariantCulture),
                Version = Version(form), Controls = controls };
        }

        private static string Version(dynamic form)
        {
            List<object> nodes;
            List<VbePropertyInfo> properties;
            int nodeCount;
            return TreeVersion(form, out nodes, out properties, out nodeCount);
        }

        // Bridge-only probe until hierarchical identity and revision behavior are tested in Excel.
        public object AddNestedControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            if (!VbeControlCatalog.IsCandidate(request.ControlType ?? string.Empty, BuiltInControls))
                throw new ArgumentException("ControlType must be a native MSForms or installed x64 CATID_Control ProgID.");
            if (string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ExpectedTreeVersion is required from form_tree.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ParentPath))
                throw new InvalidOperationException("ParentPath is not a canonical path in form_tree.");
            dynamic controls = ResolveNestedControls(form.Designer, request.ParentPath);
            foreach (dynamic existing in controls)
                if (string.Equals((string)existing.Name, request.Control, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with this name already exists in the container.");
            dynamic control = controls.Add(request.ControlType, request.Control, true);
            try
            {
                control.Left = request.Left;
                control.Top = request.Top;
                control.Width = request.Width;
                control.Height = request.Height;
                if (request.Caption != null) control.Caption = request.Caption;
            }
            catch
            {
                try { controls.Remove(request.Control); } catch { }
                throw;
            }
            dynamic after = Tree(request.Project, request.Form);
            if (string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The nested control was not reflected in the UserForm tree.");
            return after;
        }

        public object SetNodeProperty(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.Property) || request.Value == null ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath, Property, Value and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string[] propertyPath = request.Property.Split('.');
            if (propertyPath.Length < 1 || propertyPath.Length > 2 ||
                propertyPath.Any(part => string.IsNullOrWhiteSpace(part)))
                throw new ArgumentException("Property must be a property name or one object member path.");
            PropertyDescriptor root = TypeDescriptor.GetProperties(target).Find(propertyPath[0], true);
            if (root == null) throw new InvalidOperationException("Property is not exposed: " + propertyPath[0]);
            if (string.Equals(root.Name, "_Font_Reserved", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("_Font_Reserved is a COM reserved member whose getter is unavailable in the tested VBE.");
            if (propertyPath.Length == 1)
            {
                // A ComboBox.ColumnCount=2 write followed by three AddItem
                // calls was followed by Excel heap corruption. The exact
                // trigger in that sequence has not been isolated.
                if (string.Equals(TypeDescriptor.GetClassName(target), "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(root.Name, "ColumnCount", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "ComboBox.ColumnCount editing is temporarily disabled: Excel crashed after a multicolumn design-time write sequence.");
                // A single TextBox.ScrollBars=Vertical write was followed by
                // Excel heap corruption after a successful readback.
                if (string.Equals(TypeDescriptor.GetClassName(target), "TextBox", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(root.Name, "ScrollBars", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "TextBox.ScrollBars editing is temporarily disabled: Excel crashed after a design-time write.");
                // A disposable Excel session crashed with heap corruption
                // shortly after a sequence of these SpinButton setters. The
                // responsible member has not yet been isolated.
                if (string.Equals(TypeDescriptor.GetClassName(target), "SpinButton", StringComparison.OrdinalIgnoreCase) &&
                    new[] { "Min", "Max", "Value", "Delay", "SmallChange" }
                        .Any(name => string.Equals(root.Name, name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException(
                        "SpinButton." + root.Name + " editing is temporarily disabled: Excel crashed after a design-time property write sequence containing this member.");
                // Two disposable Excel sessions crashed during teardown after
                // ToggleButton.Value=true was set in the designer. Until the
                // host interaction is isolated, refuse every Value write for
                // this type before invoking its COM setter.
                if (string.Equals(root.Name, "Value", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(TypeDescriptor.GetClassName(target), "ToggleButton", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "ToggleButton.Value editing is temporarily disabled: Excel crashed during teardown after a design-time Value=true write.");
                // A real Excel Label exposed Cancel as writable through
                // PropertyDescriptor, but the COM setter returned member-not-found.
                // Reject the known bad path before invoking native COM.
                if (string.Equals(root.Name, "Cancel", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(TypeDescriptor.GetClassName(target), "Label", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Cancel has no usable designer setter in the tested Excel VBE.");
                if (root.IsReadOnly) throw new InvalidOperationException("Property is read-only: " + root.Name);
                object oldValue = root.GetValue(target);
                object converted = ConvertDescriptorValue(request.Value, root.PropertyType, oldValue);
                root.SetValue(target, converted);
                object actual = root.GetValue(target);
                if (!SameDescriptorValue(actual, converted))
                    throw new InvalidOperationException("The VBE did not retain property " + root.Name + ".");
            }
            else
            {
                object owner = string.Equals(root.Name, "Font", StringComparison.OrdinalIgnoreCase)
                    ? (object)((dynamic)target).Font : root.GetValue(target);
                if (owner == null || !Marshal.IsComObject(owner))
                    throw new InvalidOperationException("The object property is not exposed as an editable COM object.");
                PropertyDescriptor member = TypeDescriptor.GetProperties(owner).Find(propertyPath[1], true);
                if (member == null) throw new InvalidOperationException("Object member is not exposed: " + request.Property);
                if (member.IsReadOnly) throw new InvalidOperationException("Object member is read-only: " + request.Property);
                object oldValue = member.GetValue(owner);
                object converted = ConvertDescriptorValue(request.Value, member.PropertyType, oldValue);
                member.SetValue(owner, converted);
                if (!SameDescriptorValue(member.GetValue(owner), converted))
                    throw new InvalidOperationException("The VBE did not retain object member " + request.Property + ".");
            }
            dynamic after = Tree(request.Project, request.Form);
            return new { ControlPath = request.ControlPath, Property = request.Property, Tree = after };
        }

        public object SetNodePicture(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.Property) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath, Property and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(target).Find(request.Property, true);
            if (descriptor == null || descriptor.IsReadOnly ||
                (descriptor.PropertyType != typeof(Bitmap) && descriptor.PropertyType != typeof(Icon)))
                throw new InvalidOperationException("The selected node has no writable OLE image property by that name.");
            object picture = OlePictureLoader.Load(request.Path);
            target.GetType().InvokeMember(descriptor.Name, BindingFlags.SetProperty,
                null, target, new[] { picture });
            object installed = target.GetType().InvokeMember(descriptor.Name, BindingFlags.GetProperty,
                null, target, null);
            if (!string.Equals(OlePictureLoader.Fingerprint(installed), OlePictureLoader.Fingerprint(picture),
                StringComparison.Ordinal))
                throw new InvalidOperationException("The VBE did not retain the requested OLE image.");
            return new { ControlPath = request.ControlPath, Property = descriptor.Name,
                Tree = Tree(request.Project, request.Form) };
        }

        public object RemoveControl(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 ||
                !string.Equals(parts[parts.Length - 2], "Controls", StringComparison.Ordinal))
                throw new ArgumentException("ControlPath must identify a control, not a Page or Tab.");
            string name = parts[parts.Length - 1];
            object owner = parts.Length == 2 ? (object)form.Designer :
                ResolveTreeItem(form.Designer, string.Join("/", parts.Take(parts.Length - 2)));
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(owner).Find("Controls", true);
            if (descriptor == null)
                throw new InvalidOperationException("The selected parent has no Controls collection.");
            dynamic controls = descriptor.GetValue(owner);
            controls.Remove(name);
            dynamic after = Tree(request.Project, request.Form);
            if (TreeContainsPath((IEnumerable)after.Controls, request.ControlPath) ||
                string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The control removal was not reflected in the UserForm tree.");
            return new { RemovedPath = request.ControlPath, Applied = true, Tree = after };
        }

        public object ZOrderControl(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            if (request.ZPosition != 0 && request.ZPosition != 1)
                throw new ArgumentOutOfRangeException("ZPosition", "Use 0 for front or 1 for back.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts[parts.Length - 2] != "Controls")
                throw new ArgumentException("ControlPath must identify a control, not a Page or Tab.");
            object control = ResolveTreeItem(form.Designer, request.ControlPath);
            ((dynamic)control).ZOrder(request.ZPosition);
            dynamic after = Tree(request.Project, request.Form);
            if (!TreeContainsPath((IEnumerable)after.Controls, request.ControlPath))
                throw new InvalidOperationException("The control is no longer present after ZOrder.");
            return new { ControlPath = request.ControlPath, ZPosition = request.ZPosition,
                Executed = true, Verification = "Unverified",
                VerificationPending = true,
                VerificationLimit = "MSForms does not expose z-order through Controls or form_tree; compare the visible overlap in the designer.",
                TreeBefore = before, TreeAfter = after };
        }

        public object AddPageOrTab(Request request, string collectionName)
        {
            ValidateName(request.NewName, "NewName");
            if (string.IsNullOrWhiteSpace(request.ParentPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ParentPath and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ParentPath))
                throw new InvalidOperationException("ParentPath is not a canonical path in form_tree.");
            object parent = ResolveTreeItem(form.Designer, request.ParentPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(parent).Find(collectionName, true);
            if (descriptor == null ||
                !string.Equals(TypeDescriptor.GetClassName(parent),
                    collectionName == "Pages" ? "MultiPage" : "TabStrip",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected parent is not a compatible MultiPage or TabStrip.");
            dynamic collection = descriptor.GetValue(parent);
            foreach (dynamic item in collection)
                if (string.Equals((string)item.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("An item with this name already exists in the collection.");
            int count = (int)collection.Count;
            if (request.InsertIndex.HasValue &&
                (request.InsertIndex.Value < 0 || request.InsertIndex.Value > count))
                throw new ArgumentOutOfRangeException("InsertIndex", "InsertIndex must be between zero and Count.");
            string caption = request.Caption ?? request.NewName;
            if (request.InsertIndex.HasValue)
                collection.Add(request.NewName, caption, request.InsertIndex.Value);
            else collection.Add(request.NewName, caption);
            string newPath = request.ParentPath + "/" + collectionName + "/" + request.NewName;
            dynamic after = Tree(request.Project, request.Form);
            if (!TreeContainsPath((IEnumerable)after.Controls, newPath) ||
                string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The added Page or Tab was not reflected in the UserForm tree.");
            return new { AddedPath = newPath, Applied = true, Tree = after };
        }

        public object RemovePageOrTab(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 4 || parts.Length % 2 != 0 ||
                (parts[parts.Length - 2] != "Pages" && parts[parts.Length - 2] != "Tabs"))
                throw new ArgumentException("ControlPath must identify a Page or Tab.");
            string collectionName = parts[parts.Length - 2];
            string name = parts[parts.Length - 1];
            string parentPath = string.Join("/", parts.Take(parts.Length - 2));
            object parent = ResolveTreeItem(form.Designer, parentPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(parent).Find(collectionName, true);
            if (descriptor == null)
                throw new InvalidOperationException("The selected parent has no " + collectionName + " collection.");
            dynamic collection = descriptor.GetValue(parent);
            int index = -1;
            int current = 0;
            foreach (dynamic item in collection)
            {
                if (string.Equals((string)item.Name, name, StringComparison.Ordinal))
                {
                    if (index >= 0) throw new InvalidOperationException("Page or Tab name is ambiguous.");
                    index = current;
                }
                current++;
            }
            if (index < 0) throw new InvalidOperationException("Page or Tab no longer exists in its collection.");
            collection.Remove(index);
            dynamic after = Tree(request.Project, request.Form);
            if (TreeContainsPath((IEnumerable)after.Controls, request.ControlPath) ||
                string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The Page or Tab removal was not reflected in the UserForm tree.");
            return new { RemovedPath = request.ControlPath, Applied = true, Tree = after };
        }

        private static object ConvertDescriptorValue(object value, Type declaredType, object previous)
        {
            bool variant = declaredType != null &&
                declaredType.FullName == "System.Windows.Forms.ComponentModel.Com2Interop.Com2Variant";
            Type type = declaredType == typeof(object) || declaredType == null || variant
                ? previous?.GetType() ?? value.GetType() : declaredType;
            if (type == typeof(Color))
            {
                if (value is string && ((string)value).StartsWith("#", StringComparison.Ordinal))
                    return ColorTranslator.FromHtml((string)value);
                return ColorTranslator.FromOle(Convert.ToInt32(value, CultureInfo.InvariantCulture));
            }
            return ConvertScalar(value, type);
        }

        private static bool SameDescriptorValue(object actual, object expected)
        {
            if (Equals(actual, expected)) return true;
            if (actual == null || expected == null) return false;
            if (actual is Color && expected is Color)
                return ((Color)actual).ToArgb() == ((Color)expected).ToArgb();
            if (actual is IConvertible && expected is IConvertible)
                return string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture),
                    Convert.ToString(expected, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase);
            return false;
        }

        private static bool TreeContainsPath(IEnumerable nodes, string path)
        {
            foreach (dynamic node in nodes)
            {
                if (string.Equals((string)node.Path, path, StringComparison.Ordinal)) return true;
                if (TreeContainsPath((IEnumerable)node.Children, path)) return true;
            }
            return false;
        }

        private static object ResolveTreeItem(object designer, string path)
        {
            string[] parts = path.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 || parts.Length > 16)
                throw new ArgumentException("Invalid control hierarchy path.");
            object current = designer;
            for (int i = 0; i < parts.Length; i += 2)
            {
                PropertyDescriptor collection = TypeDescriptor.GetProperties(current).Find(parts[i], true);
                if (collection == null || (parts[i] != "Controls" && parts[i] != "Pages" && parts[i] != "Tabs"))
                    throw new InvalidOperationException("Path collection is unavailable: " + parts[i]);
                object found = null;
                foreach (dynamic candidate in (dynamic)collection.GetValue(current))
                    if (string.Equals((string)candidate.Name, parts[i + 1], StringComparison.Ordinal))
                    { found = candidate; break; }
                if (found == null) throw new InvalidOperationException("Path item is unavailable: " + parts[i + 1]);
                current = found;
            }
            return current;
        }

        private static dynamic ResolveNestedControls(dynamic designer, string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("ParentPath is required.");
            string[] parts = path.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 || parts[0] != "Controls" || parts.Length > 12)
                throw new ArgumentException("ParentPath must start with Controls/<name> and use Controls or Pages segments.");
            object current = designer;
            for (int i = 0; i < parts.Length; i += 2)
            {
                string collectionName = parts[i];
                string itemName = parts[i + 1];
                if ((collectionName != "Controls" && collectionName != "Pages") ||
                    !Regex.IsMatch(itemName, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                    throw new ArgumentException("ParentPath contains an invalid collection or name.");
                PropertyDescriptor descriptor = TypeDescriptor.GetProperties(current).Find(collectionName, true);
                if (descriptor == null)
                    throw new InvalidOperationException("The path element has no " + collectionName + " collection.");
                object collection = descriptor.GetValue(current);
                object match = null;
                foreach (dynamic candidate in (dynamic)collection)
                    if (string.Equals((string)candidate.Name, itemName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (match != null) throw new InvalidOperationException("Ambiguous parent path element: " + itemName);
                        match = candidate;
                    }
                if (match == null) throw new InvalidOperationException("Parent path element not found: " + itemName);
                current = match;
            }
            PropertyDescriptor controls = TypeDescriptor.GetProperties(current).Find("Controls", true);
            if (controls == null) throw new InvalidOperationException("The selected parent has no Controls collection.");
            return controls.GetValue(current);
        }

    }
}
