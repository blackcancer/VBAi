using System;
using System.Collections.Generic;
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
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class VbeForms
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
            int nodeCount = 0;
            var nodes = ReadChildControls(form.Designer.Controls, "Controls", 0, ref nodeCount);
            var properties = DescribeProperties(form);
            string json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }
                .Serialize(new { Form = (string)form.Name, Properties = properties, Controls = nodes });
            string version;
            using (var sha = SHA256.Create())
                version = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(json)))
                    .Replace("-", "").ToLowerInvariant();
            return new { Project = projectName, Form = formName, Version = version,
                NodeCount = nodeCount, Properties = properties, Controls = nodes };
        }

        private static List<object> ReadChildControls(dynamic collection, string path, int depth, ref int nodeCount)
        {
            var result = new List<object>();
            foreach (dynamic control in collection)
                result.Add(ReadTreeNode(control, "Control", path, depth, ref nodeCount));
            return result;
        }

        private static object ReadTreeNode(object item, string kind, string parentPath, int depth, ref int nodeCount)
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
                        children.AddRange(ReadChildControls(nested, path + "/Controls", depth + 1, ref nodeCount));
                }
                PropertyDescriptor pages = descriptors.Find("Pages", true);
                if (pages != null)
                    foreach (dynamic page in (dynamic)pages.GetValue(item))
                        children.Add(ReadTreeNode(page, "Page", path + "/Pages", depth + 1, ref nodeCount));
                PropertyDescriptor tabs = descriptors.Find("Tabs", true);
                if (tabs != null)
                    foreach (dynamic tab in (dynamic)tabs.GetValue(item))
                        children.Add(ReadTreeNode(tab, "Tab", path + "/Tabs", depth + 1, ref nodeCount));
            }
            else if (kind == "Page")
            {
                PropertyDescriptor controls = descriptors.Find("Controls", true);
                if (controls != null)
                    children.AddRange(ReadChildControls(controls.GetValue(item), path + "/Controls", depth + 1, ref nodeCount));
            }
            return new { Path = path, Name = name, Kind = kind,
                Type = TypeDescriptor.GetClassName(item), Properties = ReadObjectProperties(item), Children = children };
        }

        private static List<VbePropertyInfo> ReadObjectProperties(object item)
        {
            var result = new List<VbePropertyInfo>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(item))
            {
                var info = new VbePropertyInfo { Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName, ReadOnly = descriptor.IsReadOnly };
                try
                {
                    object value = descriptor.GetValue(item);
                    info.Kind = value != null && (Marshal.IsComObject(value) || value is Font || value is Image)
                        ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = NormalizeScalar(value);
                    else if (value is Image) info.Digest = ImageDigest((Image)value);
                    else if (value is Font) info.Members = DescribeObjectMembers(value);
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
                    ReadOnly = descriptor == null ? (bool?)null : descriptor.IsReadOnly };
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
                    ReadOnly = descriptor.IsReadOnly };
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
                    ReadOnly = descriptor.IsReadOnly, Value = value, Error = error });
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
            if (!BuiltInControls.Contains(request.ControlType ?? string.Empty))
                throw new ArgumentException("ControlType must be a built-in Microsoft Forms ProgID.");
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
            dynamic designer = form.Designer;
            var text = new StringBuilder();
            text.Append((string)form.Name).Append('|').Append((string)form.Properties.Item("Caption").Value).Append('|');
            AppendNumber(text, Convert.ToDouble(form.Properties.Item("Width").Value, CultureInfo.InvariantCulture));
            AppendNumber(text, Convert.ToDouble(form.Properties.Item("Height").Value, CultureInfo.InvariantCulture));
            foreach (var property in DescribeProperties((object)form).OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
            {
                AppendVersionText(text, property.Name);
                AppendVersionText(text, property.Kind);
                AppendVersionText(text, Convert.ToString(property.Value, CultureInfo.InvariantCulture));
                AppendVersionText(text, property.Digest);
                if (property.Members != null)
                    foreach (var member in property.Members.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        AppendVersionText(text, member.Name);
                        AppendVersionText(text, Convert.ToString(member.Value, CultureInfo.InvariantCulture));
                    }
            }
            var controls = new List<dynamic>();
            foreach (dynamic control in designer.Controls) controls.Add(control);
            foreach (dynamic control in controls.OrderBy(c => (string)c.Name, StringComparer.OrdinalIgnoreCase))
            {
                text.Append((string)control.Name).Append('|');
                AppendNumber(text, (double)control.Left);
                AppendNumber(text, (double)control.Top);
                AppendNumber(text, (double)control.Width);
                AppendNumber(text, (double)control.Height);
                try { text.Append((string)control.Caption); } catch { }
                text.Append('|');
                try
                {
                    dynamic font = control.Font;
                    text.Append((string)font.Name).Append('|');
                    AppendNumber(text, Convert.ToDouble(font.Size, CultureInfo.InvariantCulture));
                    text.Append((bool)font.Bold).Append('|');
                }
                catch { }
            }
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text.ToString())))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static void AppendNumber(StringBuilder text, double number)
        {
            text.Append(number.ToString("R", CultureInfo.InvariantCulture)).Append('|');
        }

        private static void AppendVersionText(StringBuilder text, string value)
        {
            value = value ?? "";
            text.Append(value.Length).Append(':').Append(value);
        }
    }
}
