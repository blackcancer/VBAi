using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

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
                    Type targetType = previous?.GetType() ?? descriptor?.PropertyType;
                    if (targetType == null || targetType == typeof(object))
                        throw new InvalidOperationException("Cannot determine a safe scalar type for " + propertyName);
                    property.Value = ConvertScalar(request.Value, targetType);
                }
                else
                {
                    object source = null;
                    try { source = property.Object; } catch { }
                    if (source == null)
                        throw new InvalidOperationException("The object property is absent or cannot expose members: " + propertyName);
                    PropertyDescriptor member = TypeDescriptor.GetProperties(source).Find(path[1], true);
                    if (member == null) throw new InvalidOperationException("Unknown object member: " + request.Property);
                    if (member.IsReadOnly) throw new InvalidOperationException("This object member is read-only: " + request.Property);
                    object current = member.GetValue(source);
                    object converted = ConvertScalar(request.Value, current?.GetType() ?? member.PropertyType);
                    member.SetValue(source, converted);
                }
            }
            string currentName = (string)form.Name;
            var properties = DescribeProperties(form);
            return new { Project = request.Project, Form = currentName, Property = request.Property,
                Properties = properties, State = Snapshot(request.Project, form) };
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

        private static List<VbePropertyInfo> DescribeProperties(dynamic form)
        {
            var result = new List<VbePropertyInfo>();
            PropertyDescriptorCollection descriptors = TypeDescriptor.GetProperties((object)form.Designer);
            foreach (dynamic property in form.Properties)
            {
                string name = (string)property.Name;
                PropertyDescriptor descriptor = descriptors.Find(name, true);
                var info = new VbePropertyInfo { Name = name, Type = descriptor?.PropertyType?.FullName,
                    ReadOnly = descriptor == null ? (bool?)null : descriptor.IsReadOnly };
                try { info.NumIndices = (int)property.NumIndices; }
                catch (Exception ex) { info.Error = ex.Message; }
                object raw = null;
                try { raw = property.Value; }
                catch (Exception ex) { info.Error = ex.Message; }
                if (raw == null)
                {
                    try { raw = property.Object; } catch { }
                }
                if (raw != null && info.Type == null) info.Type = raw.GetType().FullName;
                if (info.NumIndices > 0) info.Kind = "indexed";
                else if (raw != null && Marshal.IsComObject(raw))
                {
                    info.Kind = "object";
                    info.Members = DescribeObjectMembers(raw);
                }
                else
                {
                    info.Kind = "scalar";
                    info.Value = NormalizeScalar(raw);
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
