using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

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
            var result = new List<object>();
            foreach (dynamic property in form.Properties)
            {
                object value = null;
                try { value = property.Value; } catch { }
                result.Add(new { Name = (string)property.Name,
                    Value = value == null || System.Runtime.InteropServices.Marshal.IsComObject(value) ? null : value.ToString() });
            }
            return result;
        }

        public object Create(Request request)
        {
            ValidateName(request.Form, "Form");
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
    }
}
