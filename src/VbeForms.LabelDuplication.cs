using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // Deliberately limited to setters already exercised on MSForms Label in Excel.
        // COM PropertyDescriptor.IsReadOnly is not proof of a working setter: Cancel
        // reported writable and its setter failed in the first generic-copy probe.
        public object DuplicateLabel(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            ValidateName(request.NewName, "NewName");
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
                throw new ArgumentException("ControlPath must identify a Label control.");

            object source = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(source), "Label", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only native MSForms Label controls are supported by this probe.");
            object owner = parts.Length == 2 ? (object)form.Designer :
                ResolveTreeItem(form.Designer, string.Join("/", parts.Take(parts.Length - 2)));
            PropertyDescriptor controlsProperty = TypeDescriptor.GetProperties(owner).Find("Controls", true);
            if (controlsProperty == null)
                throw new InvalidOperationException("The parent has no Controls collection.");
            dynamic controls = controlsProperty.GetValue(owner);
            foreach (dynamic existing in controls)
                if (string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with the new name already exists.");

            // Read only the positive list. Do not enumerate every descriptor or copy
            // attributes such as Cancel, whose setter failed in a real Excel host.
            dynamic original = source;
            string caption = (string)original.Caption;
            double left = Convert.ToDouble(original.Left, CultureInfo.InvariantCulture);
            double top = Convert.ToDouble(original.Top, CultureInfo.InvariantCulture);
            double width = Convert.ToDouble(original.Width, CultureInfo.InvariantCulture);
            double height = Convert.ToDouble(original.Height, CultureInfo.InvariantCulture);
            object sourceBackColor = original.BackColor;
            int backColor = sourceBackColor is Color
                ? ColorTranslator.ToOle((Color)sourceBackColor)
                : Convert.ToInt32(sourceBackColor, CultureInfo.InvariantCulture);
            dynamic sourceFont = original.Font;
            string fontName = (string)sourceFont.Name;
            double fontSize = Convert.ToDouble(sourceFont.Size, CultureInfo.InvariantCulture);
            bool fontBold = (bool)sourceFont.Bold;
            if (width <= 0 || height <= 0 || fontSize <= 0 || fontSize > 200 ||
                !IsFinite(left) || !IsFinite(top) || !IsFinite(width) || !IsFinite(height) || !IsFinite(fontSize))
                throw new InvalidOperationException("The source Label has invalid geometry or font size.");

            string parentPath = string.Join("/", parts.Take(parts.Length - 2));
            string newPath = (parts.Length == 2 ? "Controls" : parentPath + "/Controls") + "/" + request.NewName;
            bool created = false;
            try
            {
                dynamic copy = controls.Add("Forms.Label.1", request.NewName, true);
                created = true;
                copy.Left = left;
                copy.Top = top;
                copy.Width = width;
                copy.Height = height;
                copy.Caption = caption;
                copy.BackColor = backColor;
                dynamic copyFont = copy.Font;
                copyFont.Name = fontName;
                copyFont.Size = fontSize;
                copyFont.Bold = fontBold;

                dynamic after = Tree(request.Project, request.Form);
                if (!TreeContainsPath((IEnumerable)after.Controls, newPath) ||
                    string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The duplicate was not reflected in form_tree.");
                dynamic installed = ResolveTreeItem(form.Designer, newPath);
                if ((string)installed.Caption != caption ||
                    Math.Abs(Convert.ToDouble(installed.Left, CultureInfo.InvariantCulture) - left) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Top, CultureInfo.InvariantCulture) - top) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Width, CultureInfo.InvariantCulture) - width) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Height, CultureInfo.InvariantCulture) - height) > 0.01 ||
                    (installed.BackColor is Color
                        ? ColorTranslator.ToOle((Color)installed.BackColor)
                        : Convert.ToInt32(installed.BackColor, CultureInfo.InvariantCulture)) != backColor ||
                    (string)installed.Font.Name != fontName ||
                    Math.Abs(Convert.ToDouble(installed.Font.Size, CultureInfo.InvariantCulture) - fontSize) > 0.01 ||
                    (bool)installed.Font.Bold != fontBold)
                    throw new InvalidOperationException("The duplicate did not retain all supported Label properties.");
                return new { SourcePath = request.ControlPath, NewPath = newPath,
                    CopiedProperties = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "BackColor", "Font.Name", "Font.Size", "Font.Bold" },
                    Completeness = "Partial", Tree = after };
            }
            catch
            {
                if (created)
                {
                    try { controls.Remove(request.NewName); }
                    catch (Exception rollback)
                    {
                        throw new InvalidOperationException(
                            "Label duplication failed and rollback also failed; inspect form_tree before retrying.", rollback);
                    }
                }
                throw;
            }
        }
    }
}
