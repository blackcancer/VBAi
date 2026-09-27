using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // Copy only the visible design shell. Value and GroupName affect radio
        // selection across sibling controls and require a separate group probe.
        public object DuplicateOptionButton(Request request)
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
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 ||
                !string.Equals(parts[parts.Length - 2], "Controls", StringComparison.Ordinal))
                throw new ArgumentException("ControlPath must identify an OptionButton control.");

            object source = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(source), "OptionButton", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only native MSForms OptionButton controls are supported by this probe.");
            object owner = parts.Length == 2 ? (object)form.Designer :
                ResolveTreeItem(form.Designer, string.Join("/", parts.Take(parts.Length - 2)));
            PropertyDescriptor controlsProperty = TypeDescriptor.GetProperties(owner).Find("Controls", true);
            if (controlsProperty == null)
                throw new InvalidOperationException("The parent has no Controls collection.");
            dynamic controls = controlsProperty.GetValue(owner);
            foreach (dynamic existing in controls)
                if (string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with the new name already exists.");

            dynamic original = source;
            string caption = (string)original.Caption;
            double left = Convert.ToDouble(original.Left, CultureInfo.InvariantCulture);
            double top = Convert.ToDouble(original.Top, CultureInfo.InvariantCulture);
            double width = Convert.ToDouble(original.Width, CultureInfo.InvariantCulture);
            double height = Convert.ToDouble(original.Height, CultureInfo.InvariantCulture);
            if (!IsFinite(left) || !IsFinite(top) || !IsFinite(width) || !IsFinite(height) ||
                left < 0 || top < 0 || width <= 0 || height <= 0 ||
                left > 32767 || top > 32767 || width > 32767 || height > 32767)
                throw new InvalidOperationException("The source OptionButton geometry is outside the supported range.");

            string parentPath = string.Join("/", parts.Take(parts.Length - 2));
            string newPath = (parts.Length == 2 ? "Controls" : parentPath + "/Controls") + "/" + request.NewName;
            bool created = false;
            try
            {
                dynamic copy = controls.Add("Forms.OptionButton.1", request.NewName, true);
                created = true;
                copy.Left = left;
                copy.Top = top;
                copy.Width = width;
                copy.Height = height;
                copy.Caption = caption;

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
                    Math.Abs(Convert.ToDouble(installed.Height, CultureInfo.InvariantCulture) - height) > 0.01)
                    throw new InvalidOperationException("The duplicate did not retain the supported OptionButton properties.");
                return new { SourcePath = request.ControlPath, NewPath = newPath,
                    CopiedProperties = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    Completeness = "Partial", SelectionCopied = false, GroupCopied = false, Tree = after };
            }
            catch
            {
                if (created)
                {
                    try { controls.Remove(request.NewName); }
                    catch (Exception rollback)
                    {
                        throw new InvalidOperationException(
                            "OptionButton duplication failed and rollback also failed; inspect form_tree before retrying.", rollback);
                    }
                }
                throw;
            }
        }
    }
}
