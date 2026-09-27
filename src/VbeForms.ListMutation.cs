using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // Kept bridge-only while persistence and teardown are qualified in a
        // disposable host. Never apply this to a bound list.
        public object AppendListItem(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) ||
                string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) ||
                request.Text == null)
                throw new ArgumentException("Project, Form, ControlPath, ExpectedTreeVersion and Text are required.");
            if (request.Text.Length > 4096)
                throw new ArgumentException("The list item text exceeds 4096 characters.");

            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string controlType = TypeDescriptor.GetClassName(target);
            if (!string.Equals(controlType, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(controlType, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ControlPath must identify a ComboBox or ListBox.");

            dynamic list = target;
            string rowSource = Convert.ToString(list.RowSource, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(rowSource))
                throw new InvalidOperationException("AddItem is unavailable while RowSource binds the list.");
            int beforeCount = Convert.ToInt32(list.ListCount, CultureInfo.InvariantCulture);
            if (beforeCount < 0 || beforeCount >= 1024)
                throw new InvalidOperationException("The list item count is outside the bounded design-time probe.");
            list.AddItem(request.Text);
            int afterCount;
            try { afterCount = Convert.ToInt32(list.ListCount, CultureInfo.InvariantCulture); }
            catch (Exception ex)
            {
                return new { ControlPath = request.ControlPath, Applied = true,
                    Verified = false, VerificationPending = true, CountBefore = beforeCount,
                    CountAfter = (int?)null, ReadbackError = ex.Message,
                    NextRead = "form_list_items" };
            }
            return new { ControlPath = request.ControlPath, Applied = true,
                Verified = afterCount == beforeCount + 1,
                VerificationPending = afterCount != beforeCount + 1,
                CountBefore = beforeCount, CountAfter = (int?)afterCount,
                ReadbackError = (string)null, NextRead = "form_list_items" };
        }
    }
}
