using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // MSForms.List(row, column) is indexed and does not appear among the
        // ordinary TypeDescriptor properties returned by form_tree.
        public object ListItems(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) ||
                string.IsNullOrWhiteSpace(request.ControlPath))
                throw new ArgumentException("Project, Form and ControlPath are required.");
            if (request.Offset < 0 || request.Limit < 0)
                throw new ArgumentOutOfRangeException("Offset and Limit must be non-negative.");

            dynamic form = GetForm(GetProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            object control = ResolveTreeItem(form.Designer, request.ControlPath);
            string controlType = TypeDescriptor.GetClassName(control);
            if (!string.Equals(controlType, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(controlType, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ControlPath must identify a ComboBox or ListBox.");

            dynamic list = control;
            int rowCount = Convert.ToInt32(list.ListCount, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(list.ColumnCount, CultureInfo.InvariantCulture);
            if (rowCount < 0 || columnCount < 0)
                throw new InvalidOperationException("MSForms returned a negative list size.");
            // An empty list can report ColumnCount=0; otherwise the default is one column.
            int effectiveColumns = columnCount == 0 ? 1 : columnCount;
            if (effectiveColumns > 32)
                throw new InvalidOperationException("The list has more than 32 columns; this bounded reader cannot enumerate it.");
            int pageSize = request.Limit == 0 ? 20 : Math.Min(request.Limit, 64);
            pageSize = Math.Min(pageSize, Math.Max(1, 128 / effectiveColumns));
            int start = Math.Min(request.Offset, rowCount);
            int end = (int)Math.Min((long)rowCount, (long)start + pageSize);
            var rows = new List<object>();
            var values = new List<object>();
            bool readError = false;
            bool versionable = true;
            for (int row = start; row < end; row++)
            {
                var cells = new List<object>();
                for (int column = 0; column < effectiveColumns; column++)
                {
                    try
                    {
                        object raw = control.GetType().InvokeMember("List", BindingFlags.GetProperty,
                            null, control, new object[] { row, column }, CultureInfo.InvariantCulture);
                        object value = NormalizeScalar(raw);
                        if (raw != null && !(raw is string) && !(raw is IConvertible)) versionable = false;
                        if (value is string && ((string)value).Length > 4096) versionable = false;
                        values.Add(value);
                        cells.Add(new { Column = column, Value = value, Error = (string)null });
                    }
                    catch (Exception ex)
                    {
                        var cause = ex is TargetInvocationException && ex.InnerException != null
                            ? ex.InnerException : ex;
                        readError = true;
                        cells.Add(new { Column = column, Value = (object)null, Error = cause.Message });
                    }
                }
                rows.Add(new { Index = row, Cells = cells });
            }
            string listVersion = null;
            string listVersionError = null;
            if (start != 0 || end != rowCount) listVersionError = "Read the complete list in one page to obtain ListVersion.";
            else if (readError) listVersionError = "One or more indexed list cells could not be read.";
            else if (!versionable) listVersionError = "One or more list cells are non-scalar or exceed 4096 characters.";
            else
            {
                try { listVersion = ComputeListVersion(controlType, request.ControlPath, rowCount, columnCount, values); }
                catch (Exception ex) { listVersionError = "ListVersion could not be calculated: " + ex.Message; }
            }
            return new { Project = request.Project, Form = request.Form,
                ControlPath = request.ControlPath, ControlType = controlType,
                TreeVersion = (string)tree.TreeVersion, TotalRows = rowCount,
                ColumnCount = columnCount, ListVersion = listVersion,
                ListVersionError = listVersionError, Offset = request.Offset,
                ReturnedRows = rows.Count, HasMore = end < rowCount, Rows = rows,
                Scope = "Read-only design-time MSForms.List. TreeVersion may reflect ListCount but does not fingerprint indexed item values." };
        }

        private static string ComputeListVersion(string type, string path, int rowCount,
            int columnCount, List<object> values)
        {
            string json = new JavaScriptSerializer { MaxJsonLength = 1024 * 1024 }.Serialize(new {
                Type = type, ControlPath = path, ListCount = rowCount,
                ColumnCount = columnCount, Values = values });
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(json)))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
