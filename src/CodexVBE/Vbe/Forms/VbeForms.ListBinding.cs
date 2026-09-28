using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        public object SetListBinding(Request request)
        {
            string accessor = ListControlAccessor(request.ControlPath);
            int columns = ExcelBindingColumns(request.SheetName, request.RangeAddress);
            if (string.IsNullOrWhiteSpace(request.ExpectedHostPath) || !Path.IsPathRooted(request.ExpectedHostPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) || string.IsNullOrWhiteSpace(request.ExpectedSha256))
                throw new ArgumentException("ExpectedHostPath, ExpectedTreeVersion and ExpectedSha256 are required.");
            dynamic project = GetDesignProject(request.Project);
            string path = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetFullPath(path), Path.GetFullPath(request.ExpectedHostPath), StringComparison.OrdinalIgnoreCase) ||
                !new[] { ".xlsm", ".xlsb", ".xlam", ".xltm", ".xls", ".xla" }.Contains(Path.GetExtension(path).ToLowerInvariant()))
                throw new InvalidOperationException("The project must be the saved Excel workbook identified by ExpectedHostPath.");
            dynamic form = GetForm(project, request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath)) throw new ArgumentException("ControlPath is not canonical in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string type = TypeDescriptor.GetClassName(target);
            if (!string.Equals(type, "ComboBox", StringComparison.OrdinalIgnoreCase) && !string.Equals(type, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The target must be a native ComboBox or ListBox.");
            if (Convert.ToInt32(((dynamic)target).ColumnCount, CultureInfo.InvariantCulture) != columns)
                throw new InvalidOperationException("The range width must match the existing ColumnCount. Designer properties are not changed.");
            dynamic module = form.CodeModule;
            string before = ReadFormCode(module);
            if (!string.Equals(FormCodeSha(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form source changed since it was read.");
            string prefix = "' CodexVBE BEGIN LIST " + request.ControlPath + " SHA256=";
            string end = "' CodexVBE END LIST " + request.ControlPath;
            string body = BindingLine(accessor, request.SheetName, request.RangeAddress);
            var generated = new[] { "    " + prefix + FormCodeSha(body), body, "    " + end };
            return ApplyManagedListBlock(request, module, before, accessor, prefix, end, generated, 0, columns);
        }

        internal static int ExcelBindingColumns(string sheet, string address)
        {
            if (string.IsNullOrWhiteSpace(sheet) || sheet.Length > 31 || sheet.StartsWith("'", StringComparison.Ordinal) || sheet.EndsWith("'", StringComparison.Ordinal) || sheet.Any(char.IsControl) || sheet.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0)
                throw new ArgumentException("SheetName must be a literal Excel worksheet name.");
            Match match = Regex.Match(address ?? "", @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$");
            if (!match.Success) throw new ArgumentException("RangeAddress must be a single bounded A1 cell or rectangle, without workbook, sheet or formula syntax.");
            int left = ExcelColumn(match.Groups[1].Value), top = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            int right = match.Groups[3].Success ? ExcelColumn(match.Groups[3].Value) : left;
            int bottom = match.Groups[4].Success ? int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) : top;
            if (left > 16384 || right > 16384 || top > 1048576 || bottom > 1048576 || right < left || bottom < top || right - left >= 10 || bottom - top >= 10000)
                throw new ArgumentException("The range must be inside Excel limits and contain at most 10 columns and 10000 rows.");
            return right - left + 1;
        }
        private static int ExcelColumn(string value)
        { int result = 0; foreach (char c in value.ToUpperInvariant()) result = result * 26 + c - 'A' + 1; return result; }
        private static string BindingLine(string accessor, string sheet, string address)
        { return "    Me." + accessor + ".RowSource = ThisWorkbook.Worksheets(\"" + sheet.Replace("\"", "\"\"") + "\").Range(\"" + address.ToUpperInvariant() + "\").Address(External:=True)"; }
        private static bool IsManagedBindingLine(string line, string accessor)
        {
            return Regex.IsMatch(line, "^    Me\\." + Regex.Escape(accessor) + "\\.RowSource = ThisWorkbook\\.Worksheets\\(\"(?:[^\"]|\"\")*\"\\)\\.Range\\(\"[A-Z0-9$:]+\"\\)\\.Address\\(External:=True\\)$");
        }
    }
}
