using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{
    /// <summary>Liaison déclarative d’une ComboBox ou ListBox à une plage de feuille Excel.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Insère ou actualise dans le code du UserForm une instruction RowSource gérée pour une plage bornée.</summary>
        /// <param name="request">Classeur, formulaire, chemin du contrôle, feuille/plage et révisions attendues.</param>
        /// <returns>Résultat de la mutation et vérification du bloc VBA généré.</returns>
        /// <exception cref="ArgumentException">Les préconditions ou le chemin de feuille/plage ne sont pas valides.</exception>
        /// <exception cref="InvalidOperationException">Le projet, l’arbre, le contrôle ou le code ne correspond plus à l’instantané.</exception>
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
            string prefix = "' VBAi BEGIN LIST " + request.ControlPath + " SHA256=";
            string end = "' VBAi END LIST " + request.ControlPath;
            string body = BindingLine(accessor, request.SheetName, request.RangeAddress);
            var generated = new[] { "    " + prefix + FormCodeSha(body), body, "    " + end };
            return ApplyManagedListBlock(request, module, before, accessor, prefix, end, generated, 0, columns);
        }

        /// <summary>Valide une référence A1 rectangulaire et renvoie son nombre de colonnes.</summary>
        /// <param name="sheet">Nom littéral d’une feuille Excel.</param>
        /// <param name="address">Adresse d’une cellule ou rectangle A1 borné.</param>
        /// <returns>Largeur de la plage, de 1 à 10 colonnes.</returns>
        /// <exception cref="ArgumentException">La feuille ou la plage est invalide ou dépasse les limites permises.</exception>
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
        /// <summary>Convertit les lettres d’une colonne Excel en indice numérique à partir de un.</summary>
        /// <param name="value">Lettres majuscules ou minuscules de la colonne.</param>
        /// <returns>Indice ordinal de la colonne.</returns>
        private static int ExcelColumn(string value)
        { int result = 0; foreach (char c in value.ToUpperInvariant()) result = result * 26 + c - 'A' + 1; return result; }
        /// <summary>Construit l’instruction VBA RowSource pour l’accès validé au contrôle et à la plage.</summary>
        /// <param name="accessor">Expression générée vers le contrôle du UserForm.</param>
        /// <param name="sheet">Nom de la feuille déjà validé.</param>
        /// <param name="address">Adresse A1 déjà validée.</param>
        /// <returns>Instruction VBA qui lie RowSource à une adresse Excel externe.</returns>
        private static string BindingLine(string accessor, string sheet, string address)
        { return "    Me." + accessor + ".RowSource = ThisWorkbook.Worksheets(\"" + sheet.Replace("\"", "\"\"") + "\").Range(\"" + address.ToUpperInvariant() + "\").Address(External:=True)"; }
        /// <summary>Vérifie qu’une ligne de code correspond à la forme contrôlée d’une liaison RowSource générée.</summary>
        /// <param name="line">Ligne VBA à vérifier.</param>
        /// <param name="accessor">Expression du contrôle cible.</param>
        /// <returns><see langword="true"/> si la ligne respecte exactement le motif attendu.</returns>
        private static bool IsManagedBindingLine(string line, string accessor)
        {
            return Regex.IsMatch(line, "^    Me\\." + Regex.Escape(accessor) + "\\.RowSource = ThisWorkbook\\.Worksheets\\(\"(?:[^\"]|\"\")*\"\\)\\.Range\\(\"[A-Z0-9$:]+\"\\)\\.Address\\(External:=True\\)$");
        }
    }
}
