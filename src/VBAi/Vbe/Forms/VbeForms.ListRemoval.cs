using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace VBAi
{
    /// <summary>Retire un élément d’une liste MSForms sous contrôle de sa version et vérifie le résultat.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Retire une ligne d’une liste à une colonne non liée sous contrôle des versions de l’arbre et de la liste.</summary>
        /// <param name="request">Projet, formulaire, chemin, versions attendues et index de ligne.</param>
        /// <returns>Rapport comprenant la valeur retirée et le statut de vérification.</returns>
        /// <exception cref="ArgumentException">Une donnée obligatoire ou RowIndex manque.</exception>
        /// <exception cref="ArgumentOutOfRangeException">RowIndex est négatif ou hors des lignes actuelles.</exception>
        /// <exception cref="InvalidOperationException">La liste est liée, multicolonne, illisible ou a changé depuis sa lecture.</exception>
        public object RemoveListItem(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) ||
                string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) ||
                string.IsNullOrWhiteSpace(request.ExpectedListVersion) ||
                !request.RowIndex.HasValue)
                throw new ArgumentException("Project, Form, ControlPath, ExpectedTreeVersion, ExpectedListVersion and RowIndex are required.");
            if (request.RowIndex.Value < 0)
                throw new ArgumentOutOfRangeException("RowIndex must be nonnegative.");

            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string type = TypeDescriptor.GetClassName(target);
            if (!string.Equals(type, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(type, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ControlPath must identify a ComboBox or ListBox.");
            dynamic list = target;
            if (!string.IsNullOrWhiteSpace(Convert.ToString(list.RowSource, CultureInfo.InvariantCulture)))
                throw new InvalidOperationException("RemoveItem is unavailable while RowSource binds the list.");
            if (Convert.ToInt32(list.ColumnCount, CultureInfo.InvariantCulture) != 1)
                throw new InvalidOperationException("RemoveItem is restricted to one-column lists after an Excel crash in a multicolumn sequence.");

            dynamic before = ListItems(new Request { Project = request.Project,
                Form = request.Form, ControlPath = request.ControlPath, Offset = 0, Limit = 64 });
            int count = (int)before.TotalRows;
            if (count > 64 || before.ListVersion == null)
                throw new InvalidOperationException("RemoveItem requires at most 64 readable one-column items; read the full list first.");
            if (request.RowIndex.Value >= count)
                throw new ArgumentOutOfRangeException("RowIndex is outside the current list.");
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase) ||
                !string.Equals((string)before.ListVersion, request.ExpectedListVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form or list changed since it was read.");
            List<object> expected = OneColumnValues((IEnumerable)before.Rows);
            object removed = expected[request.RowIndex.Value];
            expected.RemoveAt(request.RowIndex.Value);

            try { list.RemoveItem(request.RowIndex.Value); }
            catch (Exception ex)
            {
                return new { ControlPath = request.ControlPath, RowIndex = request.RowIndex.Value,
                    Applied = (bool?)null, Verified = false, VerificationPending = true,
                    NativeError = ex.Message, NextRead = "form_list_items" };
            }
            try
            {
                dynamic after = ListItems(new Request { Project = request.Project,
                    Form = request.Form, ControlPath = request.ControlPath, Offset = 0, Limit = 64 });
                List<object> actual = OneColumnValues((IEnumerable)after.Rows);
                bool verified = after.ListVersion != null && (int)after.TotalRows == count - 1 &&
                    expected.SequenceEqual(actual);
                return new { ControlPath = request.ControlPath, RowIndex = request.RowIndex.Value,
                    RemovedValue = removed, Applied = (bool?)true, Verified = verified,
                    VerificationPending = !verified, NativeError = (string)null,
                    ListVersionBefore = (string)before.ListVersion,
                    ListVersionAfter = (string)after.ListVersion,
                    TreeVersionAfter = (string)after.TreeVersion, NextRead = "form_list_items" };
            }
            catch (Exception ex)
            {
                return new { ControlPath = request.ControlPath, RowIndex = request.RowIndex.Value,
                    Applied = (bool?)true, Verified = false, VerificationPending = true,
                    NativeError = ex.Message, NextRead = "form_list_items" };
            }
        }

        /// <summary>Aplati les valeurs des cellules des lignes d’une liste à une colonne.</summary>
        /// <param name="rows">Lignes dont chaque entrée expose des cellules.</param>
        /// <returns>Valeurs de cellules dans l’ordre des lignes.</returns>
        private static List<object> OneColumnValues(IEnumerable rows)
        {
            var values = new List<object>();
            foreach (dynamic row in rows)
                foreach (dynamic cell in (IEnumerable)row.Cells)
                    values.Add((object)cell.Value);
            return values;
        }
    }
}
