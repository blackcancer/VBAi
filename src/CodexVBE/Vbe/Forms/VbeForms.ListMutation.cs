using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Ajoute un élément à une liste MSForms sous contrôle de sa version et vérifie le résultat.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Ajoute un élément à une liste à une colonne non liée, si ses versions correspondent, puis relit toutes ses valeurs.</summary>
        /// <param name="request">Projet, formulaire, chemin, versions attendues et texte à ajouter.</param>
        /// <returns>Rapport qui distingue application, vérification réussie et relecture en attente.</returns>
        /// <exception cref="ArgumentException">Une donnée obligatoire manque ou le texte dépasse 4096 caractères.</exception>
        /// <exception cref="InvalidOperationException">La liste est liée, multicolonne, trop grande, illisible ou a changé depuis sa lecture.</exception>
        public object AddListItem(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) ||
                string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) ||
                string.IsNullOrWhiteSpace(request.ExpectedListVersion) ||
                request.Text == null)
                throw new ArgumentException("Project, Form, ControlPath, ExpectedTreeVersion, ExpectedListVersion and Text are required.");
            if (request.Text.Length > 4096)
                throw new ArgumentException("The list item text exceeds 4096 characters.");

            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string type = TypeDescriptor.GetClassName(target);
            if (!string.Equals(type, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(type, "ListBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ControlPath must identify a ComboBox or ListBox.");
            dynamic list = target;
            if (!string.IsNullOrWhiteSpace(Convert.ToString(list.RowSource, CultureInfo.InvariantCulture)))
                throw new InvalidOperationException("AddItem is unavailable while RowSource binds the list.");
            if (Convert.ToInt32(list.ColumnCount, CultureInfo.InvariantCulture) != 1)
                throw new InvalidOperationException("AddItem is restricted to one-column lists after an Excel crash in a multicolumn sequence.");

            dynamic before = ListItems(new Request { Project = request.Project,
                Form = request.Form, ControlPath = request.ControlPath, Offset = 0, Limit = 64 });
            int count = (int)before.TotalRows;
            if (count >= 64 || before.ListVersion == null)
                throw new InvalidOperationException("AddItem requires fewer than 64 readable one-column items; read the full list first.");
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase) ||
                !string.Equals((string)before.ListVersion, request.ExpectedListVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form or list changed since it was read.");
            List<object> expected = OneColumnValues((IEnumerable)before.Rows);
            expected.Add(request.Text);

            try { list.AddItem(request.Text); }
            catch (Exception ex)
            {
                return new { ControlPath = request.ControlPath, Applied = (bool?)null,
                    Verified = false, VerificationPending = true,
                    NativeError = ex.Message, NextRead = "form_list_items" };
            }
            try
            {
                dynamic after = ListItems(new Request { Project = request.Project,
                    Form = request.Form, ControlPath = request.ControlPath, Offset = 0, Limit = 64 });
                List<object> actual = OneColumnValues((IEnumerable)after.Rows);
                bool verified = after.ListVersion != null && (int)after.TotalRows == count + 1 &&
                    expected.SequenceEqual(actual);
                return new { ControlPath = request.ControlPath, AddedValue = request.Text,
                    Applied = (bool?)true, Verified = verified, VerificationPending = !verified,
                    NativeError = (string)null, ListVersionBefore = (string)before.ListVersion,
                    ListVersionAfter = (string)after.ListVersion,
                    TreeVersionAfter = (string)after.TreeVersion, NextRead = "form_list_items" };
            }
            catch (Exception ex)
            {
                return new { ControlPath = request.ControlPath, Applied = (bool?)true,
                    Verified = false, VerificationPending = true,
                    NativeError = ex.Message, NextRead = "form_list_items" };
            }
        }

        // Kept bridge-only while persistence and teardown are qualified in a
        // disposable host. Never apply this to a bound list.
        /// <summary>Ajoute en place un élément à une liste à une colonne non liée et vérifie le nouveau compte de lignes.</summary>
        /// <param name="request">Projet, formulaire, chemin, version attendue et texte à ajouter.</param>
        /// <returns>Résultat avec les comptes avant/après et indication de vérification.</returns>
        /// <exception cref="ArgumentException">Une donnée obligatoire manque ou le texte dépasse 4096 caractères.</exception>
        /// <exception cref="InvalidOperationException">La liste est liée, multicolonne, hors limites ou a changé depuis sa lecture.</exception>
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
            int columnCount = Convert.ToInt32(list.ColumnCount, CultureInfo.InvariantCulture);
            if (columnCount != 1)
                throw new InvalidOperationException("The AddItem probe is restricted to one-column lists after an Excel crash in a multicolumn sequence.");
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
