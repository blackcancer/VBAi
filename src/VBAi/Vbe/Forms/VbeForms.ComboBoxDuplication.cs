using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace VBAi
{
    /// <summary>Implémente le profil de duplication partielle des ComboBox natifs.</summary>
    internal sealed partial class VbeForms
    {
        // ListWidth is the one ComboBox-specific setter already read back in
        // disposable Excel. Items, bindings and selection are not copied.
        /// <summary>Crée un ComboBox natif sous le même parent après contrôle de l’arbre et de sa version, puis vérifie sa géométrie et ListWidth.</summary>
        /// <param name="request">Projet, formulaire, chemin du contrôle source, version attendue de l’arbre et nouveau nom.</param>
        /// <returns>Rapport de duplication partielle avec les chemins source et cible ainsi que le nouvel arbre de contrôles.</returns>
        /// <exception cref="ArgumentException">Un champ obligatoire manque, le chemin ne désigne pas un contrôle ou le nouveau nom est invalide.</exception>
        /// <exception cref="InvalidOperationException">L’arbre est périmé, le contrôle ou son type ne convient pas, une propriété source est hors profil, ou la vérification échoue.</exception>
        /// <remarks>Les éléments, liaisons de données et sélection ne sont pas inclus dans le résultat copié. Si la vérification échoue après création, la méthode tente de supprimer le contrôle ajouté.</remarks>
        public object DuplicateComboBox(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            ValidateName(request.NewName, "NewName");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 ||
                !string.Equals(parts[parts.Length - 2], "Controls", StringComparison.Ordinal))
                throw new ArgumentException("ControlPath must identify a ComboBox control.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");

            object source = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(source), "ComboBox", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only native MSForms ComboBox controls are supported by this probe.");
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
            object rawListWidth = original.ListWidth;
            if (!(rawListWidth is string) || string.IsNullOrWhiteSpace((string)rawListWidth) ||
                ((string)rawListWidth).Length > 64)
                throw new InvalidOperationException("This probe only copies a nonempty textual ComboBox.ListWidth.");
            string listWidth = (string)rawListWidth;
            double left = Convert.ToDouble(original.Left, CultureInfo.InvariantCulture);
            double top = Convert.ToDouble(original.Top, CultureInfo.InvariantCulture);
            double width = Convert.ToDouble(original.Width, CultureInfo.InvariantCulture);
            double height = Convert.ToDouble(original.Height, CultureInfo.InvariantCulture);
            if (!IsFinite(left) || !IsFinite(top) || !IsFinite(width) || !IsFinite(height) ||
                left < 0 || top < 0 || width <= 0 || height <= 0 ||
                left > 32767 || top > 32767 || width > 32767 || height > 32767)
                throw new InvalidOperationException("The source ComboBox geometry is outside the supported range.");

            string parentPath = string.Join("/", parts.Take(parts.Length - 2));
            string newPath = (parts.Length == 2 ? "Controls" : parentPath + "/Controls") + "/" + request.NewName;
            bool created = false;
            try
            {
                dynamic copy = controls.Add("Forms.ComboBox.1", request.NewName, true);
                created = true;
                copy.Left = left;
                copy.Top = top;
                copy.Width = width;
                copy.Height = height;
                copy.ListWidth = listWidth;

                dynamic after = Tree(request.Project, request.Form);
                if (!TreeContainsPath((IEnumerable)after.Controls, newPath) ||
                    string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The duplicate was not reflected in form_tree.");
                dynamic installed = ResolveTreeItem(form.Designer, newPath);
                if (!string.Equals(listWidth, installed.ListWidth as string, StringComparison.Ordinal) ||
                    Math.Abs(Convert.ToDouble(installed.Left, CultureInfo.InvariantCulture) - left) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Top, CultureInfo.InvariantCulture) - top) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Width, CultureInfo.InvariantCulture) - width) > 0.01 ||
                    Math.Abs(Convert.ToDouble(installed.Height, CultureInfo.InvariantCulture) - height) > 0.01)
                    throw new InvalidOperationException("The duplicate did not retain the supported ComboBox properties.");
                return new { SourcePath = request.ControlPath, NewPath = newPath,
                    CopiedProperties = new[] { "Name", "Left", "Top", "Width", "Height", "ListWidth" },
                    Completeness = "Partial", ItemsCopied = false, BindingsCopied = false, Tree = after };
            }
            catch
            {
                if (created)
                {
                    try { controls.Remove(request.NewName); }
                    catch (Exception rollback)
                    {
                        throw new InvalidOperationException(
                            "ComboBox duplication failed and rollback also failed; inspect form_tree before retrying.", rollback);
                    }
                }
                throw;
            }
        }
    }
}
