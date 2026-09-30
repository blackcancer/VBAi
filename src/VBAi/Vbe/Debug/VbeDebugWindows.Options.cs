using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Automation;

namespace VBAi
{
    /// <summary>Inspecte et modifie les préférences du dialogue Options natif du VBE.</summary>
    internal static partial class VbeDebugWindows
    {
        /// <summary>Entrée réellement présente dans une liste native; un libellé vide ne désigne aucune couleur connue.</summary>
        internal sealed class OptionsNativeChoice
        {
                        /// <summary>Index Win32 observé, à partir de zéro.</summary>
                        /// <value>Position dans la liste native.</value>
            public int Index { get; set; }
                        /// <summary>Libellé natif, éventuellement vide pour une palette dessinée par VBE.</summary>
                        /// <value>Texte observé, chaîne vide lorsque l’entrée n’a pas de libellé.</value>
            public string Label { get; set; }
                        /// <summary>Valeur à transmettre : libellé exact ou NativeIndex:index pour une entrée sans libellé.</summary>
                        /// <value>Identifiant exact permettant de resélectionner l’entrée observée.</value>
            public string SelectionValue { get; set; }
        }

        /// <summary>Palettes observées pour une catégorie native, y compris leurs états désactivés.</summary>
        internal sealed class OptionsFormatCategory
        {
                        /// <summary>Libellé exact de la catégorie Couleurs du code.</summary>
                        /// <value>Nom natif de la catégorie sélectionnable.</value>
            public string Category { get; set; }
                        /// <summary>Valeurs et catalogues réels de ses trois palettes, sans interprétation RGB.</summary>
                        /// <value>Palettes Premier plan, Arrière-plan et Indicateur observées.</value>
            public IList<OptionsControl> Palettes { get; set; }
        }

        /// <summary>Sonde optionnelle qui permet une garde globale sur toutes les catégories de couleurs.</summary>
        internal interface IFormatCategoriesOptionsProbe
        {
                        /// <summary>Lit toutes les catégories et restaure la sélection initiale avant de retourner.</summary>
                        /// <param name="dialog">Handle du dialogue Options déjà ouvert.</param>
                        /// <param name="tabIndex">Index de l’onglet Couleurs du code.</param>
                        /// <returns>Palettes observées par catégorie, après restauration de la sélection initiale.</returns>
            IList<OptionsFormatCategory> FormatCategories(IntPtr dialog, int tabIndex);
                        /// <summary>Sélectionne une catégorie exacte dans le dialogue déjà ouvert.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="tabIndex">Index de son onglet Couleurs du code.</param>
                        /// <param name="category">Nom exact de catégorie issu de l’inventaire.</param>
            void SelectFormatCategory(IntPtr dialog, int tabIndex, string category);
        }

        /// <summary>Complète la sonde native par l'inspection temporaire des catégories de code.</summary>
        private sealed partial class NativeOptionsProbe : IFormatCategoriesOptionsProbe
        {
                        /// <summary>Énumère chaque palette avec une restauration garantie de la catégorie initiale.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="tabIndex">Index de l’onglet cible.</param>
                        /// <returns>Palettes natives observées pour chaque catégorie disponible.</returns>
            public IList<OptionsFormatCategory> FormatCategories(IntPtr dialog, int tabIndex)
            {
                SelectOptionsTab(dialog, tabIndex);
                var list = FormatCategoryList(dialog);
                if (list == null) return new OptionsFormatCategory[0];
                var palettes = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox))
                    .Cast<AutomationElement>().Where(x => IsColorPalette(x.Current.Name)).ToArray();
                if (palettes.Length != 3) throw new InvalidOperationException("The native code-category palette controls are absent or ambiguous.");
                var identities = palettes.Select(x => Tuple.Create(x, x.Current.Name, new IntPtr(x.Current.NativeWindowHandle))).ToArray();
                string selectedCategory = null;
                return CaptureOptionsFormatCategories(ReadFormatCategoryList(dialog, list),
                    category => { SelectFormatCategoryList(dialog, list, category); selectedCategory = category; },
                    () =>
                    {
                        var values = identities.Select(x => ReadFormatPalette(dialog, x.Item1, x.Item2, x.Item3)).ToArray();
                        if (!Equals(ReadFormatCategoryList(dialog, list).Value, selectedCategory))
                            throw new InvalidOperationException("The native Code Colors category changed during palette inspection.");
                        return values;
                    });
            }

                        /// <summary>Sélectionne un choix de catégorie unique et vérifie sa sélection accessible.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="tabIndex">Index de l’onglet cible.</param>
                        /// <param name="category">Libellé exact de catégorie.</param>
            public void SelectFormatCategory(IntPtr dialog, int tabIndex, string category)
            {
                SelectOptionsTab(dialog, tabIndex);
                var list = FormatCategoryList(dialog);
                if (list == null) throw new InvalidOperationException("The native Code Colors category list is absent.");
                SelectFormatCategoryList(dialog, list, category);
            }

                        /// <summary>Sélectionne seulement l'onglet demandé; un onglet déjà sélectionné n'est pas réactivé.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="index">Index de l’onglet sélectionné.</param>
            private void SelectOptionsTab(IntPtr dialog, int index)
            {
                GuardOptionsOwnedWindow(dialog, dialog, "#32770");
                if (root == null || root.Current.NativeWindowHandle != dialog.ToInt64() || tabItems == null || index < 0 || index >= tabItems.Count ||
                    tabItems[index].Current.ProcessId != System.Diagnostics.Process.GetCurrentProcess().Id ||
                    !tabItems[index].TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern))
                    throw new InvalidOperationException("The exact native Options tab identity is unreadable.");
                var selection = (SelectionItemPattern)pattern;
                if (!selection.Current.IsSelected) { selection.Select(); PauseNative(75); }
                if (!selection.Current.IsSelected) throw new InvalidOperationException("The native Options tab did not retain its selection.");
            }

                        /// <summary>Repère uniquement la liste des catégories sans inventorier les descendants de la police.</summary>
                        /// <param name="dialog">Handle du dialogue propriétaire.</param>
                        /// <returns>Liste native unique, ou nul si aucune liste prise en charge n’est visible.</returns>
            private AutomationElement FormatCategoryList(IntPtr dialog)
            {
                var lists = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List))
                    .Cast<AutomationElement>().Where(x => IsCodeColorList(x.Current.Name) && !x.Current.IsOffscreen).ToArray();
                if (lists.Length > 1) throw new InvalidOperationException("The native Code Colors list is ambiguous.");
                if (lists.Length == 0) return null;
                GuardOptionsOwnedWindow(dialog, new IntPtr(lists[0].Current.NativeWindowHandle), "ListBox");
                return lists[0];
            }

                        /// <summary>Lit le petit catalogue de catégories et sa sélection réelle avec contrôles d'identité.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="list">Liste des catégories obtenue dans ce dialogue.</param>
                        /// <returns>Catalogue borné et catégorie actuellement sélectionnée.</returns>
            private OptionsControl ReadFormatCategoryList(IntPtr dialog, AutomationElement list)
            {
                GuardOptionsOwnedWindow(dialog, new IntPtr(list.Current.NativeWindowHandle), "ListBox");
                if (list.Current.ControlType != ControlType.List || !IsCodeColorList(list.Current.Name) || list.Current.IsOffscreen || !list.Current.IsEnabled ||
                    !list.TryGetCurrentPattern(SelectionPattern.Pattern, out object pattern))
                    throw new InvalidOperationException("The native Code Colors category list is unreadable.");
                var items = list.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                if (items.Count < 1 || items.Count > 32) throw new InvalidOperationException("The native Code Colors category count is invalid.");
                var choices = items.Cast<AutomationElement>().Select(x =>
                {
                    if (x.Current.ProcessId != System.Diagnostics.Process.GetCurrentProcess().Id || !x.Current.IsEnabled ||
                        !x.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object itemPattern))
                        throw new InvalidOperationException("A native Code Colors category is unreadable.");
                    return x.Current.Name;
                }).ToArray();
                var selected = ((SelectionPattern)pattern).Current.GetSelection();
                if (choices.Any(string.IsNullOrWhiteSpace) || choices.Distinct(StringComparer.Ordinal).Count() != choices.Length || selected.Length != 1 ||
                    choices.Count(x => x == selected[0].Current.Name) != 1)
                    throw new InvalidOperationException("The native Code Colors category selection is ambiguous.");
                return new OptionsControl { Name = list.Current.Name, Type = "ControlType.List", Choices = choices, Value = selected[0].Current.Name };
            }

                        /// <summary>Sélectionne et relit uniquement la liste des catégories, avec catalogue exact validé à chaque action.</summary>
                        /// <param name="dialog">Handle du dialogue Options.</param>
                        /// <param name="list">Liste des catégories identifiée.</param>
                        /// <param name="category">Catégorie exacte à sélectionner.</param>
            private void SelectFormatCategoryList(IntPtr dialog, AutomationElement list, string category)
            {
                var before = ReadFormatCategoryList(dialog, list);
                if (before.Choices.Count(x => x == category) != 1) throw new InvalidOperationException("The exact native Code Colors category is absent.");
                var candidates = list.FindAll(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                    new PropertyCondition(AutomationElement.NameProperty, category)));
                if (candidates.Count != 1 || candidates[0].Current.ProcessId != System.Diagnostics.Process.GetCurrentProcess().Id ||
                    !candidates[0].TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern))
                    throw new InvalidOperationException("The exact native Code Colors category is unreadable or ambiguous.");
                IntPtr listHandle = new IntPtr(list.Current.NativeWindowHandle);
                GuardOptionsOwnedWindow(dialog, listHandle, "ListBox");
                IntPtr parent = OptionsComboParent(listHandle);
                GuardOptionsOwnedWindow(dialog, parent, "#32770");
                int identifier = GetDlgCtrlID(listHandle);
                ((SelectionItemPattern)pattern).Select();
                GuardOptionsOwnedWindow(dialog, listHandle, "ListBox");
                GuardOptionsOwnedWindow(dialog, parent, "#32770");
                if (OptionsComboParent(listHandle) != parent || GetDlgCtrlID(listHandle) != identifier || list.Current.NativeWindowHandle != listHandle.ToInt64())
                    throw new InvalidOperationException("The native Code Colors list identity changed during selection.");
                NotifyOptionsListSelection(listHandle, identifier, parent,
                    (window, message, argument, value) => SendMessageInt(window, message, argument, value));
                if (!Equals(before.Value, category)) PauseNative(15);
                var observed = ReadFormatCategoryList(dialog, list);
                if (!Equals(observed.Value, category))
                {
                    GuardOptionsOwnedWindow(dialog, listHandle, "ListBox");
                    int nativeIndex = SendMessageInt(listHandle, 0x188, IntPtr.Zero, IntPtr.Zero).ToInt32(); // LB_GETCURSEL; read only.
                    int requestedIndex = Array.IndexOf(before.Choices.ToArray(), category);
                    throw new InvalidOperationException("The native Code Colors category selection did not retain its value. " +
                        "Requested=" + category + "; Observed=" + observed.Value + "; NativeIndex=" +
                        nativeIndex.ToString(CultureInfo.InvariantCulture) + "; RequestedIndex=" + requestedIndex.ToString(CultureInfo.InvariantCulture) + ".");
                }
            }

                        /// <summary>Relit une seule palette par son identité Win32, y compris si elle est désactivée pour cette catégorie.</summary>
                        /// <param name="dialog">Handle du dialogue Options propriétaire.</param>
                        /// <param name="element">Élément UI Automation de la palette.</param>
                        /// <param name="name">Nom UI Automation attendu.</param>
                        /// <param name="handle">Handle ComboBox exact capturé avant inspection.</param>
                        /// <returns>Valeurs observées de la palette native.</returns>
            private OptionsControl ReadFormatPalette(IntPtr dialog, AutomationElement element, string name, IntPtr handle)
            {
                GuardOptionsOwnedWindow(dialog, handle, "ComboBox");
                if (element.Current.NativeWindowHandle != handle.ToInt64() || element.Current.Name != name || element.Current.ControlType != ControlType.ComboBox ||
                    element.Current.IsOffscreen || element.Current.IsPassword || !IsColorPalette(name))
                    throw new InvalidOperationException("The exact native code-category palette identity changed.");
                var control = new OptionsControl { Name = name, Type = "ControlType.ComboBox", Visible = true, Enabled = element.Current.IsEnabled };
                ReadOptionsCombo(handle, control);
                return control;
            }
        }

                /// <summary>Émet LBN_SELCHANGE vers le parent déjà qualifié; la sélection UIA seule ne garantit pas sa notification.</summary>
                /// <param name="list">Handle de la liste native.</param>
                /// <param name="identifier">Identifiant de contrôle dans le dialogue.</param>
                /// <param name="parent">Handle du dialogue parent vérifié.</param>
                /// <param name="send">Frontière d’émission de message native.</param>
        internal static void NotifyOptionsListSelection(IntPtr list, int identifier, IntPtr parent, Action<IntPtr, int, IntPtr, IntPtr> send)
        {
            if (list == IntPtr.Zero || parent == IntPtr.Zero || list == parent || identifier < 0 || identifier > ushort.MaxValue || send == null)
                throw new InvalidOperationException("The native Code Colors selection notification identity is invalid.");
            send(parent, 0x111, new IntPtr(identifier | (1 << 16)), list);
        }

                /// <summary>Capture seulement les palettes de chaque catégorie et restaure la sélection même en cas de lecture échouée.</summary>
                /// <param name="list">Catalogue validé de catégories et sélection initiale.</param>
                /// <param name="select">Action de sélection d’une catégorie exacte.</param>
                /// <param name="readPalettes">Lecture des palettes de la catégorie sélectionnée.</param>
                /// <returns>Palettes capturées pour chaque catégorie après restauration de la sélection initiale.</returns>
        internal static IList<OptionsFormatCategory> CaptureOptionsFormatCategories(OptionsControl list, Action<string> select, Func<IList<OptionsControl>> readPalettes)
        {
            if (list == null || !string.IsNullOrEmpty(list.Error) || !(list.Value is string original) || list.Choices == null ||
                list.Choices.Count < 1 || list.Choices.Count > 32 || list.Choices.Any(string.IsNullOrWhiteSpace) ||
                list.Choices.Distinct(StringComparer.Ordinal).Count() != list.Choices.Count || list.Choices.Count(x => x == original) != 1)
                throw new InvalidOperationException("The native Code Colors category catalogue is unreadable or ambiguous.");
            var categories = new List<OptionsFormatCategory>();
            Exception inspectionFailure = null;
            try
            {
                foreach (string category in list.Choices.ToArray())
                {
                    select(category);
                    var palettes = readPalettes();
                    if (palettes == null || palettes.Count != 3 || palettes.Any(x => x == null || !x.Visible || x.Type != "ControlType.ComboBox" ||
                        !IsColorPalette(x.Name) || !string.IsNullOrEmpty(x.Error) || x.Value == null) ||
                        palettes.Select(x => x.Name).Distinct(StringComparer.Ordinal).Count() != 3)
                        throw new InvalidOperationException("A native code-category palette is absent, ambiguous or unreadable.");
                    categories.Add(new OptionsFormatCategory { Category = category, Palettes = palettes });
                }
            }
            catch (Exception error) { inspectionFailure = error; throw; }
            finally
            {
                try { select(original); }
                catch (Exception restorationFailure) when (inspectionFailure != null)
                {
                    // Preserve both causes and make the failed selection restoration
                    // explicit. Do not repeat it or claim the original state was recovered.
                    throw new AggregateException(UiText.Get("Inspection and restoration of the native Code Colors category both failed. The original category selection is unverified."),
                        inspectionFailure, restorationFailure);
                }
            }
            return categories;
        }

        /// <summary>Teste l’appartenance d’un handle enfant au dialogue natif.</summary>
        /// <param name="parent">Handle du dialogue propriétaire.</param>
        /// <param name="child">Handle du contrôle ciblé.</param>
        /// <returns><see langword="true"/> si Windows le reconnaît comme enfant.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "IsChild")]
        private static extern bool OptionsContainsWindow(IntPtr parent, IntPtr child);
                /// <summary>Vérifie classe, PID courant et appartenance réelle au dialogue avant chaque accès ciblé.</summary>
                /// <param name="dialog">Dialogue Options qui doit contenir la fenêtre.</param>
                /// <param name="window">Handle du contrôle à vérifier.</param>
                /// <param name="kind">Classe Win32 attendue.</param>
        private static void GuardOptionsOwnedWindow(IntPtr dialog, IntPtr window, string kind)
        {
            GetWindowThreadProcessId(dialog, out uint dialogPid); GetWindowThreadProcessId(window, out uint pid);
            uint ownPid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            if (dialog == IntPtr.Zero || window == IntPtr.Zero || dialogPid != ownPid || pid != ownPid || ClassName(window) != kind ||
                (window != dialog && !OptionsContainsWindow(dialog, window)))
                throw new InvalidOperationException("The exact native Options control does not belong to its current process and dialog.");
        }

                /// <summary>Reconnaît uniquement les libellés natifs documentés de la liste de catégories.</summary>
                /// <param name="name">Libellé UI Automation du contrôle.</param>
                /// <returns><see langword="true"/> si le nom correspond à une liste Code Colors prise en charge.</returns>
        private static bool IsCodeColorList(string name) => new[] { "code colors", "couleurs du code", "color text", "texte couleur" }.Contains(NormalizeOptionName(name));
                /// <summary>Reconnaît uniquement les libellés des trois palettes de catégorie.</summary>
                /// <param name="name">Libellé UI Automation du contrôle.</param>
                /// <returns><see langword="true"/> si le nom correspond à une palette de couleur prise en charge.</returns>
        private static bool IsColorPalette(string name) => new[] { "foreground", "premier plan", "background", "arrière-plan", "indicator", "indicateur" }.Contains(NormalizeOptionName(name));
                /// <summary>Normalise un libellé pour reconnaître sa fonction sans modifier le choix exact transmis.</summary>
                /// <param name="name">Libellé d’origine.</param>
                /// <returns>Version minuscule nettoyée pour les comparaisons de noms supportés.</returns>
        private static string NormalizeOptionName(string name) => (name ?? "").Replace("&", "").Trim().TrimEnd(':').Trim().ToLowerInvariant();

                /// <summary>Décrit une liste native bornée sans attribuer un nom/RGB aux entrées non libellées.</summary>
                /// <param name="control">État de contrôle à compléter.</param>
                /// <param name="labels">Libellés lus dans l’ordre natif.</param>
                /// <param name="selectedIndex">Index sélectionné, ou -1 pour une valeur éditable.</param>
                /// <param name="editValue">Texte saisi si aucune entrée n’est sélectionnée.</param>
        internal static void DescribeOptionsNativeChoices(OptionsControl control, IList<string> labels, int selectedIndex, string editValue)
        {
            if (labels == null || labels.Count > 2000 || selectedIndex < -1 || selectedIndex >= labels.Count)
                throw new InvalidOperationException("The native options list count or selection is invalid.");
            var choices = new List<OptionsNativeChoice>();
            for (int i = 0; i < labels.Count; i++)
            {
                if (labels[i] == null || labels[i].Length > 4096)
                    throw new InvalidOperationException("A native options list label is unreadable or oversized.");
                choices.Add(new OptionsNativeChoice { Index = i, Label = labels[i],
                    SelectionValue = string.IsNullOrWhiteSpace(labels[i]) ? "NativeIndex:" + i.ToString(CultureInfo.InvariantCulture) : labels[i] });
            }
            control.NativeChoices = choices;
            control.SelectedIndex = selectedIndex;
            control.Choices = choices.Select(x => x.SelectionValue).ToArray();
            control.Value = selectedIndex >= 0 ? choices[selectedIndex].SelectionValue : editValue;
            if (control.Value == null) throw new InvalidOperationException("The native options list selection and edit value cannot be read.");
        }

        /// <summary>Obtient le parent natif d’une ComboBox d’options.</summary>
        /// <param name="window">Handle de la ComboBox.</param>
        /// <returns>Handle du parent Win32.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetParent")]
        private static extern IntPtr OptionsComboParent(IntPtr window);
        /// <summary>Lit un style Win32 de la ComboBox.</summary>
        /// <param name="window">Handle de la ComboBox.</param>
        /// <param name="index">Index de style transmis à GetWindowLongW.</param>
        /// <returns>Valeur du style natif demandé.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int OptionsComboStyle(IntPtr window, int index);
        /// <summary>Lit le texte natif d’une entrée ou de la partie éditable d’une ComboBox.</summary>
        /// <param name="window">Handle de la ComboBox.</param>
        /// <param name="message">Message CB_GETLBTEXT ou WM_GETTEXT à envoyer.</param>
        /// <param name="index">Index de liste ou capacité de texte selon le message.</param>
        /// <param name="text">Tampon recevant le texte natif.</param>
        /// <returns>Nombre de caractères retournés par Windows.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr OptionsComboReadText(IntPtr window, int message, IntPtr index, StringBuilder text);

                /// <summary>Refuse toute utilisation de messages ComboBox sur un autre processus ou une autre classe.</summary>
                /// <param name="window">Handle de la ComboBox à vérifier.</param>
        private static void GuardOptionsCombo(IntPtr window)
        {
            GetWindowThreadProcessId(window, out uint pid);
            if (window == IntPtr.Zero || pid != System.Diagnostics.Process.GetCurrentProcess().Id || ClassName(window) != "ComboBox")
                throw new InvalidOperationException("The native options ComboBox does not belong to this process.");
            if ((OptionsComboStyle(window, -16) & 0x200) == 0)
                throw new InvalidOperationException("An owner-data options ComboBox without native strings cannot be inspected.");
        }

                /// <summary>Lit le catalogue Win32 complet et le texte éditable, puis referme toute liste dépliée pour lecture.</summary>
                /// <param name="window">Handle de la ComboBox qualifiée.</param>
                /// <param name="control">État de sortie qui recevra choix, sélection et valeur éditable.</param>
        private static void ReadOptionsCombo(IntPtr window, OptionsControl control)
        {
            GuardOptionsCombo(window);
            int count = SendMessageInt(window, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32();
            bool expanded = false;
            try
            {
                if (count == 0 && SendMessageInt(window, 0x157, IntPtr.Zero, IntPtr.Zero) == IntPtr.Zero)
                {
                    SendMessageInt(window, 0x14F, new IntPtr(1), IntPtr.Zero); expanded = true;
                    count = SendMessageInt(window, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32();
                }
                if (count < 0 || count > 2000) throw new InvalidOperationException("The native options list count is invalid.");
                var labels = new List<string>();
                for (int i = 0; i < count; i++)
                {
                    int length = SendMessageInt(window, 0x149, new IntPtr(i), IntPtr.Zero).ToInt32();
                    if (length < 0 || length > 4096) throw new InvalidOperationException("A native options list label is unreadable or oversized.");
                    var text = new StringBuilder(length + 1);
                    if (OptionsComboReadText(window, 0x148, new IntPtr(i), text).ToInt32() != length)
                        throw new InvalidOperationException("A native options list label changed during inspection.");
                    labels.Add(text.ToString());
                }
                int selected = SendMessageInt(window, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();
                int valueLength = SendMessageInt(window, 0x000E, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (valueLength < 0 || valueLength > 4096) throw new InvalidOperationException("The native options edit value is oversized.");
                var value = new StringBuilder(valueLength + 1);
                if (OptionsComboReadText(window, 0x000D, new IntPtr(value.Capacity), value).ToInt32() != valueLength)
                    throw new InvalidOperationException("The native options edit value changed during inspection.");
                DescribeOptionsNativeChoices(control, labels, selected, value.ToString());
            }
            finally { if (expanded) SendMessageInt(window, 0x14F, IntPtr.Zero, IntPtr.Zero); }
        }

                /// <summary>Sélectionne une entrée exacte et notifie le parent comme une sélection native, sans frappe ni coordonnées.</summary>
                /// <param name="window">Handle de la ComboBox qualifiée.</param>
                /// <param name="choice">Valeur exacte d’un choix observé.</param>
        private static void WriteOptionsCombo(IntPtr window, string choice)
        {
            var observed = new OptionsControl(); ReadOptionsCombo(window, observed);
            var matches = observed.NativeChoices.Where(x => x.SelectionValue == choice).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The exact native option choice is absent or ambiguous.");
            int index = matches[0].Index;
            if (SendMessageInt(window, 0x14E, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
                throw new InvalidOperationException("The native options ComboBox refused the selected index.");
            IntPtr parent = OptionsComboParent(window);
            GetWindowThreadProcessId(parent, out uint pid);
            if (parent == IntPtr.Zero || pid != System.Diagnostics.Process.GetCurrentProcess().Id)
                throw new InvalidOperationException("The native options ComboBox parent does not belong to this process.");
            int id = GetDlgCtrlID(window);
            SendMessageInt(parent, 0x111, new IntPtr((id & 0xffff) | (1 << 16)), window);
            SendMessageInt(parent, 0x111, new IntPtr((id & 0xffff) | (9 << 16)), window);
        }
        /// <summary>Lit l’état actif d’une fenêtre Win32 du dialogue Options.</summary>
        /// <param name="window">Handle de la fenêtre.</param>
        /// <returns><see langword="true"/> si Windows indique que la fenêtre est activée.</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint="IsWindowEnabled")]
        private static extern bool NativeOptionsWindowEnabled(IntPtr window);
        /// <summary>Native enabled-state boundary, preserving the Options button validation.</summary>
        internal static Func<IntPtr, bool> OptionsWindowEnabled = NativeOptionsWindowEnabled;
        /// <summary>Sonde d’options limitée aux contrôles accessibles du dialogue courant.</summary>
        internal interface IWritableOptionsProbe : IOptionsProbe
        {
            /// <summary>Écrit une valeur dans un contrôle natif déjà identifié.</summary>
            /// <param name="dialog">Dialogue Options propriétaire.</param>
            /// <param name="tabIndex">Index de l’onglet qui contient le contrôle.</param>
            /// <param name="name">Nom exact du contrôle.</param>
            /// <param name="type">Type Automation du contrôle.</param>
            /// <param name="value">Valeur préalablement validée pour cette option.</param>
            void Write(IntPtr dialog, int tabIndex, string name, string type, object value);
            /// <summary>Valide le dialogue Options par son bouton OK natif.</summary>
            /// <param name="dialog">Handle du dialogue à confirmer.</param>
            void Accept(IntPtr dialog);
        }
                /// <summary>Écrit une préférence reconnue d’édition/débogage, puis ferme par validation native.</summary>
                /// <param name="request">Onglet, propriété, valeur et version attendue des options.</param>
                /// <returns>Valeurs avant/après et indication de validation/fermeture du dialogue.</returns>
        public static object SetVbeOption(Request request) => SetVbeOption(request, new NativeOptionsProbe());
                /// <summary>Orchestration injectable, sans modification des préférences tant que la version ne correspond pas.</summary>
                /// <param name="request">Onglet, propriété, valeur et version attendue des options.</param>
                /// <param name="native">Sonde utilisée pour lire, écrire, valider ou fermer le dialogue natif.</param>
                /// <returns>Résultat de l’écriture et de la relecture du contrôle avant validation.</returns>
        internal static object SetVbeOption(Request request, IWritableOptionsProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Pane) || string.IsNullOrWhiteSpace(request.Property) || string.IsNullOrWhiteSpace(request.ExpectedOptionsVersion))
                throw new ArgumentException("Pane, Property, Value and ExpectedOptionsVersion from read_vbe_options are required.");
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++) { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The native VBE Options dialog did not open.");
            bool commitRequested = false;
            try
            {
                var before = CaptureOptionsTabs(native, dialog);
                if (!string.Equals(OptionsRevision(before), request.ExpectedOptionsVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBE options changed since inspection; read them again.");
                var names = native.Tabs(dialog);
                int index = -1;
                for (int i = 0; i < names.Count; i++) if (names[i] == request.Pane) { if (index >= 0) throw new InvalidOperationException("The options tab is ambiguous."); index = i; }
                if (index < 0) throw new InvalidOperationException("The exact options tab is absent.");
                if (!string.IsNullOrEmpty(request.Query))
                {
                    if (!IsColorPalette(request.Property) || !(native is IFormatCategoriesOptionsProbe format))
                        throw new InvalidOperationException("Query is supported only for an exact native Code Colors palette category.");
                    format.SelectFormatCategory(dialog, index, request.Query);
                }
                var matches = native.Controls(dialog, index).Where(x => x.Visible && x.Enabled && x.Name == request.Property &&
                    (x.Type == "ControlType.CheckBox" || x.Type == "ControlType.RadioButton" || x.Type == "ControlType.Edit" || x.Type == "ControlType.ComboBox" || x.Type == "ControlType.List")).ToArray();
                if (matches.Length != 1 || !string.IsNullOrEmpty(matches[0].Error)) throw new InvalidOperationException("The exact option is absent, ambiguous or unreadable.");
                var selected = matches[0];
                object writeValue = ValidateEditableOption(request.Pane, selected, request.Value);
                object oldValue = selected.Value;
                native.Write(dialog, index, selected.Name, selected.Type, writeValue);
                var readback = native.Controls(dialog, index);
                if (!string.IsNullOrEmpty(request.Query) && readback.Count(x => x.Type == "ControlType.List" && IsCodeColorList(x.Name) && Equals(x.Value, request.Query)) != 1)
                    throw new InvalidOperationException("The native Code Colors category changed during mutation; Cancel will be requested.");
                var observed = readback.Where(x => x.Visible && x.Enabled && x.Name == selected.Name && x.Type == selected.Type).ToArray();
                object expected = selected.Type == "ControlType.CheckBox" ? (object)((bool)writeValue ? "On" : "Off") : selected.Type == "ControlType.RadioButton" ? (object)true : Convert.ToString(writeValue, CultureInfo.InvariantCulture);
                if (observed.Length != 1 || !string.IsNullOrEmpty(observed[0].Error) || !Equals(observed[0].Value, expected))
                    throw new InvalidOperationException("The native option did not retain the requested value; Cancel will be requested.");
                commitRequested = true;
                native.Accept(dialog);
                bool closed = false;
                for (int attempt = 0; attempt < 40; attempt++) { if (native.Dialog() == IntPtr.Zero) { closed = true; break; } native.Pause(50); }
                return new { request.Pane, request.Property, Category = request.Query, Before = oldValue, After = observed[0].Value,
                    CommitRequested = true, DialogClosed = closed, ControlValueVerified = true,
                    PersistenceVerified = false, NextRead = "read_vbe_options",
                    Limit = "Reopen Options to verify committed preferences. A requested OK or a closed dialog alone is not a restart persistence proof. Do not retry automatically." };
            }
            finally
            {
                if (!commitRequested || native.Dialog() != IntPtr.Zero) native.Close(dialog);
            }
        }
                /// <summary>Valide une préférence reconnue; les listes exigent un choix natif exact et unique.</summary>
                /// <param name="tab">Nom natif de l’onglet.</param>
                /// <param name="control">État lu du contrôle cible.</param>
                /// <param name="value">Valeur proposée par l’appelant.</param>
                /// <returns>Valeur normalisée acceptée pour le contrôle.</returns>
        internal static object ValidateEditableOption(string tab, OptionsControl control, object value)
        {
            string normalized = (control.Name ?? "").Replace("&", "").Trim().TrimEnd(':').Trim().ToLowerInvariant();
            string tabName = (tab ?? "").Replace("&", "").Trim().ToLowerInvariant();
            bool editor = tabName == "editor" || tabName == "éditeur" || tabName == "editeur";
            bool general = tabName == "general" || tabName == "général";
            bool format = new[] { "editor format", "format de l'éditeur", "format de l’éditeur", "format de l'editeur" }.Contains(tabName);
            bool docking = tabName == "docking" || tabName == "ancrage";
            if (!editor && !general && !format && !docking) throw new InvalidOperationException("The native options tab is not supported.");
            var editorChecks = new[] { "auto syntax check", "require variable declaration", "auto list members", "auto quick info", "auto data tips", "auto indent",
                "drag-and-drop text editing", "default to full module view", "procedure separator",
                "vérification automatique de la syntaxe", "déclaration des variables obligatoire", "exiger une déclaration de variable", "liste des membres automatique", "liste automatique des membres", "info rapide automatique", "informations rapides automatiques", "complément automatique des instructions", "info express automatique", "info-bulles automatiques", "conseils sur les données automatiques", "retrait automatique",
                "modification de texte par glisser-déplacer", "édition de texte par glisser-déplacer", "affichage module complet par défaut", "affichage du module complet par défaut", "séparateur de procédure", "séparation des procédures" };
            var generalChecks = new[] { "compile on demand", "background compile", "show grid", "align controls to grid", "notify before state loss", "show tooltips", "collapse proj. hides windows",
                "compilation à la demande", "compilation sur demande", "compiler à la demande", "compilation en arrière-plan", "compiler en arrière-plan", "afficher la grille", "aligner les contrôles sur la grille", "avertir avant la perte d'état", "avertir avant la perte d’état", "notifier avant la perte d’état", "notifier avant la perte d'état", "afficher les info-bulles", "réduire le proj. masque les fenêtres", "réduire le projet masque les fenêtres" };
            var formatChecks = new[] { "margin indicator bar", "barre des indicateurs en marge", "barre d'indicateurs en marge", "barre d’indicateurs en marge" };
            var dockChecks = new[] { "immediate window", "locals window", "watch window", "project explorer", "properties window", "object browser", "toolbox",
                "fenêtre exécution", "fenêtre variables locales", "fenêtre espions", "explorateur de projets", "explorateur de projet", "fenêtre propriétés", "explorateur d'objets", "explorateur d’objets", "boîte à outils" };
            if (control.Type == "ControlType.CheckBox" && (editor ? editorChecks : general ? generalChecks : format ? formatChecks : dockChecks).Contains(normalized))
            { if (!(value is bool)) throw new ArgumentException("This option requires a boolean Value."); return value; }
            if (general && control.Type == "ControlType.RadioButton" &&
                new[] { "break on all errors", "break in class module", "break on unhandled errors", "arrêt sur toutes les erreurs", "arrêt dans le module de classe", "arrêt sur les erreurs non gérées" }.Contains(normalized))
            { if (!(value is bool) || !(bool)value) throw new ArgumentException("Select an error-trapping radio option with Value=true."); return true; }
            bool width = editor && new[] { "tab width", "largeur de tabulation", "largeur de la tabulation" }.Contains(normalized);
            bool grid = general && new[] { "width", "height", "largeur", "hauteur" }.Contains(normalized);
            if (control.Type == "ControlType.Edit" && (width || grid))
            {
                string text = Convert.ToString(value, CultureInfo.InvariantCulture);
                int minimum = grid ? 2 : 1, maximum = grid ? 60 : 32;
                if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int size) || size < minimum || size > maximum)
                    throw new ArgumentException("The option must be an integer between " + minimum + " and " + maximum + ".");
                return size.ToString(CultureInfo.InvariantCulture);
            }
            bool formatChoice = format && ((control.Type == "ControlType.ComboBox" &&
                new[] { "font", "police", "size", "taille", "foreground", "premier plan", "background", "arrière-plan", "indicator", "indicateur" }.Contains(normalized)) ||
                (control.Type == "ControlType.List" && new[] { "code colors", "couleurs du code", "color text", "texte couleur" }.Contains(normalized)));
            if (formatChoice)
            {
                if (!(value is string choice) || string.IsNullOrWhiteSpace(choice)) throw new ArgumentException("Select an exact string from the observed native Choices.");
                if (control.Choices == null || control.Choices.Count(x => string.Equals(x, choice, StringComparison.Ordinal)) != 1)
                    throw new InvalidOperationException("The exact native choice is absent, ambiguous or unreadable.");
                return choice;
            }
            throw new InvalidOperationException("This native option is not in the supported preference list.");
        }
                /// <summary>Capture bornée commune à la lecture et au contrôle de version avant écriture.</summary>
                /// <param name="native">Sonde qui lit les onglets et contrôles visibles.</param>
                /// <param name="dialog">Handle du dialogue Options ouvert.</param>
                /// <returns>Contenu des onglets et contrôles pertinents dans des limites bornées.</returns>
        private static List<object> CaptureOptionsTabs(IOptionsProbe native, IntPtr dialog)
        {
            var tabs = new List<object>(); var names = native.Tabs(dialog);
            if (names.Count < 1 || names.Count > 8) throw new InvalidOperationException("Unexpected native VBE Options tab count: " + names.Count + ".");
            for (int i = 0; i < names.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(names[i])) throw new InvalidOperationException("A native VBE Options tab is unreadable.");
                var observed = native.Controls(dialog, i);
                if (observed.Count > 2000) throw new InvalidOperationException("The Options dialog has too many controls to inspect safely: " + observed.Count + ".");
                var controls = observed.Where(x => x.Visible && x.Enabled && (!string.IsNullOrWhiteSpace(x.Name) || x.Type != "ControlType.Text"))
                    .Select(x => (object)new { x.Name, x.Type, x.Value, x.Error, x.Choices, x.NativeChoices, x.SelectedIndex }).ToList();
                var categories = native is IFormatCategoriesOptionsProbe format ? format.FormatCategories(dialog, i) : new OptionsFormatCategory[0];
                tabs.Add(new { Tab = names[i], Controls = controls, Count = controls.Count, FormatCategories = categories });
            }
            return tabs;
        }
                /// <summary>Version du contenu visible des options, indépendante des handles transitoires.</summary>
                /// <param name="tabs">Instantané des onglets et contrôles capturés.</param>
                /// <returns>Empreinte SHA-256 de l’instantané sérialisé.</returns>
        private static string OptionsRevision(List<object> tabs)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(tabs)))).Replace("-", "").ToLowerInvariant();
        }
    }
}
