using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Resources;
using System.Windows.Forms;

namespace VBAi
{
    // Explicit culture avoids changing the host's thread culture or VBA number/date formatting.
    /// <summary>Fournit les chaînes localisées sans modifier la culture du fil ou le formatage VBA.</summary>
    internal static class UiText
    {

        /// <summary>Fallback resource manager for invariant English strings embedded in the assembly.</summary>
        private static readonly ResourceManager English = new ResourceManager("VBAi.Localization.UiStrings", typeof(UiText).Assembly);

        /// <summary>Resource managers indexed by the supported culture names.</summary>
        private static readonly Dictionary<string, ResourceManager> Catalogues = CreateCatalogues();

        /// <summary>Creates one localized resource manager for each culture in the supported language catalog.</summary>
        /// <returns>Culture-name to resource-manager mapping used by <see cref="Get"/>.</returns>
        private static Dictionary<string, ResourceManager> CreateCatalogues()
        {
            var result = new Dictionary<string, ResourceManager>();
            foreach (var language in UiLanguages.All)
                result[language.CultureName] = new ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
            return result;
        }

        /// <summary>Culture selected for interface strings after VBE language detection.</summary>
        /// <value>A supported culture; initialization falls back to the current UI culture when VBE menus are unavailable.</value>
        internal static CultureInfo Culture { get; private set; } = Supported(CultureInfo.CurrentUICulture);

        /// <summary>Maps an arbitrary culture to the closest culture supported by the UI catalogs.</summary>
        /// <param name="culture">Requested UI culture.</param>
        /// <returns>The supported culture selected by <see cref="UiLanguages.For(CultureInfo)"/>.</returns>
        internal static CultureInfo Supported(CultureInfo culture)
        {
            return CultureInfo.GetCultureInfo(UiLanguages.For(culture).CultureName);
        }

        /// <summary>Détecte la langue des menus VBE, met à jour la culture active et inscrit le résultat au journal.</summary>
        /// <param name="vbe">VBE Automation object whose menu captions are used when readable.</param>
        internal static void Initialize(object vbe)
        {
            Culture = Detect(vbe, CultureInfo.CurrentUICulture);
            LoadLog.Write("VBAi interface language: " + Culture.Name);
        }

        /// <summary>Privilégie les légendes des menus VBE et utilise la culture système si elles sont indisponibles.</summary>
        /// <param name="vbe">VBE Automation object; inaccessible or absent menus cause the fallback to be used.</param>
        /// <param name="fallback">Culture used when no menu bar can be read or no VBE menu is available.</param>
        /// <returns>Supported culture inferred from VBE captions, or mapped from <paramref name="fallback"/>.</returns>
        internal static CultureInfo Detect(object vbe, CultureInfo fallback)
        {
            bool hasMenu = false;
            var captions = new List<string>();
            // Actual VBE captions take precedence over the operating system's display language.
            try
            {
                foreach (dynamic bar in ((dynamic)vbe).CommandBars)
                {
                    if ((int)bar.Type != 1) continue;
                    hasMenu = true;
                    foreach (dynamic control in bar.Controls)
                    {
                        captions.Add((string)control.Caption);
                    }
                }
            }
            catch (Exception ex) { LoadLog.Write("VBE UI language unavailable: " + ex.Message); }
            return hasMenu ? UiLanguages.FromMenus(captions, fallback) : Supported(fallback);
        }

        /// <summary>Looks up an English resource key in the active catalog and falls back to embedded English.</summary>
        /// <param name="english">English resource key or literal text to localize.</param>
        /// <returns>Localized text, the English value when no resource exists, or null when the input is null.</returns>
        internal static string Get(string english)
        {
            if (english == null) return null;
            if (english == "DOCUMENT CONVERSATIONS") return Get("Document conversations").ToUpperInvariant();
            if (english == "YOU") return Get("You").ToUpperInvariant();
            return Catalogues[Culture.Name].GetString(english, CultureInfo.InvariantCulture)
                ?? English.GetString(english, CultureInfo.InvariantCulture) ?? english;
        }

        // Only called immediately after InitializeComponent, before document/user data is populated.
        // Designer captions stay editable in English; all translated strings live in resource files.
        /// <summary>Localizes initialized controls and applies registered help hints recursively, while preserving technical text direction.</summary>
        /// <param name="control">Initialized root control; Designer mode is left untouched.</param>
        /// <param name="components">Optional component container used to find ToolTip instances.</param>
        /// <param name="additionalTips">Tooltips not stored in <paramref name="components"/> whose text should also be translated.</param>
        internal static void Apply(Control control, IContainer components, params ToolTip[] additionalTips)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            var form = control as Form;
            if (form != null)
            {
                form.RightToLeft = Culture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
                form.RightToLeftLayout = Culture.TextInfo.IsRightToLeft;
            }
            // Technical content must never be mirrored, including URLs, tokens, VBA and diffs.
            if (control is DataGridView || control is ComboBox ||
                Array.IndexOf(new[] { "remote", "branch", "branchName", "branchList", "baseContent", "openAiEndpoint", "ollamaEndpoint", "openAiKey", "manualModels", "resolutionText", "contextPreview", "details" }, control.Name) >= 0)
                control.RightToLeft = RightToLeft.No;
            control.Text = Get(control.Text);
            control.AccessibleName = Get(control.AccessibleName);
            control.AccessibleDescription = Get(control.AccessibleDescription);
            foreach (var tips in additionalTips) tips.SetToolTip(control, Get(tips.GetToolTip(control)));
            if (components != null)
                foreach (IComponent component in components.Components)
                {
                    var tips = component as ToolTip;
                    if (tips != null) tips.SetToolTip(control, Get(tips.GetToolTip(control)));
                }
            UiHelpHints.Apply(control, components, additionalTips);
            var combo = control as ComboBox;
            if (combo != null)
                for (int i = 0; i < combo.Items.Count; i++)
                    if (combo.Items[i] is string) combo.Items[i] = Get((string)combo.Items[i]);
            var grid = control as DataGridView;
            if (grid != null)
                foreach (DataGridViewColumn column in grid.Columns) column.HeaderText = Get(column.HeaderText);
            foreach (Control child in control.Controls) Apply(child, components, additionalTips);
            if (components != null)
                foreach (IComponent component in components.Components)
                    if (component is ContextMenuStrip menu) { menu.ShowItemToolTips = true; ApplyItems(menu.Items); UiTheme.ApplyMenu(menu); }
            if (form != null) UiTheme.Attach(form);
        }

        /// <summary>Translates menu captions, tooltips, accessibility names, and nested drop-down items.</summary>
        /// <param name="items">Menu items to localize in place.</param>
        private static void ApplyItems(ToolStripItemCollection items)
        {
            foreach (ToolStripItem item in items)
            {
                UiHelpHints.Apply(item);
                item.Text = Get(item.Text);
                item.ToolTipText = Get(item.ToolTipText);
                item.AccessibleName = Get(item.AccessibleName);
                var dropdown = item as ToolStripDropDownItem;
                if (dropdown != null) ApplyItems(dropdown.DropDownItems);
            }
        }

        /// <summary>Traduit les étiquettes de rôle affichées dans le transcript et conserve les autres valeurs telles quelles.</summary>
        /// <param name="token">Étiquette française reconnue ou texte à préserver.</param>
        /// <returns>Étiquette localisée ou valeur d’origine.</returns>
        internal static string Speaker(string token)
        {
            switch (token)
            {
                case "Vous": return Get("You");
                case "Réflexion": return Get("Reasoning");
                case "Vérification": return Get("Verification");
                case "Outil": return Get("Tool");
                case "Erreur": return Get("Error");
                case "Intervention": return Get("Turn");
                case "Assistant": return Get("Assistant");
                default: return token;
            }
        }
    }
}
