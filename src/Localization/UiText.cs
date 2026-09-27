using System;
using System.ComponentModel;
using System.Globalization;
using System.Resources;
using System.Windows.Forms;

namespace CodexVBE
{
    // Explicit culture avoids changing the host's thread culture or VBA number/date formatting.
    internal static class UiText
    {
        private static readonly ResourceManager English = new ResourceManager("CodexVBE.Localization.UiStrings", typeof(UiText).Assembly);
        private static readonly ResourceManager French = new ResourceManager("CodexVBE.Localization.UiStringsFrench", typeof(UiText).Assembly);
        internal static CultureInfo Culture { get; private set; } = Supported(CultureInfo.CurrentUICulture);

        internal static CultureInfo Supported(CultureInfo culture)
        {
            return CultureInfo.GetCultureInfo(culture?.TwoLetterISOLanguageName == "fr" ? "fr-FR" : "en-US");
        }

        internal static void Initialize(object vbe)
        {
            Culture = Detect(vbe, CultureInfo.CurrentUICulture);
            LoadLog.Write("VBAi interface language: " + Culture.Name);
        }

        internal static CultureInfo Detect(object vbe, CultureInfo fallback)
        {
            bool hasMenu = false;
            // Actual VBE captions take precedence over the operating system's display language.
            try
            {
                foreach (dynamic bar in ((dynamic)vbe).CommandBars)
                {
                    if ((int)bar.Type != 1) continue;
                    hasMenu = true;
                    foreach (dynamic control in bar.Controls)
                    {
                        string caption = ((string)control.Caption ?? "").Replace("&", "").Trim().ToLowerInvariant();
                        if (caption == "affichage" || caption == "outils") return Supported(CultureInfo.GetCultureInfo("fr"));
                        if (caption == "view" || caption == "tools") return Supported(CultureInfo.GetCultureInfo("en"));
                    }
                }
            }
            catch (Exception ex) { LoadLog.Write("VBE UI language unavailable: " + ex.Message); }
            return Supported(hasMenu ? CultureInfo.GetCultureInfo("en") : fallback);
        }

        internal static string Get(string english)
        {
            if (english == null) return null;
            if (english == "DOCUMENT CONVERSATIONS") return Get("Document conversations").ToUpperInvariant();
            if (english == "YOU") return Get("You").ToUpperInvariant();
            return (Culture.TwoLetterISOLanguageName == "fr" ? French.GetString(english, CultureInfo.InvariantCulture) : null)
                ?? English.GetString(english, CultureInfo.InvariantCulture) ?? english;
        }

        // Only called immediately after InitializeComponent, before document/user data is populated.
        // Designer captions stay editable in English; all translated strings live in resource files.
        internal static void Apply(Control control, IContainer components, params ToolTip[] additionalTips)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
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
            var combo = control as ComboBox;
            if (combo != null)
                for (int i = 0; i < combo.Items.Count; i++)
                    if (combo.Items[i] is string) combo.Items[i] = Get((string)combo.Items[i]);
            var grid = control as DataGridView;
            if (grid != null)
                foreach (DataGridViewColumn column in grid.Columns) column.HeaderText = Get(column.HeaderText);
            foreach (Control child in control.Controls) Apply(child, components, additionalTips);
            if (control is Form && components != null)
                foreach (IComponent component in components.Components)
                    if (component is ContextMenuStrip) ApplyItems(((ContextMenuStrip)component).Items);
        }

        private static void ApplyItems(ToolStripItemCollection items)
        {
            foreach (ToolStripItem item in items)
            {
                item.Text = Get(item.Text);
                item.ToolTipText = Get(item.ToolTipText);
                item.AccessibleName = Get(item.AccessibleName);
                var dropdown = item as ToolStripDropDownItem;
                if (dropdown != null) ApplyItems(dropdown.DropDownItems);
            }
        }

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
                default: return token;
            }
        }
    }
}
