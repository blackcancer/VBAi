using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CodexVBE
{
    /// <summary>Mode de thème choisi pour l’interface.</summary>
    internal enum ThemeChoice
    {
        /// <summary>Suit la préférence d’apparence de Windows.</summary>
        System,
        /// <summary>Force les couleurs claires.</summary>
        Light,
        /// <summary>Force les couleurs sombres.</summary>
        Dark
    }
        /// <summary>Résout les couleurs de l’interface et applique le thème aux contrôles WinForms.</summary>
    internal static class UiTheme
    {
        /// <summary>Applique un thème visuel natif à une fenêtre donnée.</summary>
        /// <param name="window">Poignée Win32 de la fenêtre à styliser.</param>
        /// <param name="app">Nom du thème natif à appliquer, ou null pour le thème par défaut.</param>
        /// <param name="ids">Identifiant de sous-style natif, ou null.</param>
        /// <returns>Code de résultat HRESULT retourné par uxtheme.</returns>
        [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr window, string app, string ids);
        /// <summary>Configure un attribut DWM sur une fenêtre donnée.</summary>
        /// <param name="window">Poignée Win32 de la fenêtre à styliser.</param>
        /// <param name="attribute">Identifiant de l’attribut DWM à configurer.</param>
        /// <param name="value">Valeur de l’attribut DWM transmis par référence.</param>
        /// <param name="size">Taille en octets de la valeur fournie.</param>
        /// <returns>Code de résultat HRESULT retourné par DWM.</returns>
        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        // Environment reads and storage can be substituted without changing Windows preferences.
        /// <summary>Chemin du fichier qui conserve le choix de thème utilisateur.</summary>
        internal static string FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "theme.txt");
        /// <summary>Fournit l’état de contraste élevé du système.</summary>
        internal static Func<bool> HighContrast = () => SystemInformation.HighContrast;
        /// <summary>Fournit la couleur de fenêtre système pour le contraste élevé.</summary>
        internal static Func<Color> WindowColor = () => SystemColors.Window;
        /// <summary>Lit la préférence Windows AppsUseLightTheme.</summary>
        internal static Func<object> ReadSystemTheme = () => Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1);
        /// <summary>Lit le choix de thème stocké.</summary>
        internal static Func<string, string> ReadTheme = File.ReadAllText;
        /// <summary>Enregistre le choix de thème sélectionné.</summary>
        internal static Action<string, string> WriteTheme = File.WriteAllText;
        /// <summary>Choix persistant du thème, initialisé au chargement du fichier.</summary>
        /// <value>Choix courant du thème.</value>
        internal static ThemeChoice Choice { get; private set; } = Load();
        /// <summary>Notifies subscribers when changed occurs.</summary>
        internal static event Action Changed;
        /// <summary>Indique si le thème effectivement résolu est sombre.</summary>
        /// <value>Valeur déterminée selon le contraste élevé, le choix et la préférence Windows.</value>
        internal static bool Dark
        {
            get {
                if (HighContrast()) return WindowColor().GetBrightness() < 0.5f;
                if (Choice != ThemeChoice.System) return Choice == ThemeChoice.Dark;
                try { return (int?)ReadSystemTheme() == 0; } catch { return false; }
            }
        }
        /// <summary>Couleur de surface des champs, listes et grilles.</summary>
        /// <value>Couleur utilisée par les surfaces de saisie et les listes.</value>
        internal static Color Surface { get { return HighContrast() ? SystemColors.Window : Dark ? Color.FromArgb(30, 34, 42) : Color.White; } }
        /// <summary>Couleur de fond des conteneurs.</summary>
        /// <value>Couleur de fond utilisée pour les autres contrôles.</value>
        internal static Color Background { get { return HighContrast() ? SystemColors.Control : Dark ? Color.FromArgb(22, 26, 33) : Color.FromArgb(248, 250, 252); } }
        /// <summary>Couleur du texte au premier plan.</summary>
        /// <value>Couleur du texte selon le thème actif.</value>
        internal static Color Foreground { get { return HighContrast() ? SystemColors.WindowText : Dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(30, 41, 59); } }
        /// <summary>Subtle boundary shared by cards, fields and menus.</summary>
        internal static Color Border => HighContrast() ? SystemColors.WindowText : Dark ? Color.FromArgb(61, 68, 80) : Color.FromArgb(213, 220, 230);
        /// <summary>Common keyboard focus outline for all input controls.</summary>
        internal static Color FocusBorder => HighContrast() ? SystemColors.Highlight : Dark ? Color.FromArgb(96, 165, 250) : Color.FromArgb(37, 99, 235);
        /// <summary>Secondary text without reducing disabled-state legibility.</summary>
        internal static Color Muted => HighContrast() ? SystemColors.GrayText : Dark ? Color.FromArgb(155, 165, 180) : Color.FromArgb(94, 106, 124);
        /// <summary>Detects hosted Designer controls even after the design license context has ended.</summary>
        internal static bool IsDesignPreview(Control control)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return true;
            for (var current = control; current != null; current = current.Parent)
                if (current.Site?.DesignMode == true) return true;
            return false;
        }
        /// <summary>Uses the form's actual palette in Visual Studio instead of the user's runtime preference.</summary>
        internal static Color BackgroundFor(Control control) => control.Parent?.BackColor ?? control.BackColor;
        /// <summary>Selected surfaces match the design form in Visual Studio.</summary>
        internal static Color SurfaceFor(Control control) => SystemInformation.HighContrast ? SystemColors.Window : BackgroundFor(control).GetBrightness() < .5f ? Color.FromArgb(30, 34, 42) : Color.White;
        /// <summary>Designer labels inherit the form's foreground.</summary>
        internal static Color ForegroundFor(Control control) => control.Parent?.ForeColor ?? control.ForeColor;
        /// <summary>Field boundaries use the preview palette when hosted by a Designer.</summary>
        internal static Color BorderFor(Control control) => SystemInformation.HighContrast ? SystemColors.WindowText : BackgroundFor(control).GetBrightness() < .5f ? Color.FromArgb(61, 68, 80) : Color.FromArgb(213, 220, 230);
        /// <summary>Designer focus follows Windows rather than application settings.</summary>
        internal static Color FocusBorderFor(Control control) => SystemInformation.HighContrast ? SystemColors.Highlight : BackgroundFor(control).GetBrightness() < .5f ? Color.FromArgb(96, 165, 250) : Color.FromArgb(37, 99, 235);
        /// <summary>Couleur de fond d’un changement VBA ajouté.</summary>
        /// <value>Couleur de fond des changements ajoutés.</value>
        internal static Color Added { get { return Dark ? Color.FromArgb(24, 64, 42) : Color.FromArgb(232, 247, 237); } }
        /// <summary>Couleur de fond d’un changement VBA supprimé.</summary>
        /// <value>Couleur de fond des changements supprimés.</value>
        internal static Color Removed { get { return Dark ? Color.FromArgb(78, 35, 40) : Color.FromArgb(255, 240, 240); } }
        /// <summary>Abonne le thème aux changements de préférences Windows.</summary>
        static UiTheme() { SystemEvents.UserPreferenceChanged += PreferencesChanged; }
        /// <summary>Propage un changement de préférence visuelle du système.</summary>
        /// <param name="sender">Objet système ou contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de la préférence Windows modifiée.</param>
        private static void PreferencesChanged(object sender, UserPreferenceChangedEventArgs e) { Changed?.Invoke(); }
        /// <summary>Charge le thème enregistré ou retourne le suivi du système si le fichier est absent ou invalide.</summary>
        /// <returns>Choix valide enregistré, ou suivi des préférences système.</returns>
        private static ThemeChoice Load() { try { ThemeChoice value; if (Enum.TryParse(ReadTheme(FileName), out value) && Enum.IsDefined(typeof(ThemeChoice), value)) return value; } catch { } return ThemeChoice.System; }
        /// <summary>Enregistre le choix fourni et notifie les vues abonnées.</summary>
        /// <param name="choice">Mode de thème à enregistrer.</param>
        internal static void Select(ThemeChoice choice)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FileName)); WriteTheme(FileName, choice.ToString());
            Choice = choice; Changed?.Invoke();
        }
        /// <summary>Lie l’application du thème et les désabonnements au cycle de vie de la fenêtre.</summary>
        /// <param name="form">Fenêtre dont le cycle de vie pilote l’application du thème.</param>
        internal static void Attach(Form form)
        {
            if (IsDesignPreview(form)) return;
            Action update = () => { if (!form.IsDisposed) { if (form.InvokeRequired) form.BeginInvoke(new Action(() => Apply(form))); else Apply(form); } };
            Changed += update;
            form.Disposed += (s, e) => { Changed -= update; };
            Apply(form);
        }
        /// <summary>Applique les couleurs, styles et gestionnaires de dessin au contrôle et à ses enfants.</summary>
        /// <param name="control">Contrôle dont les propriétés visuelles sont mises à jour.</param>
        internal static void Apply(Control control)
        {
            if (IsDesignPreview(control)) return;
            control.HandleCreated -= ApplyNativeTheme;
            control.HandleCreated += ApplyNativeTheme;
            if (control.IsHandleCreated) ApplyNativeTheme(control, EventArgs.Empty);
            control.BackColor = control is TextBoxBase || control is ListControl || control is DataGridView ? Surface : Background;
            control.ForeColor = Foreground;
            if (control.ContextMenuStrip != null) ApplyMenu(control.ContextMenuStrip);
            if (control is TextBoxBase textBox) {
                bool transcript = textBox.BorderStyle == BorderStyle.None && (textBox is UiTextBox || textBox is RichTextBox rich && rich.ReadOnly);
                textBox.BorderStyle = transcript ? BorderStyle.None : BorderStyle.FixedSingle;
                if (transcript) textBox.BackColor = Background;
            }
            if (control is ListBox listBox) listBox.BorderStyle = BorderStyle.FixedSingle;
            if (control is CheckBox checkBox) checkBox.FlatStyle = FlatStyle.Flat;
            if (control is Button button) { button.FlatStyle = FlatStyle.Flat; button.FlatAppearance.BorderColor = Dark ? Color.FromArgb(75, 85, 99) : Color.FromArgb(203, 213, 225); }
            if (control is DataGridView grid) { grid.BackgroundColor = Surface; grid.DefaultCellStyle.BackColor = Surface; grid.DefaultCellStyle.ForeColor = Foreground; grid.EnableHeadersVisualStyles = false; grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; grid.ColumnHeadersDefaultCellStyle.BackColor = Background; grid.ColumnHeadersDefaultCellStyle.ForeColor = Foreground;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Background; grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Foreground;
                grid.GridColor = Dark ? Color.FromArgb(71, 85, 105) : Color.FromArgb(203, 213, 225);
                grid.CellFormatting -= FormatDiffCell; grid.CellFormatting += FormatDiffCell; }
            if (control is ComboBox combo) { combo.FlatStyle = combo is UiComboBox ? FlatStyle.Standard : FlatStyle.Flat; combo.DrawMode = DrawMode.OwnerDrawFixed; combo.DrawItem -= DrawCombo; if (!(combo is UiComboBox)) combo.DrawItem += DrawCombo; }
            foreach (Control child in control.Controls) Apply(child);
            control.Invalidate();
        }
        /// <summary>Styles context commands consistently when they open, including after a theme change.</summary>
        internal static void ApplyMenu(ContextMenuStrip menu)
        {
            menu.Opening -= MenuOpening; menu.Opening += MenuOpening;
            menu.BackColor = Surface; menu.ForeColor = Foreground;
            menu.Renderer = new ToolStripProfessionalRenderer(new MenuColors());
            ApplyMenuItems(menu.Items);
        }
        private static void MenuOpening(object sender, CancelEventArgs e) { ApplyMenu((ContextMenuStrip)sender); }
        private static void ApplyMenuItems(ToolStripItemCollection items)
        {
            foreach (ToolStripItem item in items) {
                item.ForeColor = Foreground; item.BackColor = Surface;
                if (item is ToolStripDropDownItem parent) ApplyMenuItems(parent.DropDownItems);
            }
        }
        private sealed class MenuColors : ProfessionalColorTable
        {
            public override Color ToolStripDropDownBackground => Surface;
            public override Color ImageMarginGradientBegin => Surface;
            public override Color ImageMarginGradientMiddle => Surface;
            public override Color ImageMarginGradientEnd => Surface;
            public override Color MenuItemSelected => HighContrast() ? SystemColors.Highlight : Dark ? Color.FromArgb(48, 61, 81) : Color.FromArgb(229, 238, 253);
            public override Color MenuItemBorder => Border;
            public override Color MenuBorder => Border;
            public override Color SeparatorDark => Border;
            public override Color SeparatorLight => Surface;
        }
        /// <summary>Applique le thème natif à la poignée du contrôle sans modifier la préférence du processus hôte.</summary>
        /// <param name="sender">Objet système ou contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement associé.</param>
        private static void ApplyNativeTheme(object sender, EventArgs e)
        {
            var control = (Control)sender;
            if (!control.IsHandleCreated || IsDesignPreview(control)) return;
            bool dark = Dark && !HighContrast();
            // Per-window styling only: do not change Office/SOLIDWORKS process-wide theme policy.
            if (control is TextBoxBase || control is ListBox || control is ComboBox || control is DataGridView || control is ScrollableControl scroll && scroll.AutoScroll)
                SetWindowTheme(control.Handle, dark ? (control is ComboBox ? "DarkMode_CFD" : "DarkMode_Explorer") : null, null);
            if (control is Form) { int value = dark ? 1 : 0; DwmSetWindowAttribute(control.Handle, 20, ref value, sizeof(int)); }
        }
        /// <summary>Dessine une entrée de ComboBox avec les couleurs et le texte du thème.</summary>
        /// <param name="sender">Objet système ou contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement associé.</param>
        private static void DrawCombo(object sender, DrawItemEventArgs e)
        {
            var combo = (ComboBox)sender;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var brush = new SolidBrush(selected ? SystemColors.Highlight : Surface)) e.Graphics.FillRectangle(brush, e.Bounds);
            string text = e.Index >= 0 ? combo.GetItemText(combo.Items[e.Index]) : combo.Text;
            TextRenderer.DrawText(e.Graphics, text, e.Font, e.Bounds, selected ? SystemColors.HighlightText : Foreground, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }
        /// <summary>Colore les cellules de grille marquées comme lignes VBA ajoutées ou supprimées.</summary>
        /// <param name="sender">Objet système ou contrôle à l’origine de l’événement.</param>
        /// <param name="e">Données de l’événement associé.</param>
        private static void FormatDiffCell(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (grid.VirtualMode) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex >= grid.Rows.Count) return;
            string kind = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag as string;
            if (kind == "vba-added" || kind == "vba-removed") { e.CellStyle.BackColor = kind == "vba-added" ? Added : Removed; e.CellStyle.ForeColor = Foreground; }
        }
        /// <summary>Convertit certaines couleurs claires connues vers leurs équivalents du thème sombre.</summary>
        /// <param name="hex">Couleur hexadécimale à convertir.</param>
        /// <returns>Couleur hexadécimale mappée, ou valeur d’origine si aucun remplacement ne s’applique.</returns>
        internal static string Map(string hex)
        {
            if (!Dark) return hex;
            switch (hex.ToUpperInvariant())
            {
                case "#FFFFFF": return "#1E222A";
                case "#F8FAFC": return "#161A21";
                case "#F1F5F9": case "#EFF6FF": return "#273345";
                case "#DBEAFE": case "#E2E8F0": return "#526176";
                case "#0F172A": case "#1E293B": case "#334155": case "#475569": return "#E2E8F0";
                case "#64748B": return "#AEBBD0";
                case "#FEF2F2": case "#FFF0F0": return "#4E2328";
                case "#E8F7ED": return "#18402A";
                case "#15803D": case "#166534": return "#86EFAC";
                case "#991B1B": return "#FCA5A5";
                default: return hex;
            }
        }
    }
}
