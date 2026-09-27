using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CodexVBE
{
    internal enum ThemeChoice { System, Light, Dark }
    internal static class UiTheme
    {
        [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr window, string app, string ids);
        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        // Environment reads and storage can be substituted without changing Windows preferences.
        internal static string FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "theme.txt");
        internal static Func<bool> HighContrast = () => SystemInformation.HighContrast;
        internal static Func<Color> WindowColor = () => SystemColors.Window;
        internal static Func<object> ReadSystemTheme = () => Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1);
        internal static Func<string, string> ReadTheme = File.ReadAllText;
        internal static Action<string, string> WriteTheme = File.WriteAllText;
        internal static ThemeChoice Choice { get; private set; } = Load();
        internal static event Action Changed;
        internal static bool Dark
        {
            get {
                if (HighContrast()) return WindowColor().GetBrightness() < 0.5f;
                if (Choice != ThemeChoice.System) return Choice == ThemeChoice.Dark;
                try { return (int?)ReadSystemTheme() == 0; } catch { return false; }
            }
        }
        internal static Color Surface { get { return HighContrast() ? SystemColors.Window : Dark ? Color.FromArgb(30, 34, 42) : Color.White; } }
        internal static Color Background { get { return HighContrast() ? SystemColors.Control : Dark ? Color.FromArgb(22, 26, 33) : Color.FromArgb(248, 250, 252); } }
        internal static Color Foreground { get { return HighContrast() ? SystemColors.WindowText : Dark ? Color.FromArgb(226, 232, 240) : Color.FromArgb(30, 41, 59); } }
        internal static Color Added { get { return Dark ? Color.FromArgb(24, 64, 42) : Color.FromArgb(232, 247, 237); } }
        internal static Color Removed { get { return Dark ? Color.FromArgb(78, 35, 40) : Color.FromArgb(255, 240, 240); } }
        static UiTheme() { SystemEvents.UserPreferenceChanged += PreferencesChanged; }
        private static void PreferencesChanged(object sender, UserPreferenceChangedEventArgs e) { Changed?.Invoke(); }
        private static ThemeChoice Load() { try { ThemeChoice value; if (Enum.TryParse(ReadTheme(FileName), out value) && Enum.IsDefined(typeof(ThemeChoice), value)) return value; } catch { } return ThemeChoice.System; }
        internal static void Select(ThemeChoice choice)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FileName)); WriteTheme(FileName, choice.ToString());
            Choice = choice; Changed?.Invoke();
        }
        internal static void Attach(Form form)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            Action update = () => { if (!form.IsDisposed) { if (form.InvokeRequired) form.BeginInvoke(new Action(() => Apply(form))); else Apply(form); } };
            Changed += update;
            form.Disposed += (s, e) => { Changed -= update; };
            Apply(form);
        }
        internal static void Apply(Control control)
        {
            control.HandleCreated -= ApplyNativeTheme;
            control.HandleCreated += ApplyNativeTheme;
            if (control.IsHandleCreated) ApplyNativeTheme(control, EventArgs.Empty);
            control.BackColor = control is TextBoxBase || control is ListControl || control is DataGridView ? Surface : Background;
            control.ForeColor = Foreground;
            if (control is TextBoxBase textBox) textBox.BorderStyle = BorderStyle.FixedSingle;
            if (control is ListBox listBox) listBox.BorderStyle = BorderStyle.FixedSingle;
            if (control is CheckBox checkBox) checkBox.FlatStyle = FlatStyle.Flat;
            if (control is Button button) { button.FlatStyle = FlatStyle.Flat; button.FlatAppearance.BorderColor = Dark ? Color.FromArgb(75, 85, 99) : Color.FromArgb(203, 213, 225); }
            if (control is DataGridView grid) { grid.BackgroundColor = Surface; grid.DefaultCellStyle.BackColor = Surface; grid.DefaultCellStyle.ForeColor = Foreground; grid.EnableHeadersVisualStyles = false; grid.ColumnHeadersDefaultCellStyle.BackColor = Background; grid.ColumnHeadersDefaultCellStyle.ForeColor = Foreground;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Background; grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Foreground;
                grid.GridColor = Dark ? Color.FromArgb(71, 85, 105) : Color.FromArgb(203, 213, 225);
                grid.CellFormatting -= FormatDiffCell; grid.CellFormatting += FormatDiffCell; }
            if (control is ComboBox combo) { combo.FlatStyle = FlatStyle.Flat; combo.DrawMode = DrawMode.OwnerDrawFixed; combo.DrawItem -= DrawCombo; combo.DrawItem += DrawCombo; }
            foreach (Control child in control.Controls) Apply(child);
            control.Invalidate();
        }
        private static void ApplyNativeTheme(object sender, EventArgs e)
        {
            var control = (Control)sender;
            if (!control.IsHandleCreated || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            bool dark = Dark && !HighContrast();
            // Per-window styling only: do not change Office/SOLIDWORKS process-wide theme policy.
            if (control is TextBoxBase || control is ListBox || control is ComboBox || control is DataGridView)
                SetWindowTheme(control.Handle, dark ? (control is ComboBox ? "DarkMode_CFD" : "DarkMode_Explorer") : null, null);
            if (control is Form) { int value = dark ? 1 : 0; DwmSetWindowAttribute(control.Handle, 20, ref value, sizeof(int)); }
        }
        private static void DrawCombo(object sender, DrawItemEventArgs e)
        {
            var combo = (ComboBox)sender;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var brush = new SolidBrush(selected ? SystemColors.Highlight : Surface)) e.Graphics.FillRectangle(brush, e.Bounds);
            string text = e.Index >= 0 ? combo.GetItemText(combo.Items[e.Index]) : combo.Text;
            TextRenderer.DrawText(e.Graphics, text, e.Font, e.Bounds, selected ? SystemColors.HighlightText : Foreground, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }
        private static void FormatDiffCell(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (grid.VirtualMode) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex >= grid.Rows.Count) return;
            string kind = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag as string;
            if (kind == "vba-added" || kind == "vba-removed") { e.CellStyle.BackColor = kind == "vba-added" ? Added : Removed; e.CellStyle.ForeColor = Foreground; }
        }
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
