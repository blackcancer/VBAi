using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;

namespace CodexVBE
{
    internal sealed class ChatDiffView : DockPanel
    {
        private readonly string before, after;
        private readonly DataGrid grid;
        private readonly CheckBox unified, fold;
        private readonly TextBox search;
        private List<DiffRow> rows;
        private static Brush Brush(System.Drawing.Color color) { return new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B)); }
        internal ChatDiffView(string before, string after)
        {
            this.before = before; this.after = after; Height = 300; FlowDirection = FlowDirection.LeftToRight;
            var actions = new WrapPanel(); SetDock(actions, Dock.Top); Children.Add(actions);
            unified = new CheckBox { Content = UiText.Get("Unified"), IsChecked = true, Margin = new Thickness(4), Foreground = Brush(UiTheme.Foreground) };
            fold = new CheckBox { Content = UiText.Get("Fold context"), IsChecked = true, Margin = new Thickness(4), Foreground = Brush(UiTheme.Foreground) };
            search = new TextBox { Width = 120, Margin = new Thickness(4), ToolTip = UiText.Get("Search code"), Background = Brush(UiTheme.Surface), Foreground = Brush(UiTheme.Foreground) };
            actions.Children.Add(unified); actions.Children.Add(fold);
            Button add(string text, string tip, Action action) { var b = new Button { Content = text, ToolTip = tip, Padding = new Thickness(6, 2, 6, 2), Margin = new Thickness(2), Background = Brush(UiTheme.Surface), Foreground = Brush(UiTheme.Foreground) }; b.Click += (s, e) => action(); actions.Children.Add(b); return b; }
            add("↑", UiText.Get("Previous change"), () => Move(-1)); add("↓", UiText.Get("Next change"), () => Move(1));
            actions.Children.Add(search); add(UiText.Get("Find"), UiText.Get("Search code"), Find);
            search.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Enter) { Find(); e.Handled = true; } };
            grid = new DataGrid { AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false, CanUserDeleteRows = false,
                CanUserSortColumns = false, EnableRowVirtualization = true, EnableColumnVirtualization = true,
                Background = Brush(UiTheme.Surface), Foreground = Brush(UiTheme.Foreground), RowBackground = Brush(UiTheme.Surface),
                FontFamily = new FontFamily("Consolas"), FontSize = 12, RowHeight = 24, HeadersVisibility = DataGridHeadersVisibility.Column,
                GridLinesVisibility = DataGridGridLinesVisibility.None, BorderThickness = new Thickness(0) };
            Children.Add(grid);
            var headers = new Style(typeof(System.Windows.Controls.Primitives.DataGridColumnHeader));
            headers.Setters.Add(new Setter(Control.BackgroundProperty, Brush(UiTheme.Background)));
            headers.Setters.Add(new Setter(Control.ForegroundProperty, Brush(UiTheme.Foreground)));
            headers.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
            grid.ColumnHeaderStyle = headers;
            unified.Click += (s, e) => Rebuild(); fold.Click += (s, e) => Rebuild();
            grid.MouseDoubleClick += (s, e) => { if (grid.SelectedItem is DiffRow row && row.Fold) { fold.IsChecked = false; Rebuild(); } };
            Rebuild();
        }
        private void Rebuild()
        {
            bool single = unified.IsChecked == true;
            rows = DiffModel.Build(before, after, single, fold.IsChecked == true);
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridTextColumn { Header = "−", Binding = new Binding("Old"), Width = 42 });
            if (!single) AddCodeColumn("Left", UiText.Get("Before"));
            grid.Columns.Add(new DataGridTextColumn { Header = "+", Binding = new Binding("New"), Width = 42 });
            AddCodeColumn(single ? "Unified" : "Right", single ? UiText.Get("Code") : UiText.Get("After"));
            grid.ItemsSource = rows;
        }
        private void AddCodeColumn(string field, string title)
        {
            var cell = new FrameworkElementFactory(typeof(ContentControl));
            cell.SetBinding(ContentControl.ContentProperty, new Binding { Converter = new SyntaxConverter(field) });
            grid.Columns.Add(new DataGridTemplateColumn { Header = title, CellTemplate = new DataTemplate { VisualTree = cell }, Width = DataGridLength.SizeToCells, MinWidth = 240 });
        }
        private sealed class SyntaxConverter : IValueConverter
        {
            private readonly string field;
            internal SyntaxConverter(string field) { this.field = field; }
            public object Convert(object value, Type type, object parameter, CultureInfo culture)
            {
                var row = value as DiffRow; if (row == null) return null;
                string source = field == "Left" ? row.Left : field == "Right" ? row.Right : row.Unified;
                var text = new TextBlock { Padding = new Thickness(3, 1, 3, 1), Background = Brush(row.Hunk < 0 ? UiTheme.Surface : field == "Left" || (field == "Unified" && !row.New.HasValue) ? UiTheme.Removed : UiTheme.Added) };
                foreach (var part in VbaSyntax.Parts(source ?? "")) text.Inlines.Add(new Run(part.Text) { Foreground = Brush(VbaSyntax.Color(part.Kind)) });
                return text;
            }
            public object ConvertBack(object value, Type type, object parameter, CultureInfo culture) { throw new NotSupportedException(); }
        }
        private void Select(DiffRow row) { grid.SelectedItem = row; grid.ScrollIntoView(row); }
        private void Move(int step)
        {
            int start = grid.SelectedIndex; int hunk = (grid.SelectedItem as DiffRow)?.Hunk ?? -1;
            for (int n = 1; n <= rows.Count; n++) { var row = rows[(start + step * n + rows.Count * 2) % rows.Count]; if (row.Hunk >= 0 && row.Hunk != hunk) { Select(row); return; } }
        }
        private void Find()
        {
            if (search.Text.Length == 0) return;
            if (fold.IsChecked == true) { fold.IsChecked = false; Rebuild(); }
            int start = grid.SelectedIndex;
            for (int n = 1; n <= rows.Count; n++) { var row = rows[(start + n) % rows.Count]; if (((row.Left ?? "") + "\n" + row.Right).IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0) { Select(row); return; } }
        }
    }
}
