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
    /// <summary>Affiche dans un panneau WPF une comparaison navigable et recherchable de deux versions de code.</summary>
    internal sealed class ChatDiffView : DockPanel
    {
        /// <summary>Texte de la version initiale.</summary>
        private readonly string before, after;
        /// <summary>Grille contenant les lignes et colonnes de la comparaison.</summary>
        private readonly DataGrid grid;
        /// <summary>Options d’affichage côte à côte et de repli du contexte inchangé.</summary>
        private readonly CheckBox unified, fold;
        /// <summary>Champ de recherche dans les lignes affichées.</summary>
        private readonly TextBox search;
        /// <summary>Lignes actuellement produites par le modèle de diff.</summary>
        private List<DiffRow> rows;
        /// <summary>Convertit une couleur GDI en pinceau WPF opaque.</summary>
        /// <param name="color">Couleur source.</param>
        /// <returns>Un pinceau WPF de même composante RGB.</returns>
        private static Brush Brush(System.Drawing.Color color) { return new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B)); }
        /// <summary>Crée le panneau de comparaison et initialise ses commandes, sa grille et ses lignes.</summary>
        /// <param name="before">Texte initial.</param>
        /// <param name="after">Texte modifié.</param>
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
        /// <summary>Reconstruit les lignes et colonnes selon les options d’affichage sélectionnées.</summary>
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
        /// <summary>Ajoute une colonne de code dont les fragments sont colorés selon la syntaxe VBA.</summary>
        /// <param name="field">Propriété de <see cref="DiffRow"/> à afficher : gauche, droite ou unifiée.</param>
        /// <param name="title">Titre visible de la colonne.</param>
        private void AddCodeColumn(string field, string title)
        {
            var cell = new FrameworkElementFactory(typeof(ContentControl));
            cell.SetBinding(ContentControl.ContentProperty, new Binding { Converter = new SyntaxConverter(field) });
            grid.Columns.Add(new DataGridTemplateColumn { Header = title, CellTemplate = new DataTemplate { VisualTree = cell }, Width = DataGridLength.SizeToCells, MinWidth = 240 });
        }
        /// <summary>Convertit une ligne de diff en bloc de texte WPF avec coloration syntaxique.</summary>
        private sealed class SyntaxConverter : IValueConverter
        {
            /// <summary>Nom du champ de la ligne utilisé pour produire le texte.</summary>
            private readonly string field;
            /// <summary>Crée un convertisseur pour la colonne donnée.</summary>
            /// <param name="field">Champ de la ligne de diff à afficher.</param>
            internal SyntaxConverter(string field) { this.field = field; }
            /// <summary>Construit le bloc WPF correspondant à la valeur de ligne reçue.</summary>
            /// <param name="value">Ligne de diff à convertir.</param>
            /// <param name="type">Type cible demandé par WPF.</param>
            /// <param name="parameter">Paramètre de liaison éventuel, ignoré.</param>
            /// <param name="culture">Culture de conversion, ignorée.</param>
            /// <returns>Un bloc de texte coloré, ou <see langword="null"/> si la valeur n’est pas une ligne de diff.</returns>
            public object Convert(object value, Type type, object parameter, CultureInfo culture)
            {
                var row = value as DiffRow; if (row == null) return null;
                string source = field == "Left" ? row.Left : field == "Right" ? row.Right : row.Unified;
                var text = new TextBlock { Padding = new Thickness(3, 1, 3, 1), Background = Brush(row.Hunk < 0 ? UiTheme.Surface : field == "Left" || (field == "Unified" && !row.New.HasValue) ? UiTheme.Removed : UiTheme.Added) };
                foreach (var part in VbaSyntax.Parts(source ?? "")) text.Inlines.Add(new Run(part.Text) { Foreground = Brush(VbaSyntax.Color(part.Kind)) });
                return text;
            }
            /// <summary>La conversion inverse n’est pas prise en charge par cette liaison en lecture seule.</summary>
            /// <param name="value">Valeur de source proposée par WPF.</param>
            /// <param name="type">Type cible demandé par WPF.</param>
            /// <param name="parameter">Paramètre de liaison éventuel.</param>
            /// <param name="culture">Culture de conversion.</param>
            /// <returns>Cette méthode ne retourne pas de valeur.</returns>
            /// <exception cref="NotSupportedException">La conversion inverse n’est pas implémentée.</exception>
            public object ConvertBack(object value, Type type, object parameter, CultureInfo culture) { throw new NotSupportedException(); }
        }
        /// <summary>Sélectionne une ligne et la fait défiler dans la grille.</summary>
        /// <param name="row">Ligne à sélectionner.</param>
        private void Select(DiffRow row) { grid.SelectedItem = row; grid.ScrollIntoView(row); }
        /// <summary>Parcourt les lignes jusqu’à la prochaine zone modifiée distincte, dans le sens demandé.</summary>
        /// <param name="step">Direction du parcours : -1 vers les changements précédents, 1 vers les suivants.</param>
        private void Move(int step)
        {
            int start = grid.SelectedIndex; int hunk = (grid.SelectedItem as DiffRow)?.Hunk ?? -1;
            for (int n = 1; n <= rows.Count; n++) { var row = rows[(start + step * n + rows.Count * 2) % rows.Count]; if (row.Hunk >= 0 && row.Hunk != hunk) { Select(row); return; } }
        }
        /// <summary>Recherche le texte saisi en parcourant les lignes, déplie le contexte si nécessaire et sélectionne la première correspondance.</summary>
        private void Find()
        {
            if (search.Text.Length == 0) return;
            if (fold.IsChecked == true) { fold.IsChecked = false; Rebuild(); }
            int start = grid.SelectedIndex;
            for (int n = 1; n <= rows.Count; n++) { var row = rows[(start + n) % rows.Count]; if (((row.Left ?? "") + "\n" + row.Right).IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0) { Select(row); return; } }
        }
    }
}
