using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Présente une comparaison virtualisée de deux versions de code, avec navigation et recherche des changements.</summary>
    public sealed partial class CodeDiffView : UserControl
    {
        /// <summary>Lignes de différence actuellement affichées.</summary>
        private List<DiffRow> visible = new List<DiffRow>();
        /// <summary>Copies des textes source utilisés pour reconstruire la comparaison.</summary>
        private string before = "", after = "";
        /// <summary>Initialise les contrôles et applique les libellés localisés.</summary>
        public CodeDiffView() { InitializeComponent(); UiText.Apply(this, components); }
                /// <summary>Choisit l’affichage unifié ; la propriété reste éditable dans le Designer.</summary>
                /// <value>True to display changes in a single column; otherwise use side-by-side columns.</value>
        [System.ComponentModel.DefaultValue(false), System.ComponentModel.Category("Appearance")]
        public bool UnifiedDiff { get => unified.Checked; set => unified.Checked = value; }

        /// <summary>Affiche la comparaison entre deux contenus et remplace les entrées précédentes.</summary>
        /// <param name="oldCode">Contenu initial ; une valeur nulle est traitée comme une chaîne vide.</param>
        /// <param name="newCode">Contenu modifié ; une valeur nulle est traitée comme une chaîne vide.</param>
        public void ShowDiff(string oldCode, string newCode)
        {
            before = oldCode ?? ""; after = newCode ?? ""; Rebuild();
        }
        /// <summary>Reconstruit les lignes virtuelles et les colonnes selon les options courantes.</summary>
        private void Rebuild()
        {
            visible = DiffModel.Build(before, after, unified.Checked, collapse.Checked);
            grid.SuspendLayout();
            grid.Visible = false;
            try
            {
                grid.RowCount = 0;
                beforeColumn.Visible = !unified.Checked;
                afterColumn.HeaderText = UiText.Get(unified.Checked ? "Code" : "After");
                grid.RowCount = visible.Count;
            }
            finally
            {
                grid.Visible = true;
                grid.ResumeLayout();
                grid.Invalidate();
            }
        }
        /// <summary>Reconstruit l’affichage après un changement d’option.</summary>
        /// <param name="sender">Contrôle ayant déclenché l’événement.</param>
        /// <param name="e">Données de l’événement.</param>
        private void OptionsChanged(object sender, EventArgs e) { Rebuild(); }
        /// <summary>Fournit à la grille la valeur de la cellule demandée en mode virtualisé.</summary>
        /// <param name="sender">Grille ayant demandé la valeur.</param>
        /// <param name="e">Indices de la cellule et emplacement de la valeur à fournir.</param>
        private void ValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex >= visible.Count) return;
            var row = visible[e.RowIndex];
            e.Value = e.ColumnIndex == 0 ? (object)row.Old : e.ColumnIndex == 2 ? (object)row.New :
                e.ColumnIndex == 1 ? row.Left : unified.Checked ? row.Right ?? row.Left : row.Right;
        }
        /// <summary>Applique les couleurs de fond et de texte selon le côté ajouté, supprimé ou inchangé.</summary>
        /// <param name="sender">Grille qui formate la cellule.</param>
        /// <param name="e">Cellule et styles à formater.</param>
        private void FormatCell(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= visible.Count) return;
            var row = visible[e.RowIndex]; bool added = unified.Checked ? row.New.HasValue : e.ColumnIndex >= 2;
            e.CellStyle.BackColor = row.Hunk < 0 ? UiTheme.Surface : added ? UiTheme.Added : UiTheme.Removed;
            e.CellStyle.ForeColor = UiTheme.Foreground;
        }
        /// <summary>Dessine les marqueurs numériques et le code avec la coloration syntaxique VBA.</summary>
        /// <param name="sender">Grille qui dessine la cellule.</param>
        /// <param name="e">Contexte de dessin et limites de la cellule.</param>
        private void PaintCell(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            e.Handled = true;
            e.PaintBackground(e.CellBounds, true);
            if (e.Value == null) return;
            bool code = e.ColumnIndex == 1 || e.ColumnIndex == 3;
            var font = e.CellStyle.Font ?? grid.Font;
            if (!code)
            {
                TextRenderer.DrawText(e.Graphics, e.Value.ToString(), font, e.CellBounds,
                    e.CellStyle.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                return;
            }
            var state = e.Graphics.Save(); e.Graphics.SetClip(e.CellBounds);
            int x = e.CellBounds.X + 3, y = e.CellBounds.Y + 2;
            foreach (var part in VbaSyntax.Parts(e.Value.ToString()))
            {
                var color = (e.State & DataGridViewElementStates.Selected) != 0 ? e.CellStyle.SelectionForeColor : VbaSyntax.Color(part.Kind);
                TextRenderer.DrawText(e.Graphics, part.Text, font, new Point(x, y), color, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
                x += TextRenderer.MeasureText(e.Graphics, part.Text, font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
                if (x > e.CellBounds.Right) break;
            }
            e.Graphics.Restore(state);
        }
        /// <summary>Déplace la sélection au changement distinct suivant ou précédent, en bouclant dans la grille.</summary>
        /// <param name="direction">Sens de navigation : -1 vers le changement précédent, 1 vers le suivant.</param>
        private void MoveChange(int direction)
        {
            int current = grid.CurrentCell?.RowIndex ?? -1;
            int hunk = current >= 0 && current < visible.Count ? visible[current].Hunk : -1;
            for (int n = 1; n <= visible.Count; n++)
            {
                int index = (current + direction * n + visible.Count * 2) % visible.Count;
                if (visible[index].Hunk >= 0 && visible[index].Hunk != hunk) { SelectRow(index); return; }
            }
        }
        /// <summary>Sélectionne la cellule de code de la ligne indiquée et la fait défiler à l’écran.</summary>
        /// <param name="index">Index de la ligne dans la grille.</param>
        private void SelectRow(int index) { grid.CurrentCell = grid.Rows[index].Cells[unified.Checked ? 3 : 1]; grid.FirstDisplayedScrollingRowIndex = index; }
        /// <summary>Navigue au changement précédent en réponse au bouton correspondant.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Previous_Click(object sender, EventArgs e) { MoveChange(-1); }
        /// <summary>Navigue au changement suivant en réponse au bouton correspondant.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Next_Click(object sender, EventArgs e) { MoveChange(1); }
        /// <summary>Recherche le texte saisi à partir de la ligne courante et sélectionne la première correspondance.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Find_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(search.Text)) return;
            collapse.Checked = false;
            int start = grid.CurrentCell?.RowIndex ?? -1;
            for (int n = 1; n <= visible.Count; n++)
            {
                int i = (start + n) % visible.Count;
                if (((visible[i].Left ?? "") + "\n" + visible[i].Right).IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0) { SelectRow(i); return; }
            }
        }
        /// <summary>Lance la recherche lorsque l’utilisateur appuie sur Entrée dans le champ de recherche.</summary>
        /// <param name="sender">Champ de recherche.</param>
        /// <param name="e">Touche pressée ; Entrée est supprimée après déclenchement de la recherche.</param>
        private void Search_KeyDown(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { Find_Click(sender, e); e.SuppressKeyPress = true; } }
        /// <summary>Déplie le contexte lorsque l’utilisateur active une ligne repliée.</summary>
        /// <param name="sender">Grille contenant la cellule.</param>
        /// <param name="e">Indices de la cellule activée.</param>
        private void ExpandContext(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && e.RowIndex < visible.Count && visible[e.RowIndex].Fold) collapse.Checked = false; }
        /// <summary>Libère les composants détenus par le contrôle lorsqu’il est détruit.</summary>
        /// <param name="disposing"><see langword="true"/> si l’appel provient de <see cref="IDisposable.Dispose()"/> ; sinon <see langword="false"/>.</param>
        protected override void Dispose(bool disposing) { if (disposing) components.Dispose(); base.Dispose(disposing); }
    }
}
