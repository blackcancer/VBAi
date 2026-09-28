using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class CodeDiffView : UserControl
    {
        private List<DiffRow> visible = new List<DiffRow>();
        private string before = "", after = "";
        public CodeDiffView() { InitializeComponent(); UiText.Apply(this, components); }
        public void ShowDiff(string oldCode, string newCode)
        {
            before = oldCode ?? ""; after = newCode ?? ""; Rebuild();
        }
        private void Rebuild()
        {
            if (grid == null) return;
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
        private void OptionsChanged(object sender, EventArgs e) { Rebuild(); }
        private void ValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= visible.Count) return;
            var row = visible[e.RowIndex];
            e.Value = e.ColumnIndex == 0 ? (object)row.Old : e.ColumnIndex == 2 ? (object)row.New :
                e.ColumnIndex == 1 ? row.Left : unified.Checked ? row.Right ?? row.Left : row.Right;
        }
        private void FormatCell(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= visible.Count) return;
            var row = visible[e.RowIndex]; bool added = unified.Checked ? row.New.HasValue : e.ColumnIndex >= 2;
            e.CellStyle.BackColor = row.Hunk < 0 ? UiTheme.Surface : added ? UiTheme.Added : UiTheme.Removed;
            e.CellStyle.ForeColor = UiTheme.Foreground;
        }
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
        private void SelectRow(int index) { grid.CurrentCell = grid.Rows[index].Cells[unified.Checked ? 3 : 1]; grid.FirstDisplayedScrollingRowIndex = index; }
        private void Previous_Click(object sender, EventArgs e) { MoveChange(-1); }
        private void Next_Click(object sender, EventArgs e) { MoveChange(1); }
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
        private void Search_KeyDown(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { Find_Click(sender, e); e.SuppressKeyPress = true; } }
        private void ExpandContext(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex >= 0 && e.RowIndex < visible.Count && visible[e.RowIndex].Fold) collapse.Checked = false; }
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
    }
}
