namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Reflection;
    using System.Windows.Forms;
    using VBAi;

    /// <summary>Vérifie le modèle de lignes du diff dans le contrôle GitWindow.</summary>
    public sealed partial class GitWindowStateTests
    {
        /// <summary>Numérote les lignes à partir de un et marque uniquement les côtés modifiés.</summary>
        [TestMethod]
        [STATestMethod]
        public void SwitchingDiffModePreservesDesignerColumnsAndLineNumberWidth()
        {
            using (var view = new CodeDiffView())
            {
                var grid = DiffGrid(view);
                var oldLine = grid.Columns[0];
                var after = grid.Columns[3];
                var unified = (CheckBox)typeof(CodeDiffView).GetField("unified", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(view);
                oldLine.Width = 73;
                view.ShowDiff("same\nremoved", "same\nadded");
                unified.Checked = true;
                Assert.AreSame(oldLine, grid.Columns[0]);
                Assert.AreSame(after, grid.Columns[3]);
                Assert.AreEqual(73, oldLine.Width);
                Assert.IsFalse(grid.Columns[1].Visible);
                Assert.AreEqual("removed", grid.Rows[1].Cells[3].Value);
                Assert.AreEqual("added", grid.Rows[2].Cells[3].Value);
                unified.Checked = false;
                Assert.IsTrue(grid.Columns[1].Visible);
                Assert.AreEqual(73, oldLine.Width);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void DiffRowsUseOneBasedLineNumbersAndMarkOnlyChangedSides()
        {
            using (var window = new GitWindow())
            {
                var view = Field<CodeDiffView>(window, "diff");
                view.ShowDiff("same\nremoved", "same\nadded");
                var grid = DiffGrid(view);
                Assert.AreEqual(2, grid.RowCount);
                var rows = DiffModel.Build("same\nremoved", "same\nadded", false, false);
                Assert.AreEqual(1, rows[0].Old);
                Assert.AreEqual(1, rows[0].New);
                Assert.AreEqual("removed", rows[1].Left);
                Assert.AreEqual("added", rows[1].Right);
                Assert.AreEqual(2, rows[1].Old);
                Assert.AreEqual(2, rows[1].New);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Drawing;
    using System.Linq;
    using System.Windows.Forms;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    /// <summary>Vérifie le layout, les événements de pliage et le dessin des cellules virtuelles.</summary>
    [TestClass]
    public sealed class CodeDiffViewCoverageTests
    {
        /// <summary>Suit les lignes visibles lors de la navigation, recherche et changement de pliage.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void LayoutNavigationSearchAndFoldEventsTrackVisibleRows()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new CodeDiffView())
            {
                var grid = UiInvoke.Field<DataGridView>(view, "grid"); var unified = UiInvoke.Field<CheckBox>(view, "unified"); var fold = UiInvoke.Field<CheckBox>(view, "collapse"); var search = UiInvoke.Field<TextBox>(view, "search");
                view.ShowDiff(null, null); Assert.AreEqual(0, grid.RowCount); UiInvoke.Call(typeof(CodeDiffView), "MoveChange", view, 1);
                view.ShowDiff("same", "same"); UiInvoke.Call(typeof(CodeDiffView), "MoveChange", view, 1); Assert.AreEqual(1, grid.RowCount);
                var before = string.Join("\n", Enumerable.Range(0, 25).Select(i => "line " + i)); var after = before.Replace("line 5\n", "changed five\n").Replace("line 20\n", "changed twenty\n"); view.ShowDiff(before, after); Assert.AreEqual(4, grid.Columns.Count);
                grid.CurrentCell = null; UiInvoke.Call(typeof(CodeDiffView), "Next_Click", view, null, EventArgs.Empty); var first = grid.CurrentCell.RowIndex;
                UiInvoke.Call(typeof(CodeDiffView), "Next_Click", view, null, EventArgs.Empty); Assert.AreNotEqual(first, grid.CurrentCell.RowIndex);
                UiInvoke.Call(typeof(CodeDiffView), "Previous_Click", view, null, EventArgs.Empty); Assert.AreEqual(first, grid.CurrentCell.RowIndex);
                UiInvoke.Call(typeof(CodeDiffView), "Find_Click", view, null, EventArgs.Empty); Assert.IsTrue(fold.Checked);
                search.Text = "CHANGED TWENTY"; grid.CurrentCell = null; UiInvoke.Call(typeof(CodeDiffView), "Find_Click", view, null, EventArgs.Empty); Assert.IsFalse(fold.Checked); Assert.AreEqual(20, grid.CurrentCell.RowIndex);
                grid.CurrentCell = null; UiInvoke.Call(typeof(CodeDiffView), "Find_Click", view, null, EventArgs.Empty); Assert.AreEqual(20, grid.CurrentCell.RowIndex);
                var selected = grid.CurrentCell; search.Text = "absent"; UiInvoke.Call(typeof(CodeDiffView), "Find_Click", view, null, EventArgs.Empty); Assert.AreSame(selected, grid.CurrentCell);
                unified.Checked = true; Assert.AreEqual(4, grid.Columns.Count); Assert.IsFalse(grid.Columns[1].Visible); Assert.IsTrue(grid.Columns[2].Visible); Assert.IsTrue(grid.Columns[3].Visible); search.Text = "changed five";
                foreach (var key in new[] { Keys.Escape, Keys.Enter }) { var args = new KeyEventArgs(key); UiInvoke.Call(typeof(CodeDiffView), "Search_KeyDown", view, search, args); Assert.AreEqual(key == Keys.Enter, args.SuppressKeyPress); }
                Assert.AreEqual("changed five", UiInvoke.Field<System.Collections.Generic.List<DiffRow>>(view, "visible")[grid.CurrentCell.RowIndex].Right);
                fold.Checked = true; var visible = UiInvoke.Field<System.Collections.Generic.List<DiffRow>>(view, "visible"); int foldIndex = visible.FindIndex(r => r.Fold);
                foreach (var row in new[] { -1, visible.Count, visible.FindIndex(r => !r.Fold) }) { UiInvoke.Call(typeof(CodeDiffView), "ExpandContext", view, grid, new DataGridViewCellEventArgs(0, row)); Assert.IsTrue(fold.Checked); }
                UiInvoke.Call(typeof(CodeDiffView), "ExpandContext", view, grid, new DataGridViewCellEventArgs(0, foldIndex)); Assert.IsFalse(fold.Checked);
                view.ShowDiff("old", "new"); UiInvoke.Call(typeof(CodeDiffView), "MoveChange", view, 1); var rowIndex = grid.CurrentCell.RowIndex; UiInvoke.Call(typeof(CodeDiffView), "MoveChange", view, 1); Assert.AreEqual(rowIndex, grid.CurrentCell.RowIndex);
                UiInvoke.Call(typeof(CodeDiffView), "Dispose", view, false); Assert.IsFalse(view.IsDisposed);
            }
        }
        /// <summary>Fournit les valeurs virtuelles et formate les cellules dans les deux layouts.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void VirtualValuesAndCellFormattingCoverBothLayoutsAndInvalidIndices()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new CodeDiffView())
            {
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => new DataGridViewCellValueEventArgs(0, -1));
                var grid = UiInvoke.Field<DataGridView>(view, "grid"); var unified = UiInvoke.Field<CheckBox>(view, "unified"); UiInvoke.Field<CheckBox>(view, "collapse").Checked = false;
                foreach (var single in new[] { false, true })
                {
                    unified.Checked = single; view.ShowDiff("same\nremoved", "same\nadded");
                    var rows = UiInvoke.Field<System.Collections.Generic.List<DiffRow>>(view, "visible");
                    foreach (var index in new[] { -1, rows.Count })
                    {
                        if (index >= 0) { var value = new DataGridViewCellValueEventArgs(0, index) { Value = "sentinel" }; UiInvoke.Call(typeof(CodeDiffView), "ValueNeeded", view, grid, value); Assert.AreEqual("sentinel", value.Value); }
                        var format = new DataGridViewCellFormattingEventArgs(0, index, "sentinel", typeof(string), new DataGridViewCellStyle { BackColor = Color.Magenta }); UiInvoke.Call(typeof(CodeDiffView), "FormatCell", view, grid, format); Assert.AreEqual(Color.Magenta, format.CellStyle.BackColor);
                    }
                    for (int row = 0; row < rows.Count; row++)
                        for (int col = 0; col < grid.Columns.Count; col++)
                        {
                            var expected = single ? new object[] { rows[row].Old, rows[row].Left, rows[row].New, rows[row].Unified } : new object[] { rows[row].Old, rows[row].Left, rows[row].New, rows[row].Right };
                            var value = new DataGridViewCellValueEventArgs(col, row); UiInvoke.Call(typeof(CodeDiffView), "ValueNeeded", view, grid, value); Assert.AreEqual(expected[col], value.Value);
                            var format = new DataGridViewCellFormattingEventArgs(col, row, value.Value, typeof(string), new DataGridViewCellStyle()); UiInvoke.Call(typeof(CodeDiffView), "FormatCell", view, grid, format);
                            Assert.AreEqual(rows[row].Hunk < 0 ? UiTheme.Surface : (single ? rows[row].New.HasValue : col >= 2) ? UiTheme.Added : UiTheme.Removed, format.CellStyle.BackColor); Assert.AreEqual(UiTheme.Foreground, format.CellStyle.ForeColor);
                        }
                }
            }
        }
        /// <summary>Peint les entêtes, valeurs absentes, numéros, code sélectionné et contenu rogné.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void CellPaintingHandlesHeadersMissingValuesNumbersCodeSelectionAndClipping()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new CodeDiffView())
            using (var bitmap = new Bitmap(500, 100))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => new DataGridViewCellValueEventArgs(0, -1));
                var grid = UiInvoke.Field<DataGridView>(view, "grid"); var unified = UiInvoke.Field<CheckBox>(view, "unified"); UiInvoke.Field<CheckBox>(view, "collapse").Checked = false; view.ShowDiff("old", "new"); var handle = grid.Handle;
                foreach (var single in new[] { false, true })
                    foreach (var selected in new[] { false, true })
                        foreach (var customFont in new[] { false, true })
                            foreach (var size in new[] { 8, 450 })
                            {
                                unified.Checked = single;
                                for (int col = -1; col < grid.Columns.Count; col++)
                                    foreach (var value in new object[] { null, 1, "Sub Test() ' comment" })
                                    {
                                        var style = new DataGridViewCellStyle { BackColor = Color.White, ForeColor = Color.Black, SelectionForeColor = Color.Purple, Font = customFont ? view.Font : null };
                                        var args = DiffPaintFixture.Args(grid, graphics, new Rectangle(0, 0, size, 30), col < 0 ? -1 : 0, col, selected ? DataGridViewElementStates.Selected : DataGridViewElementStates.None, value, style);
                                        UiInvoke.Call(typeof(CodeDiffView), "PaintCell", view, grid, args); Assert.AreEqual(col >= 0, args.Handled);
                                        if (size == 450 && value is string && (col == 1 || col == 3)) { int ink = (selected ? Color.Purple : VbaSyntax.Color("keyword")).ToArgb(); Assert.IsTrue(Enumerable.Range(0, 450).Any(x => Enumerable.Range(0, 30).Any(y => bitmap.GetPixel(x, y).ToArgb() == ink)), "Code must draw the expected ink."); }
                                    }
                            }
                var invalid = new DataGridViewCellPaintingEventArgs(grid, graphics, new Rectangle(0, 0, 500, 100), new Rectangle(0, 0, 450, 30), 0, -1, DataGridViewElementStates.None, "code", "code", null, new DataGridViewCellStyle(), new DataGridViewAdvancedBorderStyle(), DataGridViewPaintParts.All); UiInvoke.Call(typeof(CodeDiffView), "PaintCell", view, grid, invalid); Assert.IsFalse(invalid.Handled);
                Assert.AreEqual(Color.White.ToArgb(), bitmap.GetPixel(445, 15).ToArgb());
            }
        }
    }
}
