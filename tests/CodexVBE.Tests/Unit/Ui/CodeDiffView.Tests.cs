namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.Serialization;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class GitWindowStateTests
    {
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
