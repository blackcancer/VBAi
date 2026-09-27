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
