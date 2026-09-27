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

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class GitWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void SelectingChangeWithoutSnapshotClearsPreviousPreview()
        {
            using (var window = new GitWindow())
            {
                var view = Field<CodeDiffView>(window, "diff");
                view.ShowDiff("old", "new");
                var grid = DiffGrid(view);
                Assert.IsTrue(grid.RowCount > 0);
                Invoke(window, "Changes_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(0, grid.Rows.Count);
            }
        }
    }
}
