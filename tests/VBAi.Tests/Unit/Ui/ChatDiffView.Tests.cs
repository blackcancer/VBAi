using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Forms = System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie l’intégration du diff Designer dans le transcript WPF.</summary>
    [TestClass]
    public sealed class ChatDiffViewTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void HostedDesignerDiffKeepsNavigationSearchAndLayoutOptions()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var host = new ChatDiffView("old first\nstable\nold last", "new first\nstable\nnew last"))
            {
                var view = host.Child as CodeDiffView;
                Assert.IsNotNull(view);
                Assert.AreEqual(FlowDirection.LeftToRight, host.FlowDirection);
                Assert.IsTrue(view.UnifiedDiff);
                var grid = UiInvoke.Field<Forms.DataGridView>(view, "grid");
                Assert.IsFalse(grid.Columns[1].Visible);
                Assert.AreEqual(UiTheme.Surface, grid.BackgroundColor);
                Assert.AreEqual(UiTheme.Foreground, grid.ColumnHeadersDefaultCellStyle.ForeColor);
                var window = new Window { Content = host, Width = 640, Height = 400, Left = -10000, Top = -10000, ShowInTaskbar = false };
                try
                {
                    window.Show();
                    Assert.AreEqual(UiTheme.Surface, UiInvoke.Field<Forms.TextBox>(view, "search").BackColor);
                    Assert.AreEqual(UiTheme.Surface, grid.BackgroundColor);
                    UiInvoke.Call(typeof(CodeDiffView), "Next_Click", view, null, EventArgs.Empty);
                    Assert.IsNotNull(grid.CurrentCell);
                    UiInvoke.Field<Forms.TextBox>(view, "search").Text = "NEW LAST";
                    UiInvoke.Call(typeof(CodeDiffView), "Find_Click", view, null, EventArgs.Empty);
                    var rows = UiInvoke.Field<List<DiffRow>>(view, "visible");
                    Assert.AreEqual("new last", rows[grid.CurrentCell.RowIndex].Right);
                    view.UnifiedDiff = false;
                    Assert.IsTrue(grid.Columns[1].Visible);
                    Assert.IsFalse(UiInvoke.Field<Forms.CheckBox>(view, "collapse").Checked);
                }
                finally { window.Close(); }
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void HostDisposesDesignerViewAndHandlesEmptySources()
        {
            foreach (var source in new[] { null, "", "unchanged" })
            {
                CodeDiffView view;
                using (var host = new ChatDiffView(source, source))
                {
                    view = (CodeDiffView)host.Child;
                    var rows = UiInvoke.Field<List<DiffRow>>(view, "visible");
                    Assert.IsFalse(rows.Any(row => row.Hunk >= 0));
                    UiInvoke.Call(typeof(CodeDiffView), "Next_Click", view, null, EventArgs.Empty);
                }
                Assert.IsTrue(view.IsDisposed);
            }
        }
    }
}
