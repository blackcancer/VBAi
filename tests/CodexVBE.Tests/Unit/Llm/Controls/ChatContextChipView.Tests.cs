using System;
using System.Reflection;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContextChipViewTests
    {
        private static void Click(Control control) => typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { EventArgs.Empty });
        [STATestMethod]
        public void DesignerCommandsRouteOpeningAndRemovalWithoutRebuildingControls()
        {
            using (var view = new ChatContextChipView())
            {
                int opened = 0, removed = 0;
                view.OpenRequested += (s, e) => opened++;
                view.RemoveRequested += (s, e) => removed++;
                var open = view.Controls.Find("open", true)[0]; var remove = view.Controls.Find("remove", true)[0];
                view.ShowItem("#Project.Module", true, "Open", "Remove");
                Click(open); Click(remove);
                Assert.AreEqual(1, opened); Assert.AreEqual(1, removed);
                view.ShowItem("Selected code ×", false, null, "Remove selection");
                Click(open);
                Assert.AreEqual(1, opened); Assert.AreEqual(2, removed);
                Assert.AreSame(open, view.Controls.Find("open", true)[0]);
            }
        }
    }
}
