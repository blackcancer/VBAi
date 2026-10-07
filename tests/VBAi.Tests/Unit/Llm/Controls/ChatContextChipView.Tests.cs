using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContextChipViewTests
    {
        private static void Click(Control control) => typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { EventArgs.Empty });
        /// <summary>Les commandes sans abonnés restent inertes et les ressources du Designer sont libérées.</summary>
        [STATestMethod]
        public void EmptyContextAndCommandsWithoutSubscribersRemainSafe()
        {
            using (var view = new ChatContextChipView())
            {
                var open = view.Controls.Find("open", true)[0]; var remove = view.Controls.Find("remove", true)[0];
                view.ShowItem(null, true, null, null); Assert.AreEqual("", open.Text); Click(open); Click(remove);
                view.ShowItem(null, false, null, null); Click(open);
                UiInvoke.Call(typeof(ChatContextChipView), "Dispose", view, false); Assert.IsFalse(view.IsDisposed);
                var components = UiInvoke.Field<System.ComponentModel.IContainer>(view, "components"); components.Dispose();
                typeof(ChatContextChipView).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(view, null);
            }
        }
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
