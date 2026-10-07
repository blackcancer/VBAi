using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatQueuedMessageViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedQueueControls() { using (var view = new ChatQueuedMessageView()) { Assert.AreEqual("ChatQueuedMessageView", view.Name); Assert.IsTrue(view.Controls.Count > 0); } }
        [STATestMethod]
        public void MessageTooltipAndQueueActionsPreserveSenderAndSubscription()
        {
            using (var view = new ChatQueuedMessageView())
            {
                var message = UiInvoke.Field<Label>(view, "message"); var tips = UiInvoke.Field<ToolTip>(view, "toolTips");
                foreach (var text in new[] { null, "", "queued\nmessage" }) { view.ShowMessage(text); Assert.AreEqual(text ?? "", message.Text); Assert.AreEqual(text ?? "", tips.GetToolTip(message)); }
                var send = UiInvoke.Field<ChatActionButton>(view, "sendNow"); var edit = UiInvoke.Field<ChatActionButton>(view, "edit"); var delete = UiInvoke.Field<ChatActionButton>(view, "delete");
                send.PerformClick(); edit.PerformClick(); delete.PerformClick();
                int sends = 0, edits = 0, deletes = 0; EventHandler sent = (s, e) => { Assert.AreSame(view, s); Assert.IsNotNull(e); sends++; };
                view.SendNowRequested += sent; view.EditRequested += (s, e) => { Assert.AreSame(view, s); edits++; }; view.DeleteRequested += (s, e) => { Assert.AreSame(view, s); deletes++; };
                send.PerformClick(); edit.PerformClick(); delete.PerformClick(); Assert.AreEqual(1, sends); Assert.AreEqual(1, edits); Assert.AreEqual(1, deletes);
                view.SendNowRequested -= sent; send.PerformClick(); Assert.AreEqual(1, sends);
            }
        }
    }
}
