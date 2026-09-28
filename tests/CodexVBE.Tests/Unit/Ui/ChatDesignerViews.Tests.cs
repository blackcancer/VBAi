using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Linq;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatDesignerViewsTests
    {
        [STATestMethod]
        public void EveryChatTemplateLoadsInAWinFormsDesignSurfaceWithoutHostOrProvider()
        {
            var types = new[] { typeof(ChatMessageView), typeof(ChatActivityGroupView), typeof(ChatActivityStepView), typeof(ChatChangeCardView), typeof(ChatAttachmentView), typeof(ChatFormRecoveryView), typeof(ChatWelcomeView), typeof(ChatSuggestionsView), typeof(ChatTextContentView), typeof(ChatLinkView), typeof(ChatDisclosureView), typeof(ChatQueuedMessageView) };
            foreach (var type in types)
                using (var surface = new DesignSurface(type)) {
                    Assert.IsTrue(surface.IsLoaded,type.Name);
                    Assert.AreEqual(0,surface.LoadErrors.Count,type.Name);
                    var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                    var root = (UserControl)host.RootComponent;
                    Assert.AreEqual(type,root.GetType()); Assert.IsTrue(root.Controls.Count>0,type.Name);
                    root.Width = 420; root.PerformLayout();
                    Assert.IsNotNull(host.GetDesigner(root),type.Name);
                }
        }
        [STATestMethod]
        public void TranscriptAdapterHasBoundedWidthAndDisposesNestedNativeControls()
        {
            var card = new ChatMessageView(); card.message.ShowPlain(new string('a',2000));
            using (var host = new ChatDesignerHost(card)) {
                host.Measure(new System.Windows.Size(420,double.PositiveInfinity));
                Assert.IsTrue(host.DesiredSize.Width<=420); Assert.IsTrue(host.DesiredSize.Height<32001);
                host.Measure(new System.Windows.Size(260,double.PositiveInfinity)); Assert.IsTrue(host.DesiredSize.Width<=260);
            }
            Assert.IsTrue(card.IsDisposed); Assert.IsTrue(card.message.content.IsDisposed);
        }
        [STATestMethod]
        public void LoadedTranscriptCardDisplaysNativeMessageAndExpandedDiff()
        {
            var card = new ChatChangeCardView(); card.diff.ShowDiff("old", "new"); card.section.Expanded = true;
            var message = new ChatMessageView(); message.message.ShowPlain("A visible native message");
            var reference = new ChatLinkView(); reference.link.Text = "@Budget.Module.Calculer"; message.references.Controls.Add(reference);
            using (var first = new ChatDesignerHost(message))
            using (var second = new ChatDesignerHost(card)) {
                var stack = new System.Windows.Controls.StackPanel(); stack.Children.Add(first); stack.Children.Add(second);
                var window = new System.Windows.Window { Content = stack, Width = 600, Height = 900, ShowInTaskbar = false };
                try {
                    window.Show(); window.UpdateLayout(); Application.DoEvents(); window.UpdateLayout();
                    Assert.IsTrue(message.Visible); Assert.IsTrue(message.Height > 40, "Message height " + message.Height);
                    Assert.IsTrue(message.message.Height >= 24, "Message body height " + message.message.Height);
                    Assert.IsTrue(message.message.content.Height >= 24);
                    Assert.IsTrue(message.message.content.Visible); Assert.IsTrue(message.message.content.IsHandleCreated);
                    Assert.IsTrue(message.headingActions.ClientRectangle.Contains(message.fork.Bounds), "Fork action clipped");
                    Assert.IsTrue(reference.Width > 50, "Reference width " + reference.Width);
                    Assert.IsTrue(reference.Visible); Assert.IsTrue(reference.Height > 20, "Reference height " + reference.Height);
                    Assert.IsTrue(card.diff.Visible, "Diff not visible");
                    Assert.IsTrue(card.diff.Height > 100, "Diff height " + card.diff.Height);
                    Assert.IsTrue(card.section.Height > 100, "Section height " + card.section.Height);
                } finally { window.Close(); }
            }
        }
        [STATestMethod]
        public void DisclosureBodyAcceptsDesignerComponents()
        {
            using (var surface = new DesignSurface(typeof(UserControl))) {
                var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                var disclosure = (ChatDisclosureView)host.CreateComponent(typeof(ChatDisclosureView));
                ((UserControl)host.RootComponent).Controls.Add(disclosure);
                Assert.IsInstanceOfType(host.GetDesigner(disclosure), typeof(ChatDisclosureDesigner));
                Assert.IsNotNull(disclosure.ContentPanel.Site, "Nested body must be available to the Designer");
                var content = (Button)host.CreateComponent(typeof(Button)); disclosure.ContentPanel.Controls.Add(content);
                Assert.AreSame(disclosure.ContentPanel, content.Parent); Assert.IsNotNull(host.GetDesigner(content));
            }
        }
    }
}
