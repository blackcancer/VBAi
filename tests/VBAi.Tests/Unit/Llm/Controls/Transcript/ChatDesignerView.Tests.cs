using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatDesignerViewContractTests
    {
        [STATestMethod]
        public void VerticalRowsTrackWidthOnceAndPreserveMinimumHeight()
        {
            using (var view = new ChatDesignerView())
            using (var parent = new Panel())
            using (var flow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, Width = 200, Height = 200, Padding = new Padding(5) })
            using (var row = new Panel { Width = 40, Margin = new Padding(3), MinimumSize = new Size(0, 17) })
            using (var horizontal = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight })
            {
                flow.Controls.Add(row); parent.Controls.Add(flow); parent.Controls.Add(horizontal); view.Controls.Add(parent);
                UiInvoke.Call(typeof(ChatDesignerView), "Watch", view, parent);
                flow.Width = 210; Assert.AreEqual(194, row.Width); Assert.AreEqual(new Size(194, 17), row.MinimumSize); Assert.AreEqual(new Size(194, 0), row.MaximumSize);
                int widthChanges = 0; row.SizeChanged += (s, e) => widthChanges++; flow.Height++; Assert.AreEqual(0, widthChanges);
                UiInvoke.Call(typeof(ChatDesignerView), "ResizeRows", null, flow); Assert.AreEqual(0, widthChanges);
                flow.Width = 10; Assert.AreEqual(20, row.Width); flow.Width = 0; Assert.AreEqual(20, row.Width);
                using (var second = new Panel { Width = 1, Margin = Padding.Empty }) { flow.Width = 100; flow.Controls.Add(second); Assert.AreEqual(90, second.Width); }
                using (var ordinary = new Panel()) view.Controls.Add(ordinary);
            }
        }
    }
}