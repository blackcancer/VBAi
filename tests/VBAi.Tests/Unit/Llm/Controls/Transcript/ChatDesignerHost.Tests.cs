using System;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass,TestCategory("Unit")]
    public sealed class ChatDesignerHostContractTests
    {
        [STATestMethod]
        public void AdapterBoundsWidthHeightAndRejectsMissingOrDisposedChildren()
        {
            using(var card=new ChatMessageView())
            using(var host=new ChatDesignerHost(card)) {
                card.message.ShowPlain(new string('a',2000)); Assert.AreSame(card,host.View);
                foreach(var width in new[]{double.PositiveInfinity,420d,420d,10d}) {
                    var size=(System.Windows.Size)UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(width,double.PositiveInfinity));
                    Assert.AreEqual(double.IsInfinity(width)?500:Math.Max(40,width),size.Width); Assert.IsTrue(size.Height>=24&&size.Height<=32000);
                    Assert.AreEqual(0,card.MaximumSize.Width, "Measurement must not pin a native pixel width to a WPF DIP value.");
                }
                card.Dispose(); Assert.AreEqual(new System.Windows.Size(),UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)));
                host.Child=null; Assert.IsNull(host.View); Assert.AreEqual(new System.Windows.Size(),UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)));
            }
            using(var card=new PreferredHeightView())
            using(var host=new ChatDesignerHost(card)) {
                foreach(var height in new[]{0,40000}) { card.HeightWanted=height; var size=(System.Windows.Size)UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)); Assert.AreEqual(height==0?24:32000,size.Height); }
            }
        }

        [STATestMethod]
        public void TransformedHostKeepsTheFrameworkPixelConversionContract()
        {
            using (var card = new PreferredHeightView { HeightWanted = 120 })
            using (var referenceCard = new PreferredHeightView { HeightWanted = 120 })
            using (var host = new ChatDesignerHost(card))
            using (var reference = new System.Windows.Forms.Integration.WindowsFormsHost { Child = referenceCard })
            {
                var panel = new System.Windows.Controls.StackPanel();
                panel.Children.Add(host); panel.Children.Add(reference);
                var window = new System.Windows.Window { Content = panel, Width = 700, Height = 900, ShowInTaskbar = false };
                try
                {
                    window.Show();
                    foreach (double scale in new[] { 1d, 1.5d, 2d })
                    {
                        host.LayoutTransform = new System.Windows.Media.ScaleTransform(scale, scale);
                        reference.LayoutTransform = new System.Windows.Media.ScaleTransform(scale, scale);
                        host.Width = reference.Width = 240;
                        window.UpdateLayout(); Application.DoEvents(); window.UpdateLayout();
                        Assert.AreEqual(referenceCard.Width, card.Width, "Native width must use the same transform/DPI conversion as WindowsFormsHost.");
                        Assert.AreEqual(reference.DesiredSize.Height, host.DesiredSize.Height, 1d);
                    }
                }
                finally { window.Close(); }
            }
        }
    }
}
