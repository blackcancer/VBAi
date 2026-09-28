using System;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
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
                    Assert.AreEqual((int)size.Width,card.Width); Assert.AreEqual((int)size.Width,card.MinimumSize.Width); Assert.AreEqual((int)size.Width,card.MaximumSize.Width);
                }
                card.MaximumSize=System.Drawing.Size.Empty;
                UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(card.Width,100)); Assert.AreEqual(card.Width,card.MaximumSize.Width);
                card.Dispose(); Assert.AreEqual(new System.Windows.Size(),UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)));
                host.Child=null; Assert.IsNull(host.View); Assert.AreEqual(new System.Windows.Size(),UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)));
            }
            using(var card=new PreferredHeightView())
            using(var host=new ChatDesignerHost(card)) {
                foreach(var height in new[]{0,40000}) { card.HeightWanted=height; var size=(System.Windows.Size)UiInvoke.Call(typeof(ChatDesignerHost),"MeasureOverride",host,new System.Windows.Size(200,100)); Assert.AreEqual(height==0?24:32000,size.Height); }
            }
        }
    }
}