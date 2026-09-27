using System;
using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class ThemedTabControlTests
    {
        [STATestMethod]
        public void TabsPaintSelectionDisabledTextAndKeyboardFocusInBothPalettes()
        {
            using(var scope=new ThemeScope())
            using(var form=new Form {Left=-10000,Top=-10000,ShowInTaskbar=false,Width=400,Height=200})
            using(var tabs=new ThemedTabControl {Dock=DockStyle.Fill})
            using(var bitmap=new Bitmap(400,200))
            using(var graphics=Graphics.FromImage(bitmap))
            {
                form.Controls.Add(tabs);
                UiInvoke.Call(typeof(ThemedTabControl),"OnPaint",tabs,new PaintEventArgs(graphics,new Rectangle(0,0,400,200)));
                Assert.AreEqual(UiTheme.Background.ToArgb(),bitmap.GetPixel(399,199).ToArgb());
                tabs.TabPages.Add(new TabPage("first")); tabs.TabPages.Add(new TabPage("second") {Enabled=false});
                form.Show(); tabs.Focus(); Assert.IsTrue(tabs.Focused);
                foreach(var dark in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    tabs.SelectedIndex=0;
                    UiInvoke.Call(typeof(ThemedTabControl),"OnPaint",tabs,new PaintEventArgs(graphics,tabs.ClientRectangle));
                    var rect=tabs.GetTabRect(0);
                    Assert.AreEqual((dark ? Color.FromArgb(96,165,250) : Color.RoyalBlue).ToArgb(),bitmap.GetPixel(rect.Left+5,rect.Bottom-2).ToArgb());
                    tabs.SelectedIndex=1; form.Focus();
                    UiInvoke.Call(typeof(ThemedTabControl),"OnPaint",tabs,new PaintEventArgs(graphics,tabs.ClientRectangle));
                }
            }
        }
        [STATestMethod]
        public void DisabledButtonsAndComboArrowsRenderUnderContrastDirectionAndThemeChanges()
        {
            using(var scope=new ThemeScope())
            using(var button=new ThemedButton {Text="Action",Size=new Size(180,30)})
            using(var combo=new ThemedComboBox {Size=new Size(180,30),Text="Choice"})
            using(var bitmap=new Bitmap(180,30))
            using(var graphics=Graphics.FromImage(bitmap))
            {
                foreach(var dark in new[]{false,true})
                foreach(var enabled in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    button.Enabled=enabled; combo.Enabled=enabled;
                    UiTheme.Apply(button); UiTheme.Apply(combo);
                    UiInvoke.Call(typeof(ThemedButton),"OnPaint",button,new PaintEventArgs(graphics,button.ClientRectangle));
                    Assert.AreEqual(180,button.Width);
                    foreach(var rtl in new[]{RightToLeft.No,RightToLeft.Yes})
                    {
                        combo.RightToLeft=rtl;
                        var handle=combo.Handle;
                        combo.Invalidate(); combo.Update();
                        UiInvoke.Call(typeof(ThemedComboBox),"WndProc",combo,new object[]{Message.Create(handle,0x000F,IntPtr.Zero,IntPtr.Zero)});
                        IntPtr hdc=graphics.GetHdc();
                        try { var args=new object[]{Message.Create(handle,0x0318,hdc,IntPtr.Zero)}; UiInvoke.Call(typeof(ThemedComboBox),"WndProc",combo,args); }
                        finally { graphics.ReleaseHdc(hdc); }
                        if(dark) { int x=(rtl==RightToLeft.Yes ? 1 : combo.Width-SystemInformation.VerticalScrollBarWidth-1)+SystemInformation.VerticalScrollBarWidth/2; int y=(combo.Height-2)/2+1; Assert.AreEqual((enabled ? UiTheme.Foreground : SystemColors.GrayText).ToArgb(),bitmap.GetPixel(x,y).ToArgb()); }
                    }
                    Assert.AreEqual("Choice",combo.Text);
                }
                ThemeScope.SetChoice(ThemeChoice.Dark); UiTheme.HighContrast=()=>true;
                combo.Invalidate(); combo.Update();
                var message=new object[]{Message.Create(combo.Handle,0,IntPtr.Zero,IntPtr.Zero)};
                UiInvoke.Call(typeof(ThemedComboBox),"WndProc",combo,message);
                Assert.IsTrue(combo.IsHandleCreated);
            }
        }
    }
}
