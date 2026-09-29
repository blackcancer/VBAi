using System;
using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie le rendu des onglets et contrôles thémés sous différentes palettes.</summary>
    [TestClass]
    public sealed class ThemedTabControlTests
    {
        /// <summary>Checks tab and close-command hover transitions without selecting or closing a document.</summary>
        [STATestMethod]
        public void HoverAndCloseHoverRepaintWithoutChangingDocumentSelection()
        {
            using (var scope = new ThemeScope())
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, Width = 450, Height = 220 })
            using (var tabs = new ThemedTabControl { Dock = DockStyle.Fill })
            using (var image = new Bitmap(450, 220))
            using (var graphics = Graphics.FromImage(image))
            {
                form.Controls.Add(tabs);
                tabs.TabPages.Add(new TabPage("First"));
                tabs.TabPages.Add(new TabPage("Second"));
                tabs.TabPages.Add(new TabPage("Disabled") { Enabled = false });
                form.Show(); tabs.SelectedIndex = 0;
                int closes = 0; tabs.CloseRequested += (sender, args) => closes++;
                foreach (var choice in new[] { ThemeChoice.Light, ThemeChoice.Dark })
                foreach (bool showClose in new[] { false, true })
                {
                    ThemeScope.SetChoice(choice); UiTheme.Apply(form); tabs.ShowCloseButtons = showClose;
                    for (int index = 0; index < tabs.TabCount; index++)
                    {
                        var tab = tabs.GetTabRect(index);
                        var close = (Rectangle)UiInvoke.Call(typeof(ThemedTabControl), "CloseBounds", tabs, index);
                        foreach (var point in new[] { new Point(tab.Left + 8, tab.Top + tab.Height / 2), new Point(close.Left + close.Width / 2, close.Top + close.Height / 2) })
                        {
                            var args = new MouseEventArgs(MouseButtons.None, 0, point.X, point.Y, 0);
                            UiInvoke.Call(typeof(ThemedTabControl), "OnMouseMove", tabs, args);
                            UiInvoke.Call(typeof(ThemedTabControl), "OnMouseMove", tabs, args);
                            Assert.AreEqual(index, UiInvoke.Field<int>(tabs, "hoveredTab"));
                            Assert.AreEqual(showClose && close.Contains(point), UiInvoke.Field<bool>(tabs, "closeHovered"));
                            Assert.AreEqual(showClose && close.Contains(point) ? Cursors.Hand : Cursors.Default, tabs.Cursor);
                            UiInvoke.Call(typeof(ThemedTabControl), "OnPaint", tabs, new PaintEventArgs(graphics, tabs.ClientRectangle));
                            Assert.AreEqual(0, tabs.SelectedIndex); Assert.AreEqual(0, closes);
                        }
                    }
                    UiInvoke.Call(typeof(ThemedTabControl), "OnMouseMove", tabs, new MouseEventArgs(MouseButtons.None, 0, -1, -1, 0));
                    Assert.AreEqual(-1, UiInvoke.Field<int>(tabs, "hoveredTab"));
                    UiInvoke.Call(typeof(ThemedTabControl), "OnMouseLeave", tabs, EventArgs.Empty);
                    Assert.IsFalse(UiInvoke.Field<bool>(tabs, "closeHovered")); Assert.AreEqual(Cursors.Default, tabs.Cursor);
                }
            }
        }

        /// <summary>Peint sélection, texte désactivé et focus clavier dans les deux palettes.</summary>
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
                Assert.AreEqual(form.BackColor.ToArgb(),bitmap.GetPixel(399,199).ToArgb());
                tabs.TabPages.Add(new TabPage("first")); tabs.TabPages.Add(new TabPage("second") {Enabled=false});
                form.Show(); tabs.Focus(); Assert.IsTrue(tabs.Focused);
                foreach(var dark in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    UiTheme.Apply(form);
                    tabs.SelectedIndex=0;
                    UiInvoke.Call(typeof(ThemedTabControl),"OnPaint",tabs,new PaintEventArgs(graphics,tabs.ClientRectangle));
                    var rect=tabs.GetTabRect(0);
                    Assert.AreEqual(UiTheme.Surface.ToArgb(), bitmap.GetPixel(rect.Left+10,rect.Top+6).ToArgb());
                    tabs.SelectedIndex=1; form.Focus();
                    UiInvoke.Call(typeof(ThemedTabControl),"OnPaint",tabs,new PaintEventArgs(graphics,tabs.ClientRectangle));
                }
            }
        }
        /// <summary>Rend boutons et flèches de liste sous contraste, direction et changement de thème.</summary>
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
                        if(dark) { int x=(rtl==RightToLeft.Yes ? 1 : combo.Width-SystemInformation.VerticalScrollBarWidth-1)+SystemInformation.VerticalScrollBarWidth/2; int y=(combo.Height-2)/2+1; Assert.AreEqual((enabled ? UiTheme.Foreground : SystemColors.GrayText).ToArgb(),bitmap.GetPixel(x,y+2).ToArgb()); }
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
        [STATestMethod]
        public void CloseButtonsRouteOnlyOwnedLeftEventsWithoutChangingTabSelection()
        {
            using (var scope = new ThemeScope())
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, Width = 400, Height = 200 })
            using (var tabs = new ThemedTabControl { Dock = DockStyle.Fill })
            {
                form.Controls.Add(tabs); form.Show(); int ordinary = 0, closes = 0;
                tabs.MouseDown += (sender, args) => ordinary++;
                Action<MouseButtons, Point> dispatch = (button, point) => UiInvoke.Call(typeof(ThemedTabControl), "OnMouseDown", tabs, new MouseEventArgs(button, 1, point.X, point.Y, 0));
                tabs.ShowCloseButtons = true; dispatch(MouseButtons.Left, Point.Empty); Assert.AreEqual(1, ordinary);
                tabs.TabPages.Add(new TabPage("First")); tabs.TabPages.Add(new TabPage("Second")); tabs.SelectedIndex = 0;
                Func<int, Point> closeCenter = index => { var rectangle = (Rectangle)UiInvoke.Call(typeof(ThemedTabControl), "CloseBounds", tabs, index); return new Point(rectangle.Left + rectangle.Width / 2, rectangle.Top + rectangle.Height / 2); };
                tabs.ShowCloseButtons = false; dispatch(MouseButtons.Left, closeCenter(0)); Assert.AreEqual(2, ordinary);
                tabs.ShowCloseButtons = true; dispatch(MouseButtons.Right, closeCenter(0)); Assert.AreEqual(3, ordinary);
                dispatch(MouseButtons.Left, new Point(tabs.ClientRectangle.Left, tabs.ClientRectangle.Bottom - 1)); Assert.AreEqual(4, ordinary);
                dispatch(MouseButtons.Left, closeCenter(0)); Assert.AreEqual(4, ordinary); Assert.AreEqual(0, tabs.SelectedIndex);
                tabs.CloseRequested += (sender, args) => { closes++; Assert.AreSame(tabs, sender); Assert.AreEqual(1, args.TabPageIndex); Assert.AreSame(tabs.TabPages[1], args.TabPage); Assert.AreEqual(TabControlAction.Deselecting, args.Action); };
                dispatch(MouseButtons.Left, closeCenter(1)); Assert.AreEqual(1, closes); Assert.AreEqual(4, ordinary); Assert.AreEqual(0, tabs.SelectedIndex); Assert.AreEqual(2, tabs.TabCount);
            }
        }
    }
}
