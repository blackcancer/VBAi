using System;
using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class UiActionButtonTests
    {
        /// <summary>Uses Windows highlight colors for contrasted commands and preserves captions when no icon is selected.</summary>
        [STATestMethod]
        public void ContrastedCommandsAndFontFallbackKeepAccessibleRendering()
        {
            Assert.AreEqual("Segoe Fluent Icons", UiInvoke.Call(typeof(UiActionButton), "ChooseSymbolFont", null, "Segoe Fluent Icons"));
            Assert.AreEqual("Segoe MDL2 Assets", UiInvoke.Call(typeof(UiActionButton), "ChooseSymbolFont", null, "Arial"));
            using (var scope = new ThemeScope())
            using (var form = new Form { StartPosition = FormStartPosition.Manual, Location = new Point(-10000, -10000) })
            using (var button = new UiActionButton { Text = "Save", Size = new Size(180, 40) })
            using (var image = new Bitmap(180, 40))
            using (var graphics = Graphics.FromImage(image))
            {
                form.Controls.Add(button); form.Show(); button.Focus();
                foreach (bool contrast in new[] { false, true })
                foreach (bool primary in new[] { false, true })
                foreach (bool enabled in new[] { false, true })
                foreach (bool pressed in new[] { false, true })
                {
                    UiTheme.HighContrast = () => contrast;
                    button.Primary = primary; button.Enabled = enabled;
                    if (enabled) { button.Focus(); Assert.IsTrue(button.Focused); }
                    NativeUiState.SendMessage(button.Handle, 0x128, new IntPtr(0x10002), IntPtr.Zero);
                    UiInvoke.Call(typeof(UiActionButton), "OnMouseLeave", button, EventArgs.Empty);
                    UiInvoke.Call(typeof(UiActionButton), "OnMouseEnter", button, EventArgs.Empty);
                    if (pressed) UiInvoke.Call(typeof(UiActionButton), "OnMouseDown", button, new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0));
                    UiInvoke.Call(typeof(UiActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle));
                    if (contrast) Assert.AreEqual((enabled ? SystemColors.Highlight : form.BackColor).ToArgb(), image.GetPixel(15, 20).ToArgb());
                    Assert.AreEqual("Save", button.AccessibilityObject.Name);
                }
                button.IconOnly = true; button.Symbol = UiSymbol.None;
                button.Text = null; Assert.AreEqual(string.Empty, button.Text);
                button.Text = "Save";
                Assert.IsTrue(button.GetPreferredSize(Size.Empty).Width > 32);
                Assert.AreEqual("Save", button.AccessibilityObject.Name);
            }
        }

        /// <summary>Checks command interaction, caption accessibility and fallback rendering across surface and button states.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void CommandStatesRenderWithoutChangingCaptionOrPreferredSize()
        {
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false })
            using (var button = new UiActionButton { Text = "Save document", Size = new Size(160, 36) })
            using (var image = new Bitmap(160, 36))
            using (var graphics = Graphics.FromImage(image))
            {
                form.Controls.Add(button); form.Show(); button.Focus();
                foreach (var background in new[] { Color.White, Color.FromArgb(22, 26, 33) })
                foreach (bool primary in new[] { false, true })
                foreach (bool enabled in new[] { false, true })
                foreach (var symbol in new[] { UiSymbol.None, UiSymbol.Copy, (UiSymbol)'!' })
                foreach (bool iconOnly in new[] { false, true })
                foreach (var direction in new[] { RightToLeft.No, RightToLeft.Yes })
                {
                    form.BackColor = background; button.ForeColor = background == Color.White ? Color.Black : Color.White;
                    button.Primary = primary; button.Enabled = enabled; button.Symbol = symbol; button.IconOnly = iconOnly; button.RightToLeft = direction;
                    var expected = button.GetPreferredSize(Size.Empty);
                    foreach (int state in new[] { 0, 1, 2, 3 })
                    {
                        UiInvoke.Call(typeof(UiActionButton), "OnMouseLeave", button, EventArgs.Empty);
                        if (state == 1) UiInvoke.Call(typeof(UiActionButton), "OnMouseEnter", button, EventArgs.Empty);
                        if (state == 2 || state == 3) UiInvoke.Call(typeof(UiActionButton), "OnMouseDown", button, new MouseEventArgs(state == 2 ? MouseButtons.Left : MouseButtons.Right, 1, 0, 0, 0));
                        Assert.AreEqual(state == 2, UiInvoke.Field<bool>(button, "pressed"));
                        UiInvoke.Call(typeof(UiActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle));
                        Assert.AreEqual("Save document", button.AccessibilityObject.Name);
                        Assert.AreEqual(expected, button.GetPreferredSize(Size.Empty));
                        UiInvoke.Call(typeof(UiActionButton), "OnMouseUp", button, new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0));
                        Assert.IsFalse(UiInvoke.Field<bool>(button, "pressed"));
                    }
                }
                button.Size = new Size(1, 1);
                UiInvoke.Call(typeof(UiActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle));
                button.Size = new Size(160, 1);
                UiInvoke.Call(typeof(UiActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle));
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void EveryBundledSymbolRendersAtNormalAndHighDpiWithoutAFontFallback()
        {
            foreach (UiSymbol symbol in Enum.GetValues(typeof(UiSymbol)))
            {
                if (symbol == UiSymbol.None) continue;
                foreach (int dpi in new[] { 96, 192 })
                using (var image = new Bitmap(64,64))
                using (var graphics = Graphics.FromImage(image))
                {
                    graphics.Clear(Color.Black);
                    Assert.IsTrue(UiCommandIcons.Draw(graphics,symbol,new Rectangle(0,0,64,64),Color.White,dpi),symbol.ToString());
                    int pixels = 0;
                    for (int y=0;y<64;y++) for (int x=0;x<64;x++) if (image.GetPixel(x,y).R>0) pixels++;
                    Assert.IsTrue(pixels>0,symbol.ToString());
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void CaptionButtonSizeStaysStableAcrossRepeatedLayoutPasses()
        {
            using (var button = new UiActionButton { Text = "Save", Symbol = UiSymbol.Save, AutoSize = true })
            {
                var expected = button.GetPreferredSize(Size.Empty);
                for (int i = 0; i < 20; i++)
                {
                    button.Size = button.GetPreferredSize(Size.Empty);
                    Assert.AreEqual(expected, button.GetPreferredSize(Size.Empty));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void IconCaptionRemainsAccessibleAndLocalizedAfterStateChanges()
        {
            using (var button = new UiActionButton { Text = "Envoyer", Symbol = UiSymbol.Upload, IconOnly = true })
            {
                Assert.AreEqual("Envoyer", button.AccessibilityObject.Name);
                button.Text = "Arrêter"; button.Symbol = UiSymbol.Stop;
                Assert.AreEqual("Arrêter", button.AccessibilityObject.Name);
                Assert.IsTrue(button.GetPreferredSize(Size.Empty).Width >= 30);
                button.Enabled=false;
                using(var image = new Bitmap(32,30)) button.DrawToBitmap(image,new Rectangle(0,0,32,30));
            }
        }
    }
}
