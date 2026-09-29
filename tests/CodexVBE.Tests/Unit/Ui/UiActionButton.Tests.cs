using System;
using System.Drawing;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class UiActionButtonTests
    {
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
