using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Reflection;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks icon parsing failures and rendering isolation for missing or clipped command symbols.</summary>
    [TestClass]
    public sealed class UiCommandIconsTests
    {
        /// <summary>Handles absent and empty SVGs, reads only their first path and rejects entity expansion.</summary>
        [TestMethod]
        public void IconResourcesValidateXmlAndReadOnlyTheFirstSupportedPath()
        {
            Assert.IsNull(UiInvoke.Call(typeof(UiCommandIcons), "ReadPath", null, new object[] { null }));
            foreach (var xml in new[] { "<svg />", "<svg><metadata>no path</metadata></svg>", "<svg><path d='M 1 2 L 3 4'/><path d='Q 1 2'/></svg>" })
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
            using (var path = (GraphicsPath)UiInvoke.Call(typeof(UiCommandIcons), "ReadPath", null, stream))
            {
                if (xml.Contains("<path")) { Assert.AreEqual(2, path.PointCount); Assert.AreEqual(new PointF(1, 2), path.PathPoints[0]); }
                else Assert.IsNull(path);
            }
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("<!DOCTYPE svg [<!ENTITY unsafe 'expanded'>]><svg>&unsafe;</svg>")))
            {
                var error = Assert.ThrowsException<TargetInvocationException>(() => UiInvoke.Call(typeof(UiCommandIcons), "ReadPath", null, stream));
                Assert.IsInstanceOfType(error.InnerException, typeof(System.Xml.XmlException));
            }
        }

        /// <summary>Rejects unsupported path commands and accepts decimal SVG coordinates regardless of the current culture.</summary>
        [TestMethod]
        public void UnsupportedSvgCommandsFailAndValidPathsRemainReusableAcrossCultures()
        {
            var previous = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var error = Assert.ThrowsException<TargetInvocationException>(() => UiInvoke.Call(typeof(UiCommandIcons), "Parse", null, "M 1 2 Q 3 4"));
                Assert.IsInstanceOfType(error.InnerException, typeof(FormatException));
                using (var path = (GraphicsPath)UiInvoke.Call(typeof(UiCommandIcons), "Parse", null, "M 1.5 2 L 3 4 C 5 6 7 8 9 10 Z"))
                using (var empty = (GraphicsPath)UiInvoke.Call(typeof(UiCommandIcons), "Parse", null, ""))
                {
                    Assert.IsTrue(path.PointCount > 0); Assert.AreEqual(1.5f, path.PathPoints[0].X);
                    Assert.AreEqual(0, empty.PointCount);
                }
            }
            finally { CultureInfo.CurrentCulture = previous; }
        }

        /// <summary>Leaves the caller's pixels and transform unchanged for missing resources and bounds too small to render.</summary>
        [TestMethod]
        public void MissingAndClippedIconsDoNotChangeCallerGraphics()
        {
            using (var image = new Bitmap(40, 40))
            using (var graphics = Graphics.FromImage(image))
            {
                graphics.Clear(Color.Magenta); graphics.TranslateTransform(2, 3);
                var before = graphics.Transform.Elements;
                Assert.IsFalse(UiCommandIcons.Draw(graphics, (UiSymbol)0xFFFF, new Rectangle(0, 0, 40, 40), Color.White, 96));
                Assert.IsTrue(UiCommandIcons.Draw(graphics, UiSymbol.Copy, new Rectangle(0, 0, 2, 2), Color.White, 96));
                Assert.AreEqual(Color.Magenta.ToArgb(), image.GetPixel(20, 20).ToArgb());
                CollectionAssert.AreEqual(before, graphics.Transform.Elements);
                Assert.IsTrue(UiCommandIcons.Draw(graphics, UiSymbol.Copy, new Rectangle(0, 0, 30, 30), Color.White, 96));
                CollectionAssert.AreEqual(before, graphics.Transform.Elements);
            }
        }
    }
}
