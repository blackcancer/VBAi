namespace VBAi.Tests.Unit
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeWindowIconsTests
    {
        /// <summary>Exercises the same ICO-to-bitmap conversion used by Office menus against the source artwork.</summary>
        [TestMethod]
        public void SmallIcoFramesDecodeAsArtworkRatherThanPngBytes()
        {
            foreach (string name in new[] { "assistant", "github", "settings" })
            foreach (int size in new[] { 16, 20, 24, 32, 48 })
            using (var source = System.Drawing.Image.FromFile(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IconSources", name + ".png")))
            using (var icon = VbeWindowIcons.Icon(name))
            using (var small = new System.Drawing.Icon(icon, size, size))
            using (var actual = small.ToBitmap())
            using (var expected = new System.Drawing.Bitmap(size, size))
            {
                using (var graphics = System.Drawing.Graphics.FromImage(expected))
                {
                    graphics.Clear(System.Drawing.Color.Transparent);
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    double scale = size / (double)Math.Max(source.Width, source.Height);
                    int width = (int)Math.Round(source.Width * scale), height = (int)Math.Round(source.Height * scale);
                    graphics.DrawImage(source, (size - width) / 2, (size - height) / 2, width, height);
                }
                long error = 0;
                for (int y = 0; y < size; y++)
                for (int x = 0; x < size; x++)
                {
                    var a = actual.GetPixel(x, y); var e = expected.GetPixel(x, y);
                    error += Math.Abs(a.A - e.A);
                    if (e.A > 0) error += Math.Abs(a.R - e.R) + Math.Abs(a.G - e.G) + Math.Abs(a.B - e.B);
                }
                Assert.IsTrue(error / (double)(size * size * 4) < 3, name + " " + size + "px artwork differs from its source: " + error);
            }
        }

        [TestMethod]
        public void NativeEmbeddedImagesAndMissingNamesReturnOwnedCopiesOrNull()
        {
            Assert.IsNull(VbeWindowIcons.Icon("missing-coverage-resource"));
            Assert.IsNull(VbeWindowIcons.Image("missing-coverage-resource"));
            foreach (string name in new[] { "assistant", "github", "settings" })
                using (var icon = VbeWindowIcons.Icon(name)) { Assert.IsNotNull(icon); Assert.IsTrue(icon.Width > 0); }
            foreach (string name in new[] { "github", "settings" })
                using (var image = VbeWindowIcons.Image(name)) { Assert.IsNotNull(image); Assert.IsTrue(image.Width > 0); }
        }
    }
}
