using System;
using System.Drawing;
using System.Reflection;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeChromeTests
    {
        private static int Rgb(Color value) => value.ToArgb() & 0xffffff;
        private static T Evaluate<T>(string method, Bitmap bitmap) => (T)typeof(VbeNativeChrome)
            .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new object[] { bitmap });

        [TestMethod]
        public void ChromePaletteMappingHasExactColorsAndPreservesAlpha()
        {
            int[] inputs = { 0xffffff, 0, 0xff0000, 0x008000, 0x00ff00, 0xffff00, 0x0000ff, 0x000080, 0xb4c0ce };
            int[] outputs = { 0x20242b, 0xe2e8f0, 0xf48771, 0x6a9955, 0x6a9955, 0xdcdcaa, 0x569cd6, 0x569cd6, 0x344452 };
            foreach (int alpha in new[] { 0, unchecked((int)0x80000000), unchecked((int)0xff000000) })
                for (int index = 0; index < inputs.Length; index++)
                {
                    Assert.AreEqual(alpha | outputs[index], VbeNativeChrome.MapPixel(alpha | inputs[index]));
                    Assert.AreEqual(alpha | outputs[index], VbeNativeChrome.MapPixel(alpha | outputs[index]));
                }
            // Weighted intensity of (100, 50, 0) is 60; neutral curve is (180,186,194).
            Assert.AreEqual(0xb4bac2, VbeNativeChrome.MapPixel(0x643200));
        }

        [TestMethod]
        public void CodePaletteMappingPreservesMarkersAndDistinguishesFaceBorderAndRamps()
        {
            int[] inputs = { 0xf0f0f0, 0xe3e3e3, 0, 0xc0c0c0, 0x008000, 0x008080,
                0x800000, 0x000080, 0x800080, 0x808000, 0xff0000, 0xffff00, 0xffffff };
            int[] outputs = { 0x282d35, 0x3e4651, 0x282d35, 0xdcdcdc, 0x57a64a, 0x569cd6,
                0x800000, 0x000080, 0x800080, 0x808000, 0xff0000, 0xffff00, 0xffffff };
            for (int index = 0; index < inputs.Length; index++)
            {
                int alpha = unchecked((int)0xab000000);
                Assert.AreEqual(alpha | outputs[index], VbeNativeChrome.MapCodePixel(alpha | inputs[index]));
                Assert.AreEqual(alpha | outputs[index], VbeNativeChrome.MapCodePixel(alpha | outputs[index]));
            }
        }

        [TestMethod]
        public void DarkBackgroundRequiresAllSamplesAndAllChannelsBelow64()
        {
            using (var bitmap = new Bitmap(9, 9))
            {
                using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.FromArgb(63, 63, 63));
                Assert.IsTrue(Evaluate<bool>("HasDarkBackground", bitmap));
                bitmap.SetPixel(0, 0, Color.White);
                Assert.IsTrue(Evaluate<bool>("HasDarkBackground", bitmap), "Pixels outside the sample positions do not determine editor background.");
                foreach (Color channel in new[] { Color.FromArgb(64, 63, 63), Color.FromArgb(63, 64, 63), Color.FromArgb(63, 63, 64) })
                {
                    bitmap.SetPixel(3, 3, channel);
                    Assert.IsFalse(Evaluate<bool>("HasDarkBackground", bitmap));
                }
            }
        }

        [TestMethod]
        public void CodeMarginUsesLastExactFacePixelWithinBothScanLimits()
        {
            using (var tiny = new Bitmap(3, 2)) Assert.AreEqual(0, Evaluate<int>("LightCodeMargin", tiny));
            using (var bitmap = new Bitmap(80, 8))
            {
                Assert.AreEqual(0, Evaluate<int>("LightCodeMargin", bitmap));
                bitmap.SetPixel(2, 4, Color.FromArgb(240, 240, 240));
                bitmap.SetPixel(19, 4, Color.FromArgb(240, 240, 240));
                bitmap.SetPixel(20, 4, Color.FromArgb(240, 240, 240));
                Assert.AreEqual(20, Evaluate<int>("LightCodeMargin", bitmap));
                bitmap.SetPixel(19, 4, Color.FromArgb(239, 240, 240));
                Assert.AreEqual(3, Evaluate<int>("LightCodeMargin", bitmap));
            }
            using (var bitmap = new Bitmap(240, 8))
            {
                bitmap.SetPixel(47, 4, Color.FromArgb(240, 240, 240));
                bitmap.SetPixel(48, 4, Color.FromArgb(240, 240, 240));
                Assert.AreEqual(48, Evaluate<int>("LightCodeMargin", bitmap));
            }
        }

        [TestMethod]
        public void PropertyRowConvertsOnlyRequestedRegionAndRemainsStable()
        {
            var bounds = new VbeNativeTheme.NativeRect { Left = 2, Top = 3, Right = 11, Bottom = 8 };
            using (var bitmap = NativeChromeCanvas.Paint(Color.White, dc =>
            {
                VbeNativeChrome.PaintPropertyRow(dc, bounds);
                VbeNativeChrome.PaintPropertyRow(dc, bounds);
            }))
                for (int y = 0; y < bitmap.Height; y++)
                    for (int x = 0; x < bitmap.Width; x++)
                        Assert.AreEqual(x >= 2 && x < 11 && y >= 3 && y < 8 ? 0x20242b : 0xffffff, Rgb(bitmap.GetPixel(x, y)));
        }

        [TestMethod]
        public void InvalidPropertyBoundsAndMissingWindowLeaveCanvasUnchanged()
        {
            using (var bitmap = NativeChromeCanvas.Paint(Color.Red, dc =>
            {
                foreach (var bounds in new[] {
                    new VbeNativeTheme.NativeRect { Right = 0, Bottom = 4 },
                    new VbeNativeTheme.NativeRect { Right = 4, Bottom = 0 },
                    new VbeNativeTheme.NativeRect { Right = -1, Bottom = 4 },
                    new VbeNativeTheme.NativeRect { Right = 16385, Bottom = 4 },
                    new VbeNativeTheme.NativeRect { Right = 4, Bottom = 2049 } })
                    VbeNativeChrome.PaintPropertyRow(dc, bounds);
                VbeNativeChrome.PaintPropertyRow(IntPtr.Zero, new VbeNativeTheme.NativeRect { Right = 4, Bottom = 4 });
                VbeNativeChrome.Paint(IntPtr.Zero, true, dc);
                VbeNativeChrome.PaintBorder(IntPtr.Zero);
                VbeNativeChrome.PaintComboButton(IntPtr.Zero);
            }))
                for (int y = 0; y < bitmap.Height; y++)
                    for (int x = 0; x < bitmap.Width; x++) Assert.AreEqual(0xff0000, Rgb(bitmap.GetPixel(x, y)));
        }
        [TestMethod]
        public void PropertyRowAllocationFailureIsLoggedAndLeavesMemoryCanvasUnchanged()
        {
            var previousFactory = VbeNativeChrome.CreatePropertyRowBitmap;
            var previousLog = LoadLog.AppendText;
            var messages = new System.Collections.Generic.List<string>();
            int allocations = 0;
            try
            {
                VbeNativeChrome.CreatePropertyRowBitmap = (width, height) =>
                {
                    Assert.AreEqual(4, width); Assert.AreEqual(5, height); allocations++;
                    throw new OutOfMemoryException("synthetic property row allocation failed");
                };
                LoadLog.AppendText = (path, message) => messages.Add(message);
                using (var bitmap = NativeChromeCanvas.Paint(Color.White, dc =>
                    VbeNativeChrome.PaintPropertyRow(dc, new VbeNativeTheme.NativeRect { Right = 4, Bottom = 5 })))
                    for (int y = 0; y < bitmap.Height; y++)
                        for (int x = 0; x < bitmap.Width; x++) Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(x, y)));
                Assert.AreEqual(1, allocations); Assert.AreEqual(1, messages.Count);
                StringAssert.Contains(messages[0], "Native property row painting failed: synthetic property row allocation failed");
            }
            finally { VbeNativeChrome.CreatePropertyRowBitmap = previousFactory; LoadLog.AppendText = previousLog; }
        }
    }
}
