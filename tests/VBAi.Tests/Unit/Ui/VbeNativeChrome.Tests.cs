using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Drawing;
using System.Reflection;

namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeChromePaintTests
    {
        private static int Rgb(System.Drawing.Color color) => color.ToArgb() & 0xffffff;
        private static void AssertSolid(System.Drawing.Bitmap bitmap, int expected)
        { for (int y = 0; y < bitmap.Height; y++) for (int x = 0; x < bitmap.Width; x++) Assert.AreEqual(expected, Rgb(bitmap.GetPixel(x, y))); }

        [TestMethod]
        public void GeometryAndDcFailuresLeaveMemoryCanvasUnchanged()
        {
            foreach (int guard in new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 })
                using (var fixture = new NativeChromePaintFixture())
                {
                    if (guard == 0) fixture.BoundsAvailable = false;
                    if (guard == 1) fixture.ClientAvailable = false;
                    if (guard == 2) fixture.OriginAvailable = false;
                    if (guard == 3) fixture.Width = 0;
                    if (guard == 4) fixture.Height = 0;
                    if (guard == 5) fixture.Width = 16389;
                    if (guard == 6) fixture.Height = 16390;
                    if (guard == 7) fixture.MissingDc = true;
                    if (guard == 8) { fixture.HasChild = true; fixture.ChildAvailable = false; }
                    using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () =>
                        VbeNativeChrome.Paint(fixture.Window, true, guard == 7 ? IntPtr.Zero : fixture.Dc, hostedCaption: guard == 8 || guard == 9)))
                        AssertSolid(bitmap, 0xffffff);
                    Assert.AreEqual(0, fixture.Releases); Assert.AreEqual(0, fixture.Allocations);
                }
        }

        [TestMethod]
        public void NormalNonclientAndHostedPaintingConvertOnlyCapturedAreaAndReleaseOwnedDc()
        {
            using (var fixture = new NativeChromePaintFixture())
            {
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc)))
                {
                    Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(0, 0))); Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(19, 10)));
                    Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(20, 11)));
                }
                Assert.AreEqual(0, fixture.Releases); Assert.AreEqual(0, fixture.ClientAcquires);
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, false, IntPtr.Zero)))
                { Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(23, 2))); Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(0, 3))); }
                Assert.AreEqual(1, fixture.Releases); Assert.AreEqual(1, fixture.WindowAcquires);
                fixture.HasChild = true;
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, false, fixture.Dc, true)))
                { Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(23, 5))); Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(0, 6))); }
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, IntPtr.Zero)))
                    Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(0, 0)));
                Assert.AreEqual(2, fixture.Releases); Assert.AreEqual(1, fixture.ClientAcquires);
            }
        }

        [TestMethod]
        public void CodeAndPreservedClientPoliciesDistinguishDarkAndLightRaster()
        {
            using (var fixture = new NativeChromePaintFixture())
            {
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc, codeSurface: true))) AssertSolid(bitmap, 0xffffff);
                using (var bitmap = fixture.Canvas(System.Drawing.Color.Black, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc, preserveDarkClient: true))) AssertSolid(bitmap, 0);
                using (var bitmap = fixture.Canvas(System.Drawing.Color.Black, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc, preserveDarkClient: true, codeSurface: true)))
                { Assert.AreEqual(0x282d35, Rgb(bitmap.GetPixel(0, 0))); Assert.AreEqual(0, Rgb(bitmap.GetPixel(21, 12))); }
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc, preserveDarkClient: true)))
                    Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(0, 0)));
                using (var bitmap = fixture.Canvas(System.Drawing.Color.FromArgb(32, 36, 43), () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc))) AssertSolid(bitmap, 0x20242b);
            }
        }

        [TestMethod]
        public void ReentrantPaintAndAllocationFailureDoNotLeakPaintingStateOrOwnedDc()
        {
            using (var fixture = new NativeChromePaintFixture())
            {
                fixture.Reenter = true;
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc))) Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(0, 0)));
                Assert.AreEqual(1, fixture.Allocations);
                fixture.Reenter = false; fixture.AllocationFailure = new OutOfMemoryException("synthetic chrome bitmap allocation failure");
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, IntPtr.Zero))) AssertSolid(bitmap, 0xffffff);
                Assert.AreEqual(1, fixture.Releases); StringAssert.Contains(fixture.Messages[0], "Native chrome painting failed: synthetic chrome bitmap allocation failure");
                fixture.AllocationFailure = null;
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc))) Assert.AreEqual(0x20242b, Rgb(bitmap.GetPixel(0, 0)));
                Assert.AreEqual(3, fixture.Allocations);
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.Paint(fixture.Window, true, new IntPtr(1)))) AssertSolid(bitmap, 0xffffff);
                Assert.AreEqual(4, fixture.Allocations);
            }
        }

        [TestMethod]
        public void BordersPaintEachIndependentEdgeWithoutOverwritingTheClient()
        {
            foreach (int edge in new[] { 0, 1, 2, 3 })
                using (var fixture = new NativeChromePaintFixture())
                {
                    fixture.LeftEdge = edge == 0 || edge == 3 ? 2 : 0; fixture.RightEdge = edge == 1 || edge == 3 ? 2 : 0;
                    fixture.BottomEdge = edge == 2 || edge == 3 ? 2 : 0; fixture.Caption = edge == 3 ? 3 : 0;
                    using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintBorder(fixture.Window)))
                    {
                        Assert.AreEqual(0x3e4651, Rgb(bitmap.GetPixel(edge == 0 ? 0 : edge == 1 ? 23 : 8, edge == 2 ? 15 : edge == 3 ? 0 : 8)));
                        Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(8, 8)));
                    }
                    Assert.AreEqual(1, fixture.Releases);
                }
        }

        [TestMethod]
        public void InvalidBorderGeometryMissingDcAndGraphicsFailurePreserveCanvas()
        {
            foreach (int guard in new[] { 0, 1, 2, 3, 4, 5, 6 })
                using (var fixture = new NativeChromePaintFixture())
                {
                    if (guard == 0) fixture.BoundsAvailable = false;
                    if (guard == 1) fixture.ClientAvailable = false;
                    if (guard == 2) fixture.OriginAvailable = false;
                    if (guard == 3) fixture.Width = 0;
                    if (guard == 4) fixture.Height = 0;
                    if (guard == 5) fixture.LeftEdge = fixture.RightEdge = fixture.BottomEdge = 0;
                    if (guard == 6) fixture.MissingDc = true;
                    using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintBorder(fixture.Window))) AssertSolid(bitmap, 0xffffff);
                    Assert.AreEqual(0, fixture.Releases);
                }
            using (var fixture = new NativeChromePaintFixture())
            {
                VbeNativeChrome.CreateGraphicsFromDc = dc => { throw new OutOfMemoryException("synthetic border graphics allocation"); };
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintBorder(fixture.Window))) AssertSolid(bitmap, 0xffffff);
                Assert.AreEqual(1, fixture.Releases); StringAssert.Contains(fixture.Messages[0], "Native border painting failed: synthetic border graphics allocation");
            }
        }

        [TestMethod]
        public void ComboStatesRenderExactFaceAndArrowWithoutMovingTheCursor()
        {
            foreach (int state in new[] { 0, 1, 2, 3, 4, 5, 6 })
                using (var fixture = new NativeChromePaintFixture())
                {
                    fixture.LeftEdge = fixture.RightEdge = fixture.BottomEdge = fixture.Caption = 0;
                    if (state == 0) fixture.Enabled = false;
                    if (state == 1) fixture.ButtonState = 8;
                    if (state == 3) fixture.CursorAvailable = false;
                    if (state == 4) fixture.CursorConversion = false;
                    if (state == 5) fixture.PointerX = -1;
                    if (state == 6) { fixture.Width = 6; fixture.Height = 6; fixture.ButtonLeft = fixture.ButtonTop = 1; fixture.ButtonRight = fixture.ButtonBottom = 5; }
                    using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintComboButton(fixture.Window)))
                    {
                        if (state == 6) Assert.AreEqual(0x3e4651, Rgb(bitmap.GetPixel(0, 0)));
                        else
                        {
                            Assert.AreEqual(state == 0 ? 0x20242b : state == 1 ? 0x344452 : state == 2 ? 0x3e4651 : 0x282d35, Rgb(bitmap.GetPixel(14, 4)));
                            Assert.AreEqual(state == 0 ? 0x78808b : 0xe2e8f0, Rgb(bitmap.GetPixel(18, 8)));
                            Assert.AreEqual(0xffffff, Rgb(bitmap.GetPixel(4, 4)));
                        }
                    }
                    Assert.AreEqual(1, fixture.Releases);
                }
        }

        [TestMethod]
        public void InvalidComboStateAndGraphicsFailurePreserveCanvasAndOwnership()
        {
            foreach (int guard in new[] { 0, 1, 2, 3, 4, 5, 6 })
                using (var fixture = new NativeChromePaintFixture())
                {
                    fixture.LeftEdge = fixture.RightEdge = fixture.BottomEdge = fixture.Caption = 0;
                    if (guard == 0) fixture.ComboAvailable = false;
                    if (guard == 1) fixture.ClientAvailable = false;
                    if (guard == 2) fixture.ButtonState = 0x8000;
                    if (guard == 3) fixture.ButtonRight = fixture.ButtonLeft + 3;
                    if (guard == 4) fixture.ButtonBottom = fixture.ButtonTop + 3;
                    if (guard == 5) fixture.ButtonRight = fixture.Width + 1;
                    if (guard == 6) fixture.MissingDc = true;
                    using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintComboButton(fixture.Window))) AssertSolid(bitmap, 0xffffff);
                    Assert.AreEqual(0, fixture.Releases);
                }
            using (var fixture = new NativeChromePaintFixture())
            {
                fixture.LeftEdge = fixture.RightEdge = fixture.BottomEdge = fixture.Caption = 0;
                VbeNativeChrome.CreateGraphicsFromDc = dc => { throw new OutOfMemoryException("synthetic combo graphics allocation"); };
                using (var bitmap = fixture.Canvas(System.Drawing.Color.White, () => VbeNativeChrome.PaintComboButton(fixture.Window))) AssertSolid(bitmap, 0xffffff);
                Assert.AreEqual(1, fixture.Releases); StringAssert.Contains(fixture.Messages[0], "Native combo button painting failed: synthetic combo graphics allocation");
            }
        }
        [TestMethod]
        public void NativeCodeMarginUsesChromeMappingWhileTextAreaUsesCodeMapping()
        {
            using (var fixture = new NativeChromePaintFixture())
            using (var bitmap = fixture.Canvas(System.Drawing.Color.Black, () =>
            {
                using (var graphics = System.Drawing.Graphics.FromHdc(fixture.Dc))
                using (var face = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(240, 240, 240)))
                    graphics.FillRectangle(face, 0, 0, 3, 16);
                VbeNativeChrome.Paint(fixture.Window, true, fixture.Dc, codeSurface: true);
            }))
            {
                for (int y = 0; y < 16; y++)
                    for (int x = 0; x < 24; x++)
                    {
                        int expected = x < 20 && y < 11 ? x < 3 ? 0x2b3037 : 0x282d35 : x < 3 ? 0xf0f0f0 : 0;
                        Assert.AreEqual(expected, Rgb(bitmap.GetPixel(x, y)));
                    }
            }
        }
    }
}
