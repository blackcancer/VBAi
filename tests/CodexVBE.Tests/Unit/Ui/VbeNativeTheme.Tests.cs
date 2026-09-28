using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativeThemeTests
    {
        [TestMethod]
        public void NativeWindowsVersionIsReadBeforeUsingVersionDependentOrdinals()
        {
            Version actual = VbeNativeTheme.ReadNativeWindowsVersion();
            Assert.IsTrue(actual.Major > 0 && actual.Build > 0, "The native version structure must be populated.");
            Assert.IsFalse(VbeNativeTheme.SupportsPreferredAppMode(new Version(10, 0, 17763)));
            Assert.IsTrue(VbeNativeTheme.SupportsPreferredAppMode(new Version(10, 0, 18362)));
            Assert.IsTrue(VbeNativeTheme.SupportsPreferredAppMode(new Version(10, 0, 22000)));
            Assert.IsFalse(VbeNativeTheme.SupportsPreferredAppMode(new Version(11, 0, 1)));
        }

        [TestMethod]
        public void ExperimentRequiresExactExplicitEnvironmentValue()
        {
            string previous = Environment.GetEnvironmentVariable(VbeNativeTheme.ExperimentVariable);
            try
            {
                foreach (string value in new[] { null, "", "0", "true", " 1 " })
                {
                    Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, value);
                    Assert.IsFalse(VbeNativeTheme.ExperimentEnabled());
                }
                Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, "1");
                Assert.IsTrue(VbeNativeTheme.ExperimentEnabled());
            }
            finally { Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, previous); }
        }

        [TestMethod]
        public void WindowClassPolicySeparatesManagedAddInAndNativeVbeSurfaces()
        {
            Assert.IsTrue(VbeNativeTheme.IsManagedAddInWindow("WindowsForms10.Window.8.app.0.fixture"));
            Assert.IsTrue(VbeNativeTheme.IsManagedAddInWindow("HwndWrapper[DefaultDomain;;fixture]"));
            Assert.IsTrue(VbeNativeTheme.IsManagedAddInWindow("GenericPane"));
            Assert.IsFalse(VbeNativeTheme.IsManagedAddInWindow("VbaWindow"));
            foreach (string className in new[] { "#32770", "wndclass_desked_gsk", "MDIClient", "VbaWindow", "PROJECT", "wndclass_pbrs", "ToolsPalette", "GenericPane", "MsoCommandBar", "MsoCommandBarPopup", "MsoCommandBarDock", "SysTabControl32", "VBSlider", "SysTreeView32", "ListBox", "ComboBox", "Edit" })
                Assert.IsTrue(VbeNativeTheme.ShouldSubclass(className));
            foreach (string className in new[] { null, "", "WindowsForms10.Window.8.app.0.fixture" })
                Assert.IsFalse(VbeNativeTheme.ShouldSubclass(className));
        }

        [TestMethod]
        public void ChromeColorsRemainStableAcrossRepeatedPaints()
        {
            for (int r = 0; r < 256; r += 17)
                for (int g = 0; g < 256; g += 17)
                    for (int b = 0; b < 256; b += 17)
                    {
                        int mapped = VbeNativeChrome.MapPixel((r << 16) | (g << 8) | b);
                        Assert.AreEqual(mapped, VbeNativeChrome.MapPixel(mapped));
                    }
            Assert.AreEqual(0x20242b, VbeNativeChrome.MapPixel(0xffffff));
            Assert.AreEqual(0xe2e8f0, VbeNativeChrome.MapPixel(0x000000));
        }

        [TestMethod]
        public void CodeEdgesRemainStableAndIncreaseSmoothlyFromTheBackground()
        {
            foreach (int source in new[] { 0xc0c0c0, 0x008000, 0x008080 })
            {
                int previous = VbeNativeChrome.EditorBackground;
                int scale = source == 0xc0c0c0 ? 192 : 128;
                for (int coverage = 0; coverage <= scale; coverage++)
                {
                    int input = ((source >> 16) * coverage / scale << 16) |
                        (((source >> 8) & 255) * coverage / scale << 8) |
                        ((source & 255) * coverage / scale);
                    int mapped = VbeNativeChrome.MapCodePixel(input);
                    Assert.AreEqual(mapped, VbeNativeChrome.MapCodePixel(mapped), "Repeated paints must not brighten edges.");
                    foreach (int shift in new[] { 16, 8, 0 })
                    {
                        int difference = ((mapped >> shift) & 255) - ((previous >> shift) & 255);
                        Assert.IsTrue(difference >= 0 && difference <= 2, "An antialiased edge must not introduce a color discontinuity.");
                    }
                    previous = mapped;
                }
            }
            Assert.AreEqual(0x282d35, VbeNativeChrome.MapCodePixel(0));
            Assert.AreEqual(0x569cd6, VbeNativeChrome.MapCodePixel(0x008080));
            Assert.AreEqual(0x57a64a, VbeNativeChrome.MapCodePixel(0x008000));
            Assert.AreEqual(0xffff00, VbeNativeChrome.MapCodePixel(0xffff00));
            Assert.AreEqual(0x800000, VbeNativeChrome.MapCodePixel(0x800000));
        }

        [TestMethod]
        public void CapturedClearTypeEdgesDoNotRemainOnTheBlackRampOrAccumulate()
        {
            // Pixels observed in the SOLIDWORKS VBE capture on 2026-09-28.
            foreach (int source in new[] { 0x6c2b00, 0xa56c2b, 0x00486e, 0x006e5c, 0x00486e, 0x002b6c })
            {
                int mapped = VbeNativeChrome.MapCodePixel(source);
                Assert.AreNotEqual(source, mapped);
                Assert.AreEqual(mapped, VbeNativeChrome.MapCodePixel(mapped));
                Assert.IsTrue((mapped >> 16) >= 40 && ((mapped >> 8) & 255) >= 45 && (mapped & 255) >= 53,
                    "Text edges must not retain channels darker than the editor background.");
            }
            foreach (int marker in new[] { 0x800000, 0x000080, 0x800080, 0x808000, 0xff0000, 0xffff00 })
                Assert.AreEqual(marker, VbeNativeChrome.MapCodePixel(marker));
        }

        [TestMethod]
        public void InitializationRefusesMissingEditorAndDisabledResetIsIdempotent()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeNativeTheme.Initialize(IntPtr.Zero, false));
            VbeNativeTheme.SetEnabled(false);
            VbeNativeTheme.Disconnect();
            VbeNativeTheme.Disconnect();
        }
    }
}
