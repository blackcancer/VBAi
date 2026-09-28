using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
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
        public void FailedNativeStopPreservesThemeOwnerAndResourcesAcrossCleanupPaths()
        {
            var flags = System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic;
            var owner = typeof(VbeNativeTheme).GetField("editorWindow", flags);
            var brush = typeof(VbeNativeTheme).GetField("backgroundBrush", flags);
            var hook = typeof(VbeNativeTheme).GetField("windowEventHook", flags);
            var previousOwner = owner.GetValue(null);
            var previousBrush = brush.GetValue(null);
            var previousHook = hook.GetValue(null);
            try
            {
                owner.SetValue(null, new IntPtr(0x123));
                brush.SetValue(null, new IntPtr(0x456));
                hook.SetValue(null, new IntPtr(0x789));
                using (var renderer = new NativeRendererStopFixture { Result = 1444 })
                {
                    Assert.IsFalse(VbeNativeTheme.Reset());
                    Assert.IsFalse(VbeNativeTheme.Disconnect());
                    Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.SetEnabled(false));
                    Assert.AreEqual(3, renderer.Calls);
                    Assert.IsTrue(VbeNativeRenderer.Active);
                    Assert.AreEqual(new IntPtr(0x123), owner.GetValue(null));
                    Assert.AreEqual(new IntPtr(0x456), brush.GetValue(null));
                    Assert.AreEqual(new IntPtr(0x789), hook.GetValue(null));
                }
            }
            finally
            {
                owner.SetValue(null, previousOwner);
                brush.SetValue(null, previousBrush);
                hook.SetValue(null, previousHook);
            }
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

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativeControlPaletteTests
    {
        [DllImport("user32.dll")] private static extern int SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        [TestMethod]
        public void NativeTreeRestoresCustomColorsAndSystemColorSentinel()
        {
            OnSta(() =>
            {
                using (var tree = new TreeView())
                {
                    IntPtr handle = tree.Handle;
                    foreach (int background in new[] { 0x00345678, -1 })
                    {
                        SendMessage(handle, 0x111d, IntPtr.Zero, new IntPtr(background));
                        SendMessage(handle, 0x111e, IntPtr.Zero, new IntPtr(0x00123456));
                        try
                        {
                            VbeNativeTheme.ApplyControlPalette(handle, "SysTreeView32");
                            Assert.AreNotEqual(background, SendMessage(handle, 0x111f, IntPtr.Zero, IntPtr.Zero));
                            // Repeated activation must not replace the original snapshot.
                            VbeNativeTheme.ApplyControlPalette(handle, "SysTreeView32");
                        }
                        finally { VbeNativeTheme.RestoreControlPalette(handle, "SysTreeView32"); }
                        Assert.AreEqual(background, SendMessage(handle, 0x111f, IntPtr.Zero, IntPtr.Zero));
                        Assert.AreEqual(0x00123456, SendMessage(handle, 0x1120, IntPtr.Zero, IntPtr.Zero));
                    }
                }
            });
        }

        [TestMethod]
        public void NativeListRestoresThreeIndependentColors()
        {
            OnSta(() =>
            {
                using (var list = new ListView())
                {
                    IntPtr handle = list.Handle;
                    SendMessage(handle, 0x1001, IntPtr.Zero, new IntPtr(0x00112233));
                    SendMessage(handle, 0x1024, IntPtr.Zero, new IntPtr(0x00445566));
                    SendMessage(handle, 0x1026, IntPtr.Zero, new IntPtr(0x00778899));
                    try { VbeNativeTheme.ApplyControlPalette(handle, "SysListView32"); }
                    finally { VbeNativeTheme.RestoreControlPalette(handle, "SysListView32"); }
                    Assert.AreEqual(0x00112233, SendMessage(handle, 0x1000, IntPtr.Zero, IntPtr.Zero));
                    Assert.AreEqual(0x00445566, SendMessage(handle, 0x1023, IntPtr.Zero, IntPtr.Zero));
                    Assert.AreEqual(0x00778899, SendMessage(handle, 0x1025, IntPtr.Zero, IntPtr.Zero));
                }
            });
        }

        private static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(5000), "The native control test did not complete.");
            if (failure != null) throw failure;
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeThemeOrchestrationTests
    {
        [TestMethod]
        public void InitializationAndApplyRejectMissingOwnerBeforeAnyNativeModeChange()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    Assert.ThrowsException<ArgumentException>(() => VbeNativeTheme.Initialize(IntPtr.Zero, false));
                    Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.SetEnabled(true));
                    Assert.ThrowsException<ArgumentException>(() => VbeNativeTheme.Apply(IntPtr.Zero));
                    Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.Apply(new IntPtr(-1)));
                    Assert.AreEqual(0, fixture.Applies); Assert.AreEqual(0, fixture.PreferredModes.Count);
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("editorWindow"));
                }
            });
        }

        [TestMethod]
        public void NullVbeInitializationClearsPaneAndRemainsDisabled()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    fixture.Set("immediateWindow", new IntPtr(7)); fixture.Set("immediateCaption", "Old immediate");
                    VbeNativeTheme.Initialize(fixture.Handle, false);
                    Assert.AreEqual(fixture.Handle, fixture.Read("editorWindow"));
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("immediateWindow")); Assert.IsNull(fixture.Read("immediateCaption"));
                    Assert.IsNull(fixture.Palette); Assert.AreEqual(0, fixture.Creates); Assert.AreEqual(0, fixture.Applies);
                    Assert.IsTrue(VbeNativeTheme.Disconnect()); Assert.IsTrue(VbeNativeTheme.Disconnect());
                }
            });
        }

        [TestMethod]
        public void PaneDiscoveryAndPaletteReplacementPreserveOnlyTheFirstImmediatePane()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    var noImmediate = new NativeThemeVbe { Windows = new object[] { new NativeThemePane { Type = 1 } } };
                    VbeNativeTheme.Initialize(fixture.Handle, false, noImmediate);
                    var previous = fixture.Palette;
                    Assert.IsFalse((bool)NativeThemeFixture.PaletteField(previous, "requested"));
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("immediateWindow"));
                    var vbe = new NativeThemeVbe { Windows = new object[] {
                        new NativeThemePane { Type = 1 }, new NativeThemePane { Type = 5, HWnd = 99, Caption = "Synthetic immediate" },
                        new NativeThemePane { Type = 5, HWnd = 100, Caption = "Ignored immediate" } } };
                    VbeNativeTheme.Initialize(fixture.Handle, true, vbe);
                    Assert.IsTrue((bool)NativeThemeFixture.PaletteField(previous, "disposed"));
                    Assert.AreEqual(new IntPtr(99), fixture.Read("immediateWindow"));
                    Assert.AreEqual("Synthetic immediate", fixture.Read("immediateCaption"));
                    Assert.AreEqual(2, fixture.Creates); Assert.AreEqual(1, fixture.Applies);
                    Assert.IsTrue((bool)NativeThemeFixture.PaletteField(fixture.Palette, "requested"));
                    VbeNativeTheme.SetEnabled(true);
                    Assert.AreEqual(1, fixture.Applies, "Repeated activation must not reinstall resources.");
                    VbeNativeTheme.SetEnabled(false);
                    Assert.IsFalse((bool)NativeThemeFixture.PaletteField(fixture.Palette, "requested"));
                    var activePalette = fixture.Palette;
                    Assert.IsTrue(VbeNativeTheme.Disconnect());
                    Assert.IsTrue((bool)NativeThemeFixture.PaletteField(activePalette, "disposed")); Assert.IsNull(fixture.Palette);
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("editorWindow"));
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("immediateWindow")); Assert.IsNull(fixture.Read("immediateCaption"));
                }
            });
        }

        [TestMethod]
        public void ExplicitExperimentEnablesThemeAndSuppressesProductionPaletteService()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, "1");
                    VbeNativeTheme.Initialize(fixture.Handle, false, new NativeThemeVbe());
                    Assert.AreEqual(1, fixture.Applies); Assert.AreEqual(0, fixture.Creates); Assert.IsNull(fixture.Palette);
                    Assert.AreEqual(1, fixture.Count("themedWindows"));
                    Assert.IsTrue(VbeNativeTheme.Disconnect());
                }
            });
        }

        [TestMethod]
        public void PaneDiscoveryFailureIsLoggedWhilePaletteInitializationContinues()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    VbeNativeTheme.Initialize(fixture.Handle, false, new NativeThemeBrokenVbe());
                    Assert.AreEqual(1, fixture.Creates); Assert.IsNotNull(fixture.Palette);
                    Assert.IsTrue(fixture.Messages.Exists(message => message.Contains("Native Immediate pane identification failed: synthetic pane enumeration failure")));
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("immediateWindow"));
                }
            });
        }

        [TestMethod]
        public void ApplicationFailureCleansPartialThemeAndPreservesOriginalException()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    fixture.Set("editorWindow", fixture.Handle);
                    var expected = new InvalidOperationException("synthetic partial apply failure"); fixture.ApplyFailure = expected;
                    var actual = Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.SetEnabled(true));
                    Assert.AreSame(expected, actual); Assert.AreEqual(1, fixture.Applies);
                    Assert.AreEqual(1, renderer.Calls); Assert.AreEqual(0, fixture.Count("themedWindows"));
                    Assert.AreEqual(fixture.Handle, fixture.Read("editorWindow"));
                }
            });
        }

        [TestMethod]
        public void StopRefusalPreservesEveryOwnerResourceAndDoesNotRequestPaletteRestore()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture())
                using (var fixture = new NativeThemeFixture())
                {
                    VbeNativeTheme.Initialize(fixture.Handle, true, new NativeThemeVbe());
                    IntPtr brush = fixture.AddBrush(), hook = fixture.AddEventHook();
                    fixture.Add("pendingChrome", fixture.Handle); fixture.Add("pendingCaptions", fixture.Handle);
                    fixture.Set("preferredModeChanged", true); fixture.Set("previousPreferredMode", 3);
                    fixture.ConfigureCallbacks(); renderer.Result = 1444;
                    var palette = fixture.Palette;
                    Assert.IsFalse(VbeNativeTheme.Reset()); Assert.IsFalse(VbeNativeTheme.Disconnect());
                    Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.SetEnabled(false));
                    Assert.AreEqual(3, renderer.Calls); Assert.IsTrue(VbeNativeRenderer.Active);
                    Assert.AreEqual(fixture.Handle, fixture.Read("editorWindow")); Assert.AreSame(palette, fixture.Palette);
                    Assert.IsFalse((bool)NativeThemeFixture.PaletteField(palette, "disposed"));
                    Assert.IsTrue((bool)NativeThemeFixture.PaletteField(palette, "requested"));
                    Assert.AreEqual(brush, fixture.Read("backgroundBrush")); Assert.AreEqual((uint)2, NativeThemeFixture.GetObjectType(brush));
                    Assert.AreEqual(hook, fixture.Read("windowEventHook")); Assert.AreEqual(1, fixture.Count("themedWindows"));
                    Assert.AreEqual(1, fixture.Count("pendingChrome")); Assert.AreEqual(1, fixture.Count("pendingCaptions"));
                    Assert.IsTrue((bool)fixture.Read("preferredModeChanged")); Assert.AreEqual(0, fixture.PreferredModes.Count);
                    Assert.AreEqual(0, fixture.RestoredWindows.Count); Assert.AreEqual(0, fixture.Flushes);
                }
            });
        }

        [TestMethod]
        public void ConfirmedResetRestoresPreferenceDeletesBrushAndRemovesWindowResources()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture())
                using (var fixture = new NativeThemeFixture())
                {
                    fixture.ConfigureCallbacks(); fixture.Set("editorWindow", fixture.Handle);
                    IntPtr dialog = fixture.AddDialog(); fixture.Add("themedWindows", fixture.Handle); fixture.Add("themedWindows", dialog);
                    fixture.AddSubclass(fixture.Handle); Assert.IsTrue(fixture.IsSubclassed(fixture.Handle));
                    IntPtr brush = fixture.AddBrush(); fixture.AddEventHook();
                    var propertyTabs = fixture.AddPropertyTabs();
                    fixture.Add("pendingChrome", fixture.Handle); fixture.Add("pendingCaptions", fixture.Handle);
                    fixture.Set("toolbarPaintCount", 7); fixture.Set("toolbarDeferredPaintCount", 8);
                    fixture.Set("preferredModeChanged", true); fixture.Set("previousPreferredMode", 3);
                    Assert.IsTrue(VbeNativeTheme.Reset()); Assert.IsFalse(VbeNativeRenderer.Active);
                    Assert.AreEqual((uint)0, NativeThemeFixture.GetObjectType(brush)); Assert.AreEqual(IntPtr.Zero, fixture.Read("backgroundBrush"));
                    Assert.AreEqual(IntPtr.Zero, fixture.Read("windowEventHook")); Assert.IsFalse(fixture.IsSubclassed(fixture.Handle));
                    IntPtr result; Assert.IsFalse(propertyTabs.TryHandleMessage(0x14, IntPtr.Zero, IntPtr.Zero, out result));
                    foreach (string name in new[] { "themedWindows", "subclassedWindows", "pendingChrome", "pendingCaptions", "propertyTabs", "originalControlColors" })
                        Assert.AreEqual(0, fixture.Count(name));
                    Assert.AreEqual(0, fixture.Read("toolbarPaintCount")); Assert.AreEqual(0, fixture.Read("toolbarDeferredPaintCount"));
                    CollectionAssert.AreEqual(new[] { fixture.Handle, dialog }, fixture.RestoredWindows);
                    CollectionAssert.AreEqual(new[] { 3 }, fixture.PreferredModes);
                    Assert.AreEqual(1, fixture.Flushes); Assert.IsFalse((bool)fixture.Read("preferredModeChanged"));
                    Assert.AreEqual(fixture.Handle, fixture.Read("editorWindow"));
                    Assert.IsTrue(VbeNativeTheme.Reset()); Assert.AreEqual(1, fixture.PreferredModes.Count); Assert.AreEqual(2, fixture.Flushes);
                }
            });
        }

        [TestMethod]
        public void MissingRestoreDelegatesStillClearConfirmedResourcesAndPreferenceFlag()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var renderer = new NativeRendererStopFixture(false))
                using (var fixture = new NativeThemeFixture())
                {
                    fixture.ConfigureCallbacks(false, false, false);
                    fixture.Add("themedWindows", fixture.Handle); fixture.Set("preferredModeChanged", true);
                    Assert.IsTrue(VbeNativeTheme.Reset()); Assert.AreEqual(0, fixture.Count("themedWindows"));
                    Assert.IsFalse((bool)fixture.Read("preferredModeChanged")); Assert.AreEqual(0, fixture.Flushes);
                    Assert.AreEqual(0, fixture.PreferredModes.Count); Assert.AreEqual(IntPtr.Zero, fixture.Read("editorWindow"));
                }
            });
        }
    }
}
