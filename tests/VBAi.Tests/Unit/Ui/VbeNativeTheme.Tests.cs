using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class VbeNativeThemeNativeTests
    {
        private static void WithFixture(Action<NativeThemeNativeFixture, NativeRendererStopFixture> action) => NativeThemeFixture.OnSta(() =>
        {
            using (var renderer = new NativeRendererStopFixture(false))
            using (var fixture = new NativeThemeNativeFixture()) { renderer.ConfigureLifecycle(); action(fixture, renderer); }
        });

        [TestMethod]
        public void VersionExportsAndWindowOwnershipAreVerifiedBeforeNativeApplication()
        {
            WithFixture((fixture, renderer) =>
            {
                fixture.VersionResult = 1;
                Assert.ThrowsException<PlatformNotSupportedException>(() => VbeNativeTheme.ReadNativeWindowsVersion());
                fixture.VersionResult = 0; fixture.Build = 17763;
                Assert.ThrowsException<PlatformNotSupportedException>(() => VbeNativeTheme.Apply(fixture.Main));
                fixture.Build = 22000; fixture.ModuleAvailable = false;
                Assert.ThrowsException<PlatformNotSupportedException>(() => VbeNativeTheme.Apply(fixture.Main));
                fixture.ModuleAvailable = true;
                foreach (int ordinal in new[] { 135, 133 })
                {
                    fixture.MissingExports.Add(ordinal);
                    Assert.ThrowsException<PlatformNotSupportedException>(() => VbeNativeTheme.Apply(fixture.Main));
                    fixture.MissingExports.Clear();
                }
                fixture.ForeignThread.Add(fixture.Main);
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.Apply(fixture.Main));
                fixture.ForeignThread.Clear(); fixture.ForeignProcess.Add(fixture.Main);
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativeTheme.Apply(fixture.Main));
                Assert.AreEqual(0, renderer.StartCalls); Assert.AreEqual(0, fixture.Modes.Count);
            });
        }

        [TestMethod]
        public void ApplicationUsesManagedPreferredModeOnlyInsideExplicitExperimentAndRestoresIt()
        {
            foreach (bool experiment in new[] { false, true })
                WithFixture((fixture, renderer) =>
                {
                    Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, experiment ? "1" : null);
                    Environment.SetEnvironmentVariable(VbeNativeTheme.LocalRefreshExperimentVariable, "1");
                    IntPtr child = fixture.Add("#32770", fixture.Main); fixture.Dwm20Result = 1;
                    Assert.AreEqual(2, VbeNativeTheme.Apply(fixture.Main));
                    Assert.AreEqual(1, renderer.StartCalls); Assert.IsTrue(VbeNativeRenderer.Active);
                    Assert.AreEqual(experiment ? 1 : 0, fixture.Modes.Count); Assert.AreEqual(experiment ? 1 : 0, fixture.RefreshPolicies);
                    if (experiment) Assert.AreEqual(2, fixture.Modes[0]);
                    Assert.AreEqual(1, fixture.Flushes); Assert.AreEqual(1, fixture.HookCalls); Assert.AreEqual(new IntPtr(9), fixture.State.Read("windowEventHook"));
                    Assert.IsTrue(fixture.Attributes.Exists(a => a.Item1 == child && a.Item2 == 19 && a.Item3 == 1));
                    Assert.IsTrue(fixture.Attributes.Exists(a => a.Item1 == fixture.Main && a.Item2 == 19 && a.Item3 == 1));
                    Assert.IsTrue((bool)fixture.State.Read("localChromeRefresh"));
                    Assert.IsTrue(VbeNativeTheme.Reset()); Assert.AreEqual(1, fixture.Unhooks);
                    Assert.AreEqual(experiment ? 2 : 0, fixture.Modes.Count); if (experiment) Assert.AreEqual(3, fixture.Modes[1]);
                    Assert.AreEqual(0, fixture.State.Count("themedWindows")); Assert.IsFalse(VbeNativeRenderer.Active);
                });
            WithFixture((fixture, renderer) =>
            {
                Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, "1");
                fixture.MissingExports.Add(104); fixture.MissingExports.Add(136);
                Assert.AreEqual(1, VbeNativeTheme.Apply(fixture.Main)); Assert.AreEqual(0, fixture.RefreshPolicies); Assert.AreEqual(0, fixture.Flushes);
                Assert.IsTrue(VbeNativeTheme.Reset());
            });
        }

        [TestMethod]
        public void WindowPoliciesHandleDuplicatesForeignThreadsManagedClassesAndImmediateCaption()
        {
            WithFixture((fixture, renderer) =>
            {
                fixture.ApplyWindow(IntPtr.Zero);
                IntPtr empty = fixture.Add(""); fixture.ApplyWindow(empty);
                foreach (string className in new[] { "WindowsForms10.Window.fixture", "HwndWrapper[fixture]" }) fixture.ApplyWindow(fixture.Add(className));
                IntPtr foreign = fixture.Add("Edit"); fixture.ForeignThread.Add(foreign); fixture.ApplyWindow(foreign);
                Assert.AreEqual(0, fixture.State.Count("themedWindows"));
                IntPtr generic = fixture.Add("GenericPane"); fixture.ApplyWindow(generic); fixture.ApplyWindow(generic);
                Assert.AreEqual(1, fixture.State.Count("themedWindows")); Assert.AreEqual(1, fixture.Installations.Count);
                Assert.IsTrue(fixture.State.Messages.Exists(m => m.Contains("Native theme GenericPane")));
                IntPtr toolbar = fixture.Add("MsoCommandBar"); renderer.SetActive(true); fixture.ApplyWindow(toolbar);
                Assert.AreEqual(1, renderer.RegisterCalls);
                fixture.State.Set("immediateCaption", "Target immediate");
                IntPtr wrong = fixture.Add("VbaWindow"); fixture.Captions[wrong] = "Other"; fixture.ApplyWindow(wrong);
                Assert.AreEqual(IntPtr.Zero, fixture.State.Read("immediateWindow"));
                IntPtr immediate = fixture.Add("VbaWindow"); fixture.Captions[immediate] = "Target immediate"; fixture.ApplyWindow(immediate);
                Assert.AreEqual(immediate, fixture.State.Read("immediateWindow"));
                fixture.ApplyWindow(fixture.Add("VbaWindow"));
                Assert.AreEqual(immediate, fixture.State.Read("immediateWindow"));
                fixture.SubclassSucceeds = false; IntPtr refused = fixture.Add("ComboBox"); fixture.ApplyWindow(refused);
                Assert.IsFalse(fixture.DarkCalls.Exists(c => c.Item1 == refused));
            });
        }

        [TestMethod]
        public void ListTabDialogAndButtonScopingChooseClassicRendererOnlyWhereRequired()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr properties = fixture.Add("wndclass_pbrs", fixture.Main), dialog = fixture.Add("#32770", owner: fixture.Main);
                IntPtr listOutside = fixture.Add("ListBox", fixture.Main), listInside = fixture.Add("ListBox", properties);
                fixture.ApplyWindow(listOutside); fixture.ApplyWindow(listInside);
                Assert.IsFalse(fixture.Installations.Contains(listOutside)); Assert.IsTrue(fixture.Installations.Contains(listInside));
                IntPtr tabOutside = fixture.Add("SysTabControl32", fixture.Main), tabProperties = fixture.Add("SysTabControl32", properties), tabDialog = fixture.Add("SysTabControl32", dialog);
                fixture.ApplyWindow(tabOutside); fixture.ApplyWindow(tabProperties); fixture.ApplyWindow(tabDialog);
                Assert.IsFalse(fixture.Installations.Contains(tabOutside)); Assert.IsTrue(fixture.Installations.Contains(tabProperties)); Assert.IsTrue(fixture.Installations.Contains(tabDialog));
                Assert.IsTrue(fixture.Themes.Exists(t => t.Item1 == tabProperties && t.Item2 == "" && t.Item3 == ""));
                foreach (int style in new[] { 0, 1, 2, 7, 8, 9, 10 })
                {
                    IntPtr button = fixture.Add("Button", dialog); fixture.Styles[button] = style; fixture.ApplyWindow(button);
                    Assert.AreEqual((style >= 2 && style <= 7) || style == 9, fixture.Themes.Exists(t => t.Item1 == button && t.Item2 == ""));
                }
                fixture.ApplyWindow(fixture.Add("Button", fixture.Main));
                fixture.ApplyWindow(dialog);
                Assert.IsTrue(fixture.Attributes.Exists(t => t.Item1 == dialog && t.Item2 == 20 && t.Item3 == 1));
            });
        }

        [TestMethod]
        public void OwnershipAndEventNotificationsScopeRefreshChildrenAndForgetDestroyedResources()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr child = fixture.Add("Edit", fixture.Main), popup = fixture.Add("#32770", owner: fixture.Main);
                IntPtr ownedByChild = fixture.Add("Edit", owner: child), unrelated = fixture.Add("Edit", owner: IntPtr.Zero);
                Assert.IsTrue(fixture.Belongs(fixture.Main)); Assert.IsTrue(fixture.Belongs(child)); Assert.IsTrue(fixture.Belongs(popup)); Assert.IsTrue(fixture.Belongs(ownedByChild));
                Assert.IsFalse(fixture.Belongs(IntPtr.Zero)); Assert.IsFalse(fixture.Belongs(unrelated));
                fixture.State.Set("editorWindow", IntPtr.Zero); Assert.IsFalse(fixture.Belongs(child)); fixture.State.Set("editorWindow", fixture.Main);
                fixture.Event(0x8000, child, objectId: 1); fixture.Event(0x8000, child, childId: 1); fixture.Event(0x8000, unrelated);
                Assert.AreEqual(0, fixture.Redraws.Count);
                renderer.SetActive(true); fixture.Event(0x8000, child); Assert.AreEqual(1, renderer.RefreshCalls);
                IntPtr nested = fixture.Add("ComboBox", popup); fixture.Event(0x8002, popup);
                Assert.IsTrue(fixture.Installations.Contains(nested)); Assert.AreEqual(2, renderer.RefreshCalls);
                IntPtr foreign = fixture.Add("Edit", fixture.Main); fixture.ForeignThread.Add(foreign); fixture.Event(0x8000, foreign);
                Assert.AreEqual(2, renderer.RefreshCalls); Assert.IsFalse(fixture.Installations.Contains(foreign));
                fixture.State.Add("pendingChrome", child); fixture.State.Add("pendingCaptions", child); fixture.State.Set("immediateWindow", child);
                fixture.Event(0x8001, child); Assert.AreEqual(IntPtr.Zero, fixture.State.Read("immediateWindow"));
                Assert.IsFalse(((System.Collections.IList)fixture.State.Read("themedWindows")).Contains(child));
                Assert.AreEqual(0, fixture.State.Count("pendingChrome")); Assert.AreEqual(0, fixture.State.Count("pendingCaptions"));
            });
        }

        [TestMethod]
        public void PaintRoutingHonorsNativeToolbarDirectTabsCodeSurfaceAndDcOwnership()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr toolbarChild = fixture.Add("MsoCommandBar", fixture.Main), toolbarPopup = fixture.Add("MsoCommandBar", owner: fixture.Main);
                renderer.SetActive(true); fixture.Paint(toolbarChild, "MsoCommandBar", IntPtr.Zero); Assert.AreEqual(0, fixture.Paints.Count);
                fixture.Paint(toolbarPopup, "MsoCommandBar", IntPtr.Zero); Assert.AreEqual(1, fixture.Paints.Count);
                renderer.SetActive(false); fixture.Paint(toolbarChild, "MsoCommandBar", IntPtr.Zero); Assert.AreEqual(2, fixture.Paints.Count);
                IntPtr code = fixture.Add("VbaWindow", fixture.Main); fixture.CodeChildren.Add(code);
                fixture.Paint(code, "VbaWindow", new IntPtr(4)); Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == code && p.Item2 && p.Item6 && p.Item3 == new IntPtr(4)));
                IntPtr immediate = fixture.Add("VbaWindow", fixture.Main); fixture.State.Set("immediateWindow", immediate); Assert.IsTrue(fixture.IsCode(immediate, "VbaWindow"));
                IntPtr designer = fixture.Add("VbaWindow", fixture.Main); fixture.Paint(designer, "VbaWindow", IntPtr.Zero);
                Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == designer && p.Item5 && !p.Item6));
                IntPtr generic = fixture.Add("GenericPane", fixture.Main); fixture.Paint(generic, "GenericPane", IntPtr.Zero);
                Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == generic && p.Item4 && !p.Item2));
                IntPtr combo = fixture.Add("ComboBox", fixture.Main); fixture.Paint(combo, "ComboBox", IntPtr.Zero); fixture.Paint(combo, "ComboBox", new IntPtr(4));
                Assert.AreEqual(1, fixture.Combos.Count);
                foreach (string kind in new[] { "MsoCommandBarPopup", "MsoCommandBarDock", "SysTabControl32", "Other" }) fixture.Paint(fixture.Add(kind), kind, IntPtr.Zero);
                int borders = fixture.Borders.Count; fixture.Paint(fixture.Main, "Other", IntPtr.Zero); Assert.AreEqual(borders, fixture.Borders.Count);
                var tabs = fixture.AddTabs(fixture.Main); int paints = fixture.Paints.Count; fixture.Paint(fixture.TabsWindow(tabs), "SysTabControl32", IntPtr.Zero); Assert.AreEqual(paints, fixture.Paints.Count);
            });
        }

        [TestMethod]
        public void ContainerRefreshQueuesOnlyOwnedDockToolbarAndHostedCaptionAndCoalescesPosts()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr dock = fixture.Add("MsoCommandBarDock", fixture.Main), toolbar = fixture.Add("MsoCommandBar", dock), generic = fixture.Add("GenericPane", fixture.Main);
                foreach (var window in new[] { toolbar, generic, fixture.Add("Other", dock), fixture.Add("MsoCommandBar", fixture.Main) }) fixture.State.Add("subclassedWindows", window);
                fixture.PostSucceeds = false; fixture.Queue(fixture.Main);
                Assert.AreEqual(0, fixture.State.Count("pendingChrome")); Assert.AreEqual(0, fixture.State.Count("pendingCaptions")); Assert.AreEqual(2, fixture.Posts.Count);
                fixture.PostSucceeds = true; fixture.Queue(fixture.Main); fixture.Queue(fixture.Main);
                Assert.AreEqual(1, fixture.State.Count("pendingChrome")); Assert.AreEqual(1, fixture.State.Count("pendingCaptions")); Assert.AreEqual(4, fixture.Posts.Count);
                Assert.AreEqual(new IntPtr(0), fixture.Message(toolbar, 0x8564)); Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                Assert.AreEqual(1, fixture.State.Read("toolbarDeferredPaintCount"));
                fixture.Message(generic, 0x8565); Assert.AreEqual(0, fixture.State.Count("pendingCaptions"));
                Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == generic && p.Item4));
            });
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeNativeThemeNativeTests
    {
        [DllImport("gdi32.dll")] private static extern uint GetTextColor(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern uint GetBkColor(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern uint SetTextColor(IntPtr dc, uint color);

        [TestMethod]
        public void SubclassQueryAndDeferredMessagesRespectExperimentPendingAndRendererGuards()
        {
            WithFixture((fixture, renderer) =>
            {
                fixture.State.Set("toolbarPaintCount", 7); fixture.State.Set("toolbarDeferredPaintCount", 9);
                Assert.AreEqual(fixture.NativeResult, fixture.Message(fixture.Main, 0x856b));
                Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, "1");
                Assert.AreEqual(new IntPtr(7), fixture.Message(fixture.Main, 0x856b)); Assert.AreEqual(new IntPtr(9), fixture.Message(fixture.Main, 0x856b, new IntPtr(1)));
                IntPtr window = fixture.Add("Other");
                fixture.Message(window, 0x8565); fixture.Message(window, 0x8564); Assert.AreEqual(0, fixture.Paints.Count);
                fixture.State.Add("pendingCaptions", window); fixture.State.Add("pendingChrome", window);
                fixture.Message(window, 0x8565); fixture.Message(window, 0x8564); Assert.AreEqual(0, fixture.Paints.Count);
                fixture.State.Add("subclassedWindows", window); fixture.State.Add("pendingCaptions", window); fixture.Message(window, 0x8565);
                Assert.AreEqual(1, fixture.Paints.Count); Assert.IsFalse(fixture.Paints[0].Item4);
                foreach (string name in new[] { "PROJECT", "wndclass_pbrs", "VbaWindow", "Other" })
                {
                    IntPtr target = fixture.Add(name); fixture.State.Add("subclassedWindows", target);
                    fixture.State.Add("pendingChrome", target); fixture.Message(target, 0x8564);
                    fixture.State.Set("localChromeRefresh", true); fixture.State.Add("pendingChrome", target); fixture.Message(target, 0x8564);
                    fixture.State.Set("localChromeRefresh", false);
                }
                Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                var tabs = fixture.AddTabs(fixture.Main); IntPtr tabWindow = fixture.TabsWindow(tabs);
                fixture.State.Add("subclassedWindows", tabWindow); fixture.State.Add("pendingChrome", tabWindow);
                int paints = fixture.Paints.Count; fixture.Message(tabWindow, 0x8564); Assert.AreEqual(paints, fixture.Paints.Count);
            });
        }

        [TestMethod]
        public void DirectPropertyRendererQueriesHandlePaintAndRetainNativeMessageBehavior()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr properties = fixture.Add("wndclass_pbrs", fixture.Main);
                var tabs = fixture.AddTabs(properties); IntPtr window = fixture.TabsWindow(tabs);
                Assert.IsTrue(NativeThemeNativeFixture.Invoke<bool>("TryAttachPropertyTabs", window));
                Assert.AreEqual(fixture.NativeResult, fixture.Message(window, 0x856a));
                Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, "1");
                Assert.AreEqual(new IntPtr(1), fixture.Message(window, 0x856a, new IntPtr(2)));
                Assert.AreEqual(IntPtr.Zero, fixture.Message(window, 0x856a)); Assert.AreEqual(IntPtr.Zero, fixture.Message(window, 0x856a, new IntPtr(1)));
                Assert.AreEqual(new IntPtr(1), fixture.Message(window, 0x14));
                using (var bitmap = NativeChromeCanvas.Paint(System.Drawing.Color.White, dc => fixture.Message(window, 0x318, dc, new IntPtr(4))))
                    Assert.AreEqual(24, bitmap.Width);
                Assert.AreEqual(new IntPtr(1), fixture.Message(window, 0x856a, new IntPtr(1)));
                fixture.Message(window, 0xf); Assert.AreEqual(new IntPtr(1), fixture.Message(window, 0x856a));
                NativePaletteDialogFixture.SetStyle(window, -16, 0x40800000);
                Assert.AreEqual(new IntPtr(-1), fixture.Message(window, 0x856a, new IntPtr(2)));
                NativePaletteDialogFixture.SetStyle(window, -16, 0x40000000);
                Assert.AreEqual(fixture.NativeResult, fixture.Message(window, 0x9999));
                Assert.AreEqual(0, fixture.Paints.Count); Assert.AreEqual(0, fixture.Borders.Count);
                fixture.State.Set("immediateWindow", window); fixture.State.Add("pendingChrome", window); fixture.State.Add("pendingCaptions", window);
                fixture.Message(window, 0x82); Assert.IsTrue(fixture.Removals.Contains(window));
                Assert.AreEqual(0, fixture.State.Count("propertyTabs")); Assert.AreEqual(IntPtr.Zero, fixture.State.Read("immediateWindow"));
                IntPtr handled; Assert.IsFalse(tabs.TryHandleMessage(0x14, IntPtr.Zero, IntPtr.Zero, out handled));
            });
        }

        [TestMethod]
        public void NativeEraseAndControlColorMessagesUseOneOwnedBrushAndExactGdiColors()
        {
            WithFixture((fixture, renderer) =>
            {
                foreach (string kind in new[] { "#32770", "MDIClient", "VBSlider" })
                {
                    IntPtr window = fixture.Add(kind);
                    using (var bitmap = NativeChromeCanvas.Paint(System.Drawing.Color.White, dc => Assert.AreEqual(new IntPtr(1), fixture.Message(window, 0x14, dc))))
                    {
                        Assert.AreEqual(0x20242b, bitmap.GetPixel(0, 0).ToArgb() & 0xffffff);
                        Assert.AreEqual(0xffffff, bitmap.GetPixel(8, 8).ToArgb() & 0xffffff);
                    }
                }
                IntPtr unknown = fixture.Add("Other"); Assert.AreEqual(fixture.NativeResult, fixture.Message(unknown, 0x14));
                fixture.ClientSucceeds = false; Assert.AreEqual(fixture.NativeResult, fixture.Message(fixture.Add("#32770"), 0x14)); fixture.ClientSucceeds = true;
                IntPtr brush = (IntPtr)fixture.State.Read("backgroundBrush"); Assert.AreNotEqual(IntPtr.Zero, brush);
                using (var bitmap = NativeChromeCanvas.Paint(System.Drawing.Color.White, dc =>
                {
                    foreach (uint message in new uint[] { 0x133, 0x134, 0x135, 0x136, 0x137, 0x138 })
                    {
                        Assert.AreEqual(brush, fixture.Message(unknown, message, dc));
                        Assert.AreEqual((uint)0x00f0e8e2, GetTextColor(dc)); Assert.AreEqual((uint)0x002b2420, GetBkColor(dc));
                    }
                })) Assert.AreEqual(0xffffff, bitmap.GetPixel(0, 0).ToArgb() & 0xffffff);
            });
        }

        [TestMethod]
        public void NativePropertyDrawItemVerifiesEveryPredicateAndRestoresCallerDc()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr properties = fixture.Add("wndclass_pbrs"), list = fixture.Add("ListBox", properties), other = fixture.Add("Other");
                Assert.AreEqual(fixture.NativeResult, fixture.Message(properties, 0x2b));
                using (var bitmap = NativeChromeCanvas.Paint(System.Drawing.Color.White, dc =>
                {
                    foreach (int guard in new[] { 0, 1, 2, 3, 4, 5 })
                    {
                        var item = new VbeNativeTheme.NativeDrawItem { ControlType = guard == 0 ? 1u : 2u,
                            Window = guard == 1 ? other : list, DeviceContext = guard == 2 ? IntPtr.Zero : dc,
                            Action = guard == 3 ? 0u : 1u, Bounds = new VbeNativeTheme.NativeRect { Right = 4, Bottom = 4 } };
                        IntPtr memory = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(VbeNativeTheme.NativeDrawItem)));
                        try
                        {
                            Marshal.StructureToPtr(item, memory, false);
                            var save = VbeNativeTheme.SaveDeviceContext;
                            if (guard == 4) VbeNativeTheme.SaveDeviceContext = device => 0;
                            try
                            {
                                SetTextColor(dc, 0x00123456);
                                Assert.AreEqual(fixture.NativeResult, fixture.Message(guard == 5 ? properties : properties, 0x2b, IntPtr.Zero, memory));
                                Assert.AreEqual((uint)0x00123456, GetTextColor(dc));
                            }
                            finally { VbeNativeTheme.SaveDeviceContext = save; }
                        }
                        finally { Marshal.FreeHGlobal(memory); }
                    }
                })) Assert.AreEqual(24, bitmap.Width);
                Assert.AreEqual(1, fixture.Rows.Count); Assert.AreEqual(4, fixture.Rows[0].Item2.Right);
                Assert.AreEqual(fixture.NativeResult, fixture.Message(other, 0x2b, IntPtr.Zero, new IntPtr(1)));
            });
        }

        [TestMethod]
        public void TabInitializationRetriesAttachOnlySupportedPropertiesControls()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr properties = fixture.Add("wndclass_pbrs", fixture.Main), fake = fixture.Add("SysTabControl32", properties);
                foreach (uint message in new uint[] { 0x18, 0x47, 0x1307, 0x133e }) Assert.AreEqual(fixture.NativeResult, fixture.Message(fake, message));
                Assert.AreEqual(0, fixture.State.Count("propertyTabs"));
                fixture.Message(fixture.Add("SysTabControl32", fixture.Main), 0x18);
                var original = fixture.AddTabs(properties); IntPtr window = fixture.TabsWindow(original);
                ((System.Collections.IDictionary)fixture.State.Read("propertyTabs")).Remove(window); original.Dispose(); fixture.State.Add("pendingChrome", window);
                Assert.AreEqual(fixture.NativeResult, fixture.Message(window, 0x18));
                Assert.AreEqual(1, fixture.State.Count("propertyTabs")); Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                Assert.IsTrue(fixture.State.Messages.Exists(m => m.Contains("direct text renderer attached")));
            });
        }

        [TestMethod]
        public void ComboCaptionAndCodeMouseMessagesRouteOnlyTheExpectedSurfaces()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr combo = fixture.Add("ComboBox"), code = fixture.Add("VbaWindow"), other = fixture.Add("Other"); fixture.CodeChildren.Add(code);
                foreach (uint message in new uint[] { 0x201, 0x202, 0x14f, 0xa }) fixture.Message(combo, message);
                Assert.AreEqual(4, fixture.Combos.Count);
                foreach (uint message in new uint[] { 0x85, 0x86, 0xc, 0x317 }) fixture.Message(other, message, new IntPtr(41));
                Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == other && p.Item3 == new IntPtr(41) && !p.Item2));
                int paints = fixture.Paints.Count; fixture.Message(fixture.Main, 0x86); Assert.AreEqual(paints, fixture.Paints.Count);
                fixture.Message(code, 0x200); Assert.AreEqual(paints, fixture.Paints.Count);
                fixture.Message(code, 0x200, new IntPtr(1)); Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == code && p.Item6));
                foreach (uint message in new uint[] { 0xf, 0x317, 0x318, 0x2a3 }) fixture.Message(code, message, new IntPtr(42));
                Assert.IsTrue(fixture.Paints.Exists(p => p.Item1 == code && p.Item3 == new IntPtr(42) && p.Item6));
                fixture.Message(other, 0x200);
            });
        }

        [TestMethod]
        public void EditorTextAndScrollChangesCoalesceRefreshAndRecoverFromFailedPost()
        {
            WithFixture((fixture, renderer) =>
            {
                IntPtr code = fixture.Add("VbaWindow"); fixture.CodeChildren.Add(code);
                foreach (uint message in new uint[] { 0x102, 0x101, 0x202, 0x114, 0x115, 0x20a }) fixture.Message(code, message);
                Assert.AreEqual(1, fixture.Posts.Count); Assert.AreEqual(1, fixture.State.Count("pendingChrome"));
                fixture.Message(code, 0x8564); Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                fixture.PostSucceeds = false; fixture.Message(code, 0x102); Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                fixture.Message(fixture.Add("Other"), 0x102); Assert.AreEqual(2, fixture.Posts.Count);
            });
        }

        [TestMethod]
        public void LocalCaptionAndContainerNotificationsStayScopedAndHandlePostFailures()
        {
            WithFixture((fixture, renderer) =>
            {
                fixture.State.Set("localChromeRefresh", true);
                IntPtr dock = fixture.Add("MsoCommandBarDock", fixture.Main), toolbar = fixture.Add("MsoCommandBar", dock), generic = fixture.Add("GenericPane", fixture.Main);
                fixture.State.Add("subclassedWindows", toolbar); fixture.State.Add("subclassedWindows", generic);
                foreach (string kind in new[] { "PROJECT", "wndclass_pbrs", "VbaWindow", "GenericPane", "Other" })
                {
                    IntPtr caption = fixture.Add(kind); fixture.PostSucceeds = false;
                    fixture.Message(caption, 0x86); Assert.AreEqual(0, fixture.State.Count("pendingCaptions"));
                    fixture.PostSucceeds = true; fixture.Message(caption, 0xc); fixture.Message(caption, 0xc);
                    if (kind != "Other") { Assert.AreEqual(1, fixture.State.Count("pendingCaptions")); fixture.Message(caption, 0x8565); }
                }
                IntPtr code = fixture.Add("VbaWindow", fixture.Main); fixture.CodeChildren.Add(code);
                foreach (uint message in new uint[] { 0x222, 0x102, 0x101, 0x202, 0x114, 0x115, 0x20a }) fixture.Message(code, message);
                Assert.IsTrue(fixture.State.Count("pendingChrome") > 0); Assert.IsTrue(fixture.State.Count("pendingCaptions") > 0);
                foreach (uint message in new uint[] { 0xf, 0x85, 0x222 }) fixture.Message(fixture.Main, message);
                fixture.Message(fixture.Add("MDIClient", fixture.Main), 0xf); fixture.Message(code, 0xf);
                IntPtr project = fixture.Add("PROJECT", fixture.Main); fixture.Message(project, 0xf); fixture.Message(project, 0x4e);
                fixture.Message(fixture.Add("Other"), 0xf); fixture.Message(fixture.Add("Other"), 0x4e); fixture.Message(generic, 0x47);
            });
        }

        [TestMethod]
        public void LegacyRepaintFansOutOnlySupportedSubclassCandidatesAndSkipsDirectTabs()
        {
            WithFixture((fixture, renderer) =>
            {
                string[] classes = { "MsoCommandBar", "MsoCommandBarPopup", "MsoCommandBarDock", "GenericPane", "SysTabControl32", "ListBox", "VbaWindow", "PROJECT", "wndclass_pbrs", "SysTreeView32", "ComboBox", "Edit", "Other" };
                foreach (string kind in classes) fixture.State.Add("subclassedWindows", fixture.Add(kind, fixture.Main));
                var tabs = fixture.AddTabs(fixture.Main); fixture.State.Add("subclassedWindows", fixture.TabsWindow(tabs));
                fixture.Message(fixture.Main, 0xf); Assert.AreEqual(12, fixture.State.Count("pendingChrome")); Assert.AreEqual(12, fixture.Posts.Count);
                foreach (uint message in new uint[] { 0xf, 0x85, 0x4e, 0x47 }) fixture.Message(fixture.Main, message);
                Assert.AreEqual(12, fixture.Posts.Count);
                foreach (IntPtr window in ((System.Collections.IList)fixture.State.Read("subclassedWindows"))) fixture.Message(window, 0x8564);
                fixture.PostSucceeds = false; fixture.Message(fixture.Main, 0xf); Assert.AreEqual(0, fixture.State.Count("pendingChrome"));
                Assert.AreEqual(24, fixture.Posts.Count);
            });
        }
        [TestMethod]
        public void RichEditSnapshotsPreserveCustomBackgroundAndDiscardChangedClassEntries()
        {
            WithFixture((fixture, renderer) =>
            {
                foreach (string kind in new[] { "RichEdit20A", "RICHEDIT60W" })
                    using (var rich = new RichTextBox())
                    {
                        IntPtr window = rich.Handle;
                        var original = new IntPtr(0x00332211);
                        var dark = new IntPtr(0x002b2420);
                        VbeNativeTheme.NativeSend(window, 0x443, IntPtr.Zero, original);
                        VbeNativeTheme.ApplyControlPalette(window, kind);
                        Assert.AreEqual(dark.ToInt32(), VbeNativeTheme.NativeSend(window, 0x443, IntPtr.Zero, dark));
                        VbeNativeTheme.ApplyControlPalette(window, kind);
                        VbeNativeTheme.RestoreControlPalette(window, kind);
                        Assert.AreEqual(original.ToInt32(), VbeNativeTheme.NativeSend(window, 0x443, IntPtr.Zero, original));
                        Assert.AreEqual(0, fixture.State.Count("originalControlColors"));
                        VbeNativeTheme.ApplyControlPalette(window, kind);
                        VbeNativeTheme.RestoreControlPalette(window, "Other");
                        Assert.AreEqual(0, fixture.State.Count("originalControlColors"));
                        Assert.AreEqual(dark.ToInt32(), VbeNativeTheme.NativeSend(window, 0x443, IntPtr.Zero, original));
                    }
            });
        }

        [TestMethod]
        public void OwnershipTraversalFollowsBoundedPopupChainsAndStopsAtUnrelatedRoots()
        {
            WithFixture((fixture, renderer) =>
            {
                Assert.IsFalse(fixture.Belongs(new IntPtr(-1)));
                IntPtr unrelated = fixture.Add("Other", owner: IntPtr.Zero);
                Assert.IsFalse(fixture.Belongs(fixture.Add("Other", owner: unrelated)));
                IntPtr popup = fixture.Main;
                for (int depth = 1; depth <= 17; depth++)
                {
                    popup = fixture.Add("Other", owner: popup);
                    Assert.AreEqual(depth <= 16, fixture.Belongs(popup), "Ownership depth " + depth);
                }
            });
        }
    }
}
