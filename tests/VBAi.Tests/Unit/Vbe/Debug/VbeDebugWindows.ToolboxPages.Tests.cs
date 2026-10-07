using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeToolboxPagesValidateOwnerWindowServersAndFullReadback()
        {
            foreach (string fault in new[] { "none", "zero owner", "foreign", "hidden", "class", "caption", "no servers", "too many", "foreign server", "hidden server", "server class", "provider error", "provider null", "wrong role", "duplicate groups", "changed owner", "changed server owner", "changed visible", "changed server visible", "changed caption", "changed server class" })
                using (var scene = new SystemScene())
                {
                    var editor = scene.Add("VBE", "wndclass_desked_gsk"); var window = scene.Add("Toolbox", "VbaWindow", editor); var server = scene.Add("pages", "F3 Server fixture", window);
                    var grouping = new AccessibleNode { NativeRole = System.Windows.Forms.AccessibleRole.Grouping };
                    var list = new AccessibleNode { NativeRole = System.Windows.Forms.AccessibleRole.PageTabList };
                    list.Children.Add(new AccessibleNode { Label = "Controls", NativeRole = System.Windows.Forms.AccessibleRole.PageTab, StateValue = System.Windows.Forms.AccessibleStates.Selected }); grouping.Children.Add(list); server.Accessible = grouping;
                    uint owner = window.ProcessId;
                    if (fault == "zero owner") owner = 0;
                    if (fault == "foreign") window.ProcessId = 999999;
                    if (fault == "hidden") window.Visible = false;
                    if (fault == "class") window.Class = "unrelated";
                    if (fault == "caption") window.Text = "unrelated";
                    if (fault == "no servers") scene.Windows.Remove(server);
                    if (fault == "too many") for (int i = 0; i < 64; i++) scene.Add("server", "F3 Server fixture", window);
                    if (fault == "foreign server") server.ProcessId = 999999;
                    if (fault == "hidden server") server.Visible = false;
                    if (fault == "server class") server.Class = "unrelated";
                    if (fault == "provider error") scene.AccessibilityHResult = -1;
                    if (fault == "provider null") server.Accessible = null;
                    if (fault == "wrong role") grouping.NativeRole = System.Windows.Forms.AccessibleRole.Client;
                    if (fault == "duplicate groups") scene.Add("second", "F3 Server fixture", window).Accessible = grouping;
                    list.Children[0].OnName = () =>
                    {
                        if (fault == "changed owner") window.ProcessId = 999999;
                        if (fault == "changed server owner") server.ProcessId = 999999;
                        if (fault == "changed visible") window.Visible = false;
                        if (fault == "changed server visible") server.Visible = false;
                        if (fault == "changed caption") window.Text = "Changed";
                        if (fault == "changed server class") server.Class = "changed";
                    };
                    var result = (VbeDebugWindows.NavigationSurface)Call("ReadNativeToolboxPages", window.Handle, owner);
                    Assert.AreEqual(fault == "none", result.Available, fault);
                    if (fault == "none") { Assert.AreEqual(1, result.Nodes.Length); Assert.AreEqual(true, result.Nodes[0].ObservedSelected); Assert.IsNull(result.Nodes[0].Selected); }
                    else { Assert.AreEqual(0, result.Nodes.Length); Assert.IsNull(result.Identity); Assert.IsFalse(string.IsNullOrEmpty(result.Error)); }
                    if (fault == "none") Assert.IsTrue(Native<VbeDebugWindows.INavigationSurfaceProbe>("NativeNavigationSurfaceProbe").Read("toolbox").Available);
                }
        }

        [TestMethod]
        public void NativeToolboxAdapterReturnsSimpleChildrenAndReleasesOwnedNativeInterfaces()
        {
            using (var scene = new SystemScene())
            {
                var node = new AccessibleNode { NativeRole = System.Windows.Forms.AccessibleRole.Grouping };
                node.Children.Add(null);
                var type = typeof(VbeDebugWindows).GetNestedType("ToolboxAccessibleNode", System.Reflection.BindingFlags.NonPublic);
                using (var adapter = (VbeDebugWindows.IToolboxAccessibleNode)Activator.CreateInstance(type, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, new object[] { node }, null))
                    Assert.IsNull(adapter.Child(1));
            }
            object font = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
            IntPtr identity = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(font);
            try
            {
                foreach (bool native in new[] { false, true })
                {
                    object value = native ? System.Runtime.InteropServices.Marshal.GetUniqueObjectForIUnknown(identity) : new object();
                    var node = new ChildValueAccessible { ChildValue = value };
                    var type = typeof(VbeDebugWindows).GetNestedType("ToolboxAccessibleNode", System.Reflection.BindingFlags.NonPublic);
                    using (var adapter = (VbeDebugWindows.IToolboxAccessibleNode)Activator.CreateInstance(type, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, new object[] { node }, null)) Assert.IsNull(adapter.Child(1));
                }
                using (var scene = new SystemScene())
                {
                    scene.OverrideAccessibility = true; scene.AccessibilityOverride = System.Runtime.InteropServices.Marshal.GetUniqueObjectForIUnknown(identity);
                    var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call("OpenToolboxAccessibleNode", new IntPtr(100)));
                    Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
                }
            }
            finally { System.Runtime.InteropServices.Marshal.Release(identity); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(font); }
            var readNative = VbeDebugWindows.AccessibleObjectFromWindow;
            using (var window = new System.Windows.Forms.Form())
            {
                Guid iid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71"); object accessible;
                Assert.AreEqual(0, readNative(window.Handle, 0xfffffffc, ref iid, out accessible));
                using (var scene = new SystemScene())
                {
                    scene.OverrideAccessibility = true; scene.AccessibilityOverride = accessible;
                    using (var adapter = (VbeDebugWindows.IToolboxAccessibleNode)Call("OpenToolboxAccessibleNode", new IntPtr(100))) Assert.IsTrue(adapter.Role(0) != 20);
                }
            }
        }
    }
    /// <summary>Matrice de lecture, identités, bornes, défaillances et refus des actions MSAA.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ToolboxPagesTests
    {
        private static VbeDebugWindows.NavigationSurface Read(ToolboxPagesFixture root, long window = 100, long server = 200) =>
            VbeDebugWindows.ReadToolboxPages(new IntPtr(window), new IntPtr(server), "Boîte à outils", root);

        /// <summary>Conserve l'ordre, l'état observé et des identités indépendantes des noms.</summary>
        [TestMethod]
        public void PagesExposeObservedSelectionButNoQualifiedActionOrButtons()
        {
            var root = ToolboxPagesFixture.Create(); var state = Read(root);
            Assert.IsTrue(state.Available); Assert.AreEqual("MSAA", state.Provider); Assert.AreEqual(false, state.ButtonsExposed);
            Assert.AreEqual(2, state.Nodes.Length); Assert.AreEqual("Contrôles", state.Nodes[0].Name);
            Assert.AreEqual("MSAA.PageTab", state.Nodes[0].Kind); Assert.AreEqual(true, state.Nodes[0].ObservedSelected);
            Assert.AreEqual(2097154, state.Nodes[0].MSAAState); Assert.AreEqual(0, state.MSAAContainerState);
            Assert.AreEqual(false, state.Nodes[1].ObservedSelected); Assert.IsNull(state.Nodes[0].Selected);
            Assert.IsTrue(state.Nodes[0].ActionUnavailableReason.Contains("not qualified"));
            Assert.AreEqual("msaa:100:200:0.1.1", state.Nodes[0].Token);
            Assert.AreEqual(state.Nodes[0].Token, Read(root).Nodes[0].Token);
            Assert.AreNotEqual(state.Nodes[0].Token, Read(root, 101).Nodes[0].Token);
            Assert.AreNotEqual(state.Nodes[0].Token, Read(root, 100, 201).Nodes[0].Token);
            root.Children[0].Children[1].NameValue = "Contrôles";
            Assert.AreNotEqual(Read(root).Nodes[0].Token, Read(root).Nodes[1].Token);
            root.Children.Add(null); Assert.AreEqual(2, Read(root).Nodes.Length);
            root.Children[0].Children[0].StateValue = 1;
            Assert.IsFalse(Read(root).Nodes[0].Enabled);
        }

        /// <summary>La version couvre l'ensemble des pages et le refus ne livre aucune action.</summary>
        [TestMethod]
        public void PageVersionsCoverObservedStateAndRejectSelectionBeforeDelivery()
        {
            var root = ToolboxPagesFixture.Create(); var probe = new ToolboxPagesFixture.Probe { State = Read(root) };
            dynamic all = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox" }, probe);
            dynamic page = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox", Query = "Contrôles", Limit = 1 }, probe);
            Assert.AreEqual((string)all.WindowVersion, (string)page.WindowVersion); Assert.AreEqual(1, (int)page.Total);
            Assert.AreEqual("MSAA", (string)page.Provider); Assert.AreEqual(false, (bool?)page.ButtonsExposed);
            var request = new Request { Pane = "toolbox", Action = "select", Control = probe.State.Nodes[0].Token, ExpectedWindowVersion = all.WindowVersion };
            var error = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeNavigationSurface(request, probe));
            Assert.IsTrue(error.Message.Contains("not qualified")); Assert.AreEqual(0, probe.Actions);
            root.Children[0].Children[0].StateValue = 2097152; probe.State = Read(root);
            dynamic changed = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox" }, probe);
            Assert.AreNotEqual((string)all.WindowVersion, (string)changed.WindowVersion);
            root.Children[0].Children[0].StateValue |= 4; probe.State = Read(root);
            dynamic focusChanged = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox" }, probe);
            Assert.AreNotEqual((string)changed.WindowVersion, (string)focusChanged.WindowVersion);
            root.Children[0].StateValue = 4; probe.State = Read(root);
            dynamic containerChanged = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox" }, probe);
            Assert.AreNotEqual((string)focusChanged.WindowVersion, (string)containerChanged.WindowVersion);
        }

        /// <summary>Refuse les identités invalides, arbres incomplets, listes multiples et états illisibles.</summary>
        [TestMethod]
        public void InvalidAndAmbiguousPageTreesAreRefusedCompletely()
        {
            for (int scenario = 0; scenario < 20; scenario++)
            {
                var root = ToolboxPagesFixture.Create(); long window = 100, server = 200;
                if (scenario == 0) window = 0; if (scenario == 1) server = 0;
                if (scenario == 2) root = null;
                if (scenario == 3) root.RoleValue = 16;
                if (scenario == 4) root.Children.Clear();
                if (scenario == 5) root.Children.Add(root.Children[0]);
                if (scenario == 6) root.Children[0].Children.Clear();
                if (scenario == 7) root.Children[0].CountOverride = 129;
                if (scenario == 8) root.Children[0].Children[0].RoleValue = 43;
                if (scenario == 9) root.Children[0].Children[0].NameValue = "";
                if (scenario == 10) root.Children[0].Children[0].NameValue = new string('a', 257);
                if (scenario == 11) root.Children[0].Children[0].StateValue = -1;
                if (scenario == 12) root.FailCount = true;
                if (scenario == 13) root.FailChild = true;
                if (scenario == 14) root.Children[0].FailName = true;
                if (scenario == 15) root.Children[0].FailState = true;
                if (scenario == 16) root.FailRole = true;
                if (scenario == 17) root.CountOverride = -2;
                if (scenario == 18) root.Children[0].StateValue = -1;
                if (scenario == 19) { root.Children.Clear(); root.Children.Add(null); }
                Assert.ThrowsException<InvalidOperationException>(() => Read(root, window, server), "scenario " + scenario);
            }
        }

        /// <summary>La profondeur et le volume ne sont pas tronqués silencieusement.</summary>
        [TestMethod]
        public void TreeDepthAndVolumeHaveStrictBounds()
        {
            var root = new ToolboxPagesFixture(); var branch = root;
            for (int i = 0; i < 9; i++) { var child = new ToolboxPagesFixture(16); branch.Children.Add(child); branch = child; }
            Assert.ThrowsException<InvalidOperationException>(() => Read(root));
            root = ToolboxPagesFixture.Create();
            for (int i = 0; i < 5; i++)
            {
                branch = new ToolboxPagesFixture(16);
                for (int j = 0; j < 128; j++) branch.Children.Add(new ToolboxPagesFixture(16));
                root.Children.Add(branch);
            }
            Assert.ThrowsException<InvalidOperationException>(() => Read(root));
            root = ToolboxPagesFixture.Create(); root.CountOverride = 129;
            Assert.ThrowsException<InvalidOperationException>(() => Read(root));
            root = new ToolboxPagesFixture();
            for (int i = 0; i < 3; i++) { branch = new ToolboxPagesFixture(16); for (int j = 0; j < 128; j++) branch.Children.Add(new ToolboxPagesFixture(16)); root.Children.Add(branch); }
            branch = new ToolboxPagesFixture(60); for (int i = 0; i < 128; i++) branch.Children.Add(new ToolboxPagesFixture(37) { NameValue = "Page" }); root.Children.Add(branch);
            Assert.ThrowsException<InvalidOperationException>(() => Read(root));
        }
    }
}
