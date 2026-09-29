using System;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeNavigationRejectsChangingOwnersAndDeduplicatesReparentedToolboxes()
        {
            foreach (string fault in new[] { "top French", "reparented", "foreign child", "closed root", "empty toolbox", "detached parent" })
            using (var host = new AutomationHost(new AutomationNode { Kind = System.Windows.Automation.ControlType.Pane, Name = "surface" }))
            using (var scene = new SystemScene())
            {
                var root = scene.Add("VBE", "wndclass_desked_gsk");
                var surface = scene.Add("Boîte à outils", "F3 MinFrame fixture", fault == "top French" ? null : root); surface.Handle = host.Handle;
                scene.AccessibilityHResult = -1;
                if (fault == "foreign child") surface.ProcessId = 999999;
                if (fault == "reparented") { var enumerate = VbeDebugWindows.EnumChildWindows; VbeDebugWindows.EnumChildWindows = (h, c, p) => { bool result = enumerate(h, c, p); surface.Parent = IntPtr.Zero; return result; }; }
                if (fault == "closed root") { var read = VbeDebugWindows.GetWindowThreadProcessId; int rootReads = 0; surface.ProcessId = 0; VbeDebugWindows.GetWindowThreadProcessId = (IntPtr h, out uint pid) => { uint result = read(h, out pid); if (h == root.Handle && ++rootReads > 1) pid = 0; return result; }; }
                if (fault == "empty toolbox") host.Invoke(h => System.Windows.Forms.Control.FromHandle(h).GetType().GetProperty("ControlBox").SetValue(System.Windows.Forms.Control.FromHandle(h), false));
                if (fault == "detached parent") host.Root.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.TreeItem, Name = "detached" });
                var probe = Native<VbeDebugWindows.INavigationSurfaceProbe>("NativeNavigationSurfaceProbe");
                if (fault == "detached parent") probe.GetType().GetField("ReadParent", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(probe, new Func<System.Windows.Automation.AutomationElement, System.Windows.Automation.AutomationElement>(e => null));
                var state = probe.Read("toolbox");
                if (fault == "foreign child" || fault == "closed root" || fault == "empty toolbox") Assert.IsFalse(state.Available, fault);
                else Assert.IsTrue(state.Available, fault);
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void NativeNavigationOwnsExactSurfaceAndReportsEveryAbsentOrInaccessibleState()
        {
            foreach (string fault in new[] { "missing root", "enumeration", "hidden", "wrong class", "ambiguous", "empty", "toolbox", "macros" })
            using (var host = new AutomationHost(new AutomationNode { Kind = System.Windows.Automation.ControlType.Pane, Name = "surface" }))
            using (var scene = new SystemScene())
            {
                var root = scene.Add("VBE", "wndclass_desked_gsk");
                var surface = scene.Add("Toolbox", fault == "toolbox" ? "VbaWindow" : fault == "macros" ? "#32770" : "SysTreeView32", root); surface.Handle = host.Handle;
                if (fault == "missing root") root.Class = "not VBE";
                if (fault == "enumeration") VbeDebugWindows.EnumChildWindows = (p, c, a) => { throw new InvalidOperationException("enumeration unavailable"); };
                if (fault == "hidden") surface.Visible = false;
                if (fault == "wrong class") surface.Class = "unrelated";
                if (fault == "ambiguous") scene.Add("second", "SysTreeView32", root);
                if (fault == "macros") { surface.Parent = IntPtr.Zero; surface.Text = "unrelated dialog"; }
                if (fault == "toolbox") surface.Handle = new IntPtr(987654);
                var probe = Native<VbeDebugWindows.INavigationSurfaceProbe>("NativeNavigationSurfaceProbe");
                var result = probe.Read(fault == "toolbox" ? "toolbox" : fault == "macros" ? "macros" : "project");
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(result.Available, fault);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(string.IsNullOrEmpty(result.Error), fault);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<InvalidOperationException>(() => probe.Act("project", "expired", "select"));
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void NativeNavigationReadsAndActsOnOwnedUiaTreeTabsAndMacroItems()
        {
            foreach (string pane in new[] { "project", "toolbox", "macros" })
            {
                var rootNode = new AutomationNode { Kind = System.Windows.Automation.ControlType.Pane, Name = "root" };
                var parent = rootNode.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.TreeItem, Name = "parent" }.With(System.Windows.Automation.SelectionItemPattern.Pattern, System.Windows.Automation.ExpandCollapsePattern.Pattern));
                parent.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.TreeItem, Name = "child" });
                rootNode.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.ListItem, Name = "item" }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
                rootNode.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.TabItem, Name = "tab" });
                rootNode.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.Button, Name = "button" });
                using (var host = new AutomationHost(rootNode)) using (var scene = new SystemScene())
                {
                    var root = scene.Add("VBE", "wndclass_desked_gsk");
                    var surface = scene.Add(pane == "macros" ? "Macros" : "Boîte à outils", pane == "project" ? "SysTreeView32" : pane == "toolbox" ? "F3 MinFrame fixture" : "#32770", pane == "macros" ? null : root); surface.Handle = host.Handle;
                    scene.AccessibilityHResult = -1;
                    if (pane == "toolbox") { var hidden = scene.Add("Toolbox", "F3 MinFrame fixture"); hidden.Visible = false; var foreign = scene.Add("Toolbox", "F3 MinFrame fixture"); foreign.ProcessId = 999999; scene.Add("wrong", "F3 MinFrame fixture"); scene.Add("Toolbox", "unrelated"); }
                    var probe = Native<VbeDebugWindows.INavigationSurfaceProbe>("NativeNavigationSurfaceProbe"); var state = probe.Read(pane);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(state.Available, pane);
                    var tree = System.Linq.Enumerable.First(state.Nodes, n => n.Name == "parent");
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(tree.Token, System.Linq.Enumerable.First(state.Nodes, n => n.Name == "child").ParentToken);
                    probe.Act(pane, tree.Token, "collapse"); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(parent.Expanded);
                    probe.Act(pane, tree.Token, "expand"); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(parent.Expanded);
                    probe.Act(pane, tree.Token, "select"); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(parent.Selected);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(3, scene.Pauses);
                    if (pane == "project")
                    {
                        dynamic snapshot = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = pane });
                        dynamic changed = VbeDebugWindows.ChangeNavigationSurface(new Request { Pane = pane, Action = "select", Control = tree.Token, ExpectedWindowVersion = snapshot.WindowVersion });
                        Assert.IsTrue((bool)changed.Verified);
                    }
                }
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void NativeNavigationRejectsDuplicateOrOversizedUiaAndKeepsUnreadableNodeErrors()
        {
            foreach (string fault in new[] { "duplicate", "oversized", "pattern", "name", "toolbox name" })
            {
                var provider = new AutomationNode { Kind = System.Windows.Automation.ControlType.Pane, Name = "root" };
                var node = provider.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.TreeItem, Name = "node" }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
                if (fault == "duplicate") { node.RuntimeIdentity = 500; provider.Add(new AutomationNode { Kind = node.Kind, Name = "duplicate", RuntimeIdentity = 500 }); }
                if (fault == "oversized") for (int i = 0; i < 4096; i++) provider.Add(new AutomationNode { Kind = node.Kind, Name = "node" + i });
                if (fault == "pattern") node.FailSelection = true;
                if (fault.EndsWith("name", StringComparison.Ordinal)) node.FailName = true;
                using (var host = new AutomationHost(provider)) using (var scene = new SystemScene())
                {
                    var root = scene.Add("VBE", "wndclass_desked_gsk"); var surface = scene.Add("Toolbox", fault == "toolbox name" ? "VbaWindow" : "SysTreeView32", root); surface.Handle = host.Handle; scene.AccessibilityHResult = -1;
                    var state = Native<VbeDebugWindows.INavigationSurfaceProbe>("NativeNavigationSurfaceProbe").Read(fault == "toolbox name" ? "toolbox" : "project");
                    if (fault == "pattern") { Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(state.Available); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(string.IsNullOrEmpty(state.Nodes[0].Error)); }
                    else Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(state.Available, fault);
                }
            }
        }
    }
    public sealed partial class NavigationSurfaceTests
    {
        [TestMethod]
        public void NullNativeNamesRemainFilterableAndUnavailableSurfaceCannotReceiveMutation()
        {
            var probe = new IdeSurfaceFixture.NavigationProbe(); probe.State.Nodes[0].Name = null;
            dynamic filtered = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project", Query = "absent" }, probe);
            Assert.AreEqual(0, (int)filtered.Total);
            var request = probe.Change(); probe.State.Available = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeNavigationSurface(request, probe)); Assert.AreEqual(0, probe.Actions);
        }
    }
    [TestClass, TestCategory("Unit")]
    public sealed partial class NavigationSurfaceTests
    {
        /// <summary>La pagination ne supprime ni la version complète ni les états d'inaccessibilité.</summary>
        [TestMethod]
        public void ReadingNavigationFiltersAndPagesWithoutCertifyingHiddenNodesAbsent()
        {
            var p = new IdeSurfaceFixture.NavigationProbe();
            dynamic all = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project" }, p);
            dynamic page = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project", Query = "module", Offset = 0, Limit = 1 }, p);
            Assert.AreEqual((string)all.WindowVersion, (string)page.WindowVersion);
            Assert.AreEqual(1, (int)page.Total); Assert.AreEqual("2", ((VbeDebugWindows.NavigationNode[])page.Nodes)[0].Token);
            dynamic skipped = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project", Offset = 1, Limit = 1 }, p);
            Assert.AreEqual(1, ((VbeDebugWindows.NavigationNode[])skipped.Nodes).Length);
            Assert.IsTrue((bool)((dynamic)VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project", Limit = 1 }, p)).HasMore);
            p.State.Available = false; p.State.Nodes = new VbeDebugWindows.NavigationNode[0]; p.State.Error = "provider unavailable";
            dynamic missing = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "toolbox" }, p);
            Assert.IsFalse((bool)missing.Available); Assert.IsNull((object)missing.WindowVersion); Assert.AreEqual("provider unavailable", (string)missing.Error);
        }

        /// <summary>Les paramètres et instantanés invalides sont refusés avant toute action.</summary>
        [TestMethod]
        public void NavigationRejectsInvalidRequestsAndMalformedSnapshots()
        {
            foreach (int i in Enumerable.Range(0, 9))
            {
                var p = new IdeSurfaceFixture.NavigationProbe(); var r = new Request { Pane = "project" };
                if (i == 0) r = null; if (i == 1) r.Pane = "other"; if (i == 2) r.Offset = -1; if (i == 3) r.Offset = 100001;
                if (i == 4) r.Limit = -1; if (i == 5) r.Limit = 501; if (i == 6) r.Query = new string('a', 257);
                if (i == 7) r.Pane = null; if (i == 8) r.Pane = "";
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ReadNavigationSurface(r, p)); Assert.AreEqual(0, p.Actions);
            }
            foreach (int i in Enumerable.Range(0, 6))
            {
                var p = new IdeSurfaceFixture.NavigationProbe();
                if (i == 0) p.State = null; if (i == 1) p.State.Nodes = null;
                if (i == 2) p.State.Nodes = Enumerable.Repeat(p.State.Nodes[0], 4097).ToArray();
                if (i == 3) p.State.Nodes[0] = null; if (i == 4) p.State.Nodes[0].Token = ""; if (i == 5) p.State.Identity = null;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project" }, p));
            }
        }

        /// <summary>Les trois actions relisent l'état et ne togglent pas un état déjà demandé.</summary>
        [TestMethod]
        public void SelectionAndExpansionAreIdempotentAndVerified()
        {
            foreach (string action in new[] { "expand", "collapse", "select" })
            {
                var p = new IdeSurfaceFixture.NavigationProbe(); p.State.Nodes[0].Expansion = action == "collapse" ? "Expanded" : "Collapsed";
                dynamic first = VbeDebugWindows.ChangeNavigationSurface(p.Change(action), p);
                Assert.IsTrue((bool)first.Verified); Assert.IsTrue((bool)first.Applied); Assert.AreEqual(1, p.Actions);
                dynamic second = VbeDebugWindows.ChangeNavigationSurface(p.Change(action), p);
                Assert.IsTrue((bool)second.Verified); Assert.IsFalse((bool)second.Applied); Assert.AreEqual(1, p.Actions);
            }
        }

        /// <summary>Les identités périmées, absentes, ambiguës et inaccessibles ne sont jamais livrées.</summary>
        [TestMethod]
        public void NavigationChecksEveryIdentityAndStateGuardBeforeDelivery()
        {
            foreach (int i in Enumerable.Range(0, 12))
            {
                var p = new IdeSurfaceFixture.NavigationProbe(); var r = p.Change();
                if (i == 0) r.Control = null; if (i == 1) r.ExpectedWindowVersion = null; if (i == 2) r.Action = "execute";
                if (i == 3) r.Control = new string('x', 257); if (i == 4) r.ExpectedWindowVersion = "stale";
                if (i == 5) r.Control = "absent";
                if (i == 6) { p.State.Nodes = new[] { p.State.Nodes[0], p.State.Nodes[0] }; r = p.Change(); }
                if (i == 7) { p.State.Nodes[0].Enabled = false; r = p.Change(); }
                if (i == 8) { p.State.Nodes[0].Error = "unreadable"; r = p.Change(); }
                if (i == 9) { p.State.Nodes[0].Expansion = "LeafNode"; r = p.Change(); }
                if (i == 10) { p.State.Nodes[0].Selected = null; r = p.Change("select"); }
                if (i == 11) { r.Pane = "toolbox"; }
                if (i < 4) Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeNavigationSurface(r, p));
                else Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeNavigationSurface(r, p));
                Assert.AreEqual(0, p.Actions);
            }
        }

        /// <summary>Une livraison ou relecture défaillante reste explicitement incertaine.</summary>
        [TestMethod]
        public void NavigationPreservesUncertaintyAfterActionOrReadbackFailure()
        {
            foreach (int i in Enumerable.Range(0, 6))
            {
                var p = new IdeSurfaceFixture.NavigationProbe(); var r = p.Change();
                p.AfterAction = () => {
                    if (i == 0) throw new InvalidOperationException("native failed");
                    if (i == 1) p.State.Identity = "new tree";
                    if (i == 2) p.State.Nodes = new VbeDebugWindows.NavigationNode[0];
                    if (i == 3) p.State.Nodes[0].Expansion = "Collapsed";
                    if (i == 4) p.State.Nodes[0].Error = "read failed";
                    if (i == 5) p.BeforeRead = () => { throw new InvalidOperationException("read failed"); };
                };
                dynamic result = VbeDebugWindows.ChangeNavigationSurface(r, p);
                Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.VerificationPending); Assert.AreEqual(1, p.Actions);
            }
        }
    }
}
