using System;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
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
        }
    }
}