using System;
using System.Linq;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NavigationSurfaceTests
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
