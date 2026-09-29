using System;
using System.Collections;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass]
    public sealed class EditorProjectNavigationTests
    {
        [STATestMethod]
        public void DisposalReleasesLiveHooksAndAcceptsAlreadyDestroyedTrees()
        {
            foreach (bool destroyFirst in new[] { false, true })
            using (var fixture = new EditorNavigationFixture())
            {
                var navigation = new EditorProjectNavigation(fixture.Vbe, fixture.Dispatcher, module => { });
                Assert.AreEqual(1, ((IList)EditorNavigationFixture.Get(navigation, "hooks")).Count);
                if (destroyFirst) fixture.Tree.Dispose();
                navigation.Dispose();
                Assert.AreEqual(0, ((IList)EditorNavigationFixture.Get(navigation, "hooks")).Count);
                Assert.IsFalse(((System.Windows.Forms.Timer)EditorNavigationFixture.Get(navigation, "timer")).Enabled);
            }
        }

        [STATestMethod]
        public void DiscoveryFiltersProcessClassCaptionAndDeduplicatesNestedProjectTrees()
        {
            using (var fixture = new EditorNavigationFixture())
            {
                int process = Process.GetCurrentProcess().Id;
                CollectionAssert.AreEqual(new[] { fixture.Tree.Handle }, EditorProjectNavigation.FindProjectTrees(process, EditorNavigationFixture.Caption));
                Assert.AreEqual(0, EditorProjectNavigation.FindProjectTrees(-1, EditorNavigationFixture.Caption).Length);
                Assert.AreEqual(0, EditorProjectNavigation.FindProjectTrees(process, "missing caption").Length);
                Assert.AreEqual(0, EditorProjectNavigation.FindProjectTrees(process, null).Length);
            }
        }

        [STATestMethod]
        public void RoutingAndDoubleClickGuardsQueueOnlySupportedSelectedCodeComponents()
        {
            using (var fixture = new EditorNavigationFixture())
            {
                var contract = new EditorVbeContract(); fixture.Vbe.VBProjects.Add(contract.Project);
                int opens = 0; IEditorModule opened = null;
                using (var navigation = new EditorProjectNavigation(fixture.Vbe, fixture.Dispatcher, module => { opens++; opened = module; }))
                {
                    fixture.Vbe.Windows.Add(new EditorNavigationFixture.ProjectWindow { Type = 0, Caption = "ignored" });
                    fixture.Vbe.Windows.Add(new EditorNavigationFixture.ProjectWindow { Type = 6, Caption = EditorNavigationFixture.Caption });
                    EditorNavigationFixture.Invoke(navigation, "Refresh");
                    Assert.AreEqual(1, ((IList)EditorNavigationFixture.Get(navigation, "hooks")).Count);
                    foreach (int type in new[] { 1, 2, 100, 3, 0 })
                    {
                        contract.Original.Type = type; fixture.Vbe.Selection = contract.Original;
                        int before = opens; fixture.DoubleClick(); Application.DoEvents();
                        Assert.AreEqual(before + (type == 1 || type == 2 || type == 100 ? 1 : 0), opens);
                        if (opens > before) Assert.AreEqual("Project1 · Module1", opened.Name);
                    }
                    contract.Original.Type = 1;
                    int accepted = opens;
                    foreach (int refusal in new[] { 0, 1, 2 })
                    {
                        fixture.Replies.Flags = refusal == 0 ? 0u : 4u;
                        fixture.Replies.Item = refusal == 1 ? IntPtr.Zero : new IntPtr(1);
                        fixture.Replies.Selected = refusal == 2 ? new IntPtr(2) : new IntPtr(1);
                        fixture.DoubleClick(); Application.DoEvents(); Assert.AreEqual(accepted, opens);
                    }
                    fixture.Vbe.Selection = null; Assert.IsFalse((bool)EditorNavigationFixture.Invoke(navigation, "Route"));
                    fixture.Vbe.FailSelection = true; Assert.IsFalse((bool)EditorNavigationFixture.Invoke(navigation, "Route"));
                    fixture.Vbe.FailSelection = false; fixture.Vbe.Selection = new object(); Assert.IsFalse((bool)EditorNavigationFixture.Invoke(navigation, "Route"));
                    fixture.Vbe.Selection = contract.Original; fixture.Dispatcher.Dispose(); Assert.IsFalse((bool)EditorNavigationFixture.Invoke(navigation, "Route"));
                    fixture.Tree.Dispose(); EditorNavigationFixture.Invoke(navigation, "Refresh");
                    Assert.AreEqual(0, ((IList)EditorNavigationFixture.Get(navigation, "hooks")).Count);
                    var timer = (System.Windows.Forms.Timer)EditorNavigationFixture.Get(navigation, "timer");
                    typeof(System.Windows.Forms.Timer).GetMethod("OnTick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(timer, new object[] { EventArgs.Empty });
                }
            }
            using (var dispatcher = new Form()) using (var invalid = new EditorProjectNavigation(new object(), dispatcher, module => { })) { }
        }

        [STATestMethod]
        public void ForeignOwnedStaTreesAreDiscoveredButNeverSubclassedByThisThread()
        {
            using (var ready = new ManualResetEvent(false)) using (var finished = new ManualResetEvent(false))
            {
                Exception error = null;
                var thread = new Thread(() => { try { using (var fixture = new EditorNavigationFixture()) { ready.Set(); while (!finished.WaitOne(10)) Application.DoEvents(); } } catch (Exception failure) { error = failure; ready.Set(); } });
                thread.SetApartmentState(ApartmentState.STA); thread.Start();
                try
                {
                    Assert.IsTrue(ready.WaitOne(10000)); if (error != null) throw error;
                    using (var dispatcher = new Form())
                    {
                        var host = new EditorNavigationFixture.Host(); host.Windows.Add(new EditorNavigationFixture.ProjectWindow { Type = 6, Caption = EditorNavigationFixture.Caption });
                        using (var navigation = new EditorProjectNavigation(host, dispatcher, module => { }))
                            Assert.AreEqual(0, ((IList)EditorNavigationFixture.Get(navigation, "hooks")).Count);
                    }
                }
                finally { finished.Set(); Assert.IsTrue(thread.Join(10000)); }
                if (error != null) throw error;
            }
        }
    }
}
