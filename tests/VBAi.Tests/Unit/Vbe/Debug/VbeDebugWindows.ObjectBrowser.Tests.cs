namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using System.Windows.Automation;
    using VBAi;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void BrowserReadRequiresVisibleSupportedPaneAndReturnsNativePatternsAndTextLimits()
        {
            using (var scene = new ObjectBrowserScene())
            {
                scene.Root.Visible = false;
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ReadObjectBrowser()).Available);
                scene.Root.Visible = true; scene.Pane.Visible = false;
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ReadObjectBrowser()).Available);
                scene.Pane.Visible = true;
                var list = BrowserList("Classes Alpha"); list.Children[0].Selected = true;
                scene.Add(list);
                var plain = BrowserList("Classes Plain"); plain.Patterns.Clear(); scene.Add(plain);
                var combo = BrowserLibraries("Libraries", "Excel"); combo.Text = "Excel"; scene.Add(combo, "ComboBox");
                var failing = BrowserLibraries("bad"); failing.FailValue = true; scene.Add(failing);
                scene.Add(new AutomationNode { Kind = ControlType.Document, Name = "description", Text = "  exact text  " }.With(TextPattern.Pattern));
                scene.Add(new AutomationNode { Kind = ControlType.Document, Name = "long", Text = new string('x', 16385) }.With(TextPattern.Pattern));
                scene.Add(new AutomationNode { Kind = ControlType.Document, Name = "unsupported" });
                scene.Add(new AutomationNode { Kind = ControlType.Button, Name = "unrelated" });
                var unavailable = scene.Add(BrowserList()); scene.Unreadable.Add(unavailable.Handle);
                dynamic result = VbeDebugWindows.ReadObjectBrowser();
                Assert.IsTrue((bool)result.Available); Assert.IsFalse((bool)result.SemanticResolutionVerified);
                object[] descriptions = ((IEnumerable)result.Descriptions).Cast<object>().ToArray();
                Assert.AreEqual(3, descriptions.Length);
                Assert.AreEqual("exact text", (string)((dynamic)descriptions[0]).Text);
                Assert.IsFalse((bool)((dynamic)descriptions[0]).Truncated);
                Assert.AreEqual(16384, ((string)((dynamic)descriptions[1]).Text).Length);
                Assert.IsTrue((bool)((dynamic)descriptions[1]).Truncated);
                Assert.AreEqual("Text pattern unavailable.", (string)((dynamic)descriptions[2]).Error);
                object[] selections = ((IEnumerable)result.Selections).Cast<object>().ToArray();
                Assert.AreEqual(4, selections.Length);
                CollectionAssert.AreEqual(new[] { "Classes Alpha" }, (string[])((dynamic)selections[0]).Selected);
                Assert.IsNull(((dynamic)selections[1]).Selected);
                Assert.AreEqual("Excel", (string)((dynamic)selections[2]).Value);
                Assert.IsNotNull(((dynamic)selections[3]).Error);
            }
        }
        [TestMethod]
        public void BrowserSelectionRejectsEveryInvalidInputBeforeNativeMutation()
        {
            foreach (var request in new[] {
                new Request(), new Request { ObjectName=" " }, new Request { ObjectName=new string('a',256) },
                new Request { Context="Excel", Procedure="member" }, new Request { ObjectName="Class", Procedure=" " },
                new Request { ObjectName="Class", Procedure=new string('b',256) }, new Request { Context=" " },
                new Request { Context=new string('c',256) } })
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectObjectBrowser(request));
            using (var scene = new ObjectBrowserScene())
            {
                var request = new Request { ObjectName = "Alpha" };
                scene.Root.Visible = false; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(request));
                scene.Root.Visible = true; scene.Pane.Visible = false; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(request));
                scene.Pane.Visible = true; scene.Root.Enabled = false; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(request));
                Assert.AreEqual(0, scene.Notifications.Count);
            }
        }
        [TestMethod]
        public void BrowserListsClassesMembersAndLibrariesWithFilteringPaginationAndSelectionReadability()
        {
            using (var scene = new ObjectBrowserScene())
            {
                var classes = BrowserList("Classes Alpha", "Classes Beta", "Classes Gamma"); classes.Children[0].Selected = true;
                classes.Children[1].Patterns.Clear(); scene.Add(classes);
                scene.Add(BrowserList("Members of 'Alpha' Go"));
                scene.Add(BrowserList("Membres de 'Alpha' Aller"));
                scene.Add(BrowserList()); scene.Add(BrowserList("unrelated"));
                scene.Add(new AutomationNode { Kind = ControlType.Document });
                scene.Add(BrowserLibraries("other", "wrong"), "ComboBox");
                var libraries = BrowserLibraries("Bibliothèques", "Excel", "VBA"); scene.Add(libraries, "ComboBox");
                dynamic all = VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes" });
                Assert.IsTrue((bool)all.Available); Assert.AreEqual(3, (int)all.TotalMatches); Assert.AreEqual(50, (int)all.Limit);
                var rows = ((IEnumerable)all.Items).Cast<object>().ToArray();
                Assert.IsTrue((bool)((dynamic)rows[0]).Selected); Assert.IsFalse((bool)((dynamic)rows[1]).SelectionReadable);
                Assert.IsFalse((bool)((dynamic)rows[2]).Selected);
                dynamic page = VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes", Offset = 1, Limit = 1 });
                Assert.AreEqual("Classes Beta", (string)((dynamic)((IEnumerable)page.Items).Cast<object>().Single()).Label);
                Assert.IsTrue((bool)page.HasMore);
                dynamic filtered = VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes", Query = "gAM", Limit = 200 });
                Assert.AreEqual(1, (int)filtered.TotalMatches); Assert.IsFalse((bool)filtered.HasMore);
                Assert.AreEqual(2, (int)((dynamic)VbeDebugWindows.ListObjectBrowser(new Request { Pane = "libraries" })).TotalMatches);
                // Both language-specific member lists are deliberately ambiguous until narrowed.
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ListObjectBrowser(new Request { Pane = "members" })).Available);
            }
        }
        [TestMethod]
        public void BrowserListRejectsInvalidPageBoundsMissingPaneAndAmbiguousLists()
        {
            foreach (var request in new[] { new Request { Pane = "wrong" }, new Request { Pane = "classes", Offset = -1 }, new Request { Pane = "members", Limit = -1 }, new Request { Pane = "libraries", Limit = 201 } })
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ListObjectBrowser(request));
            using (var scene = new ObjectBrowserScene())
            {
                scene.Root.Visible = false; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes" }));
                scene.Root.Visible = true; scene.Pane.Visible = false; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes" }));
                scene.Pane.Visible = true;
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes" })).Available);
                scene.Add(BrowserList("Classes one")); scene.Add(BrowserList("Classes two"));
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes" })).Available);
            }
        }
        [TestMethod]
        public void BrowserSelectsExactClassAndLocalizedMemberAfterAsynchronousListChange()
        {
            foreach (bool french in new[] { false, true })
                using (var scene = new ObjectBrowserScene())
                {
                    var classes = BrowserList("Classes Alpha"); scene.Add(classes);
                    var members = BrowserList(); scene.Add(members);
                    scene.Add(BrowserLibraries("other"), "ComboBox");
                    scene.OnPause = count => { if (count == 2) members.Add(new AutomationNode { Name = french ? "Membres de 'Alpha' Go" : "Members of 'Alpha' Go" }.With(SelectionItemPattern.Pattern)); };
                    dynamic result = VbeDebugWindows.SelectObjectBrowser(new Request { ObjectName = "Alpha", Procedure = "Go" });
                    Assert.IsTrue((bool)result.SelectionObserved); Assert.IsFalse((bool)result.SemanticResolutionVerified);
                    Assert.IsTrue(classes.Children[0].Selected); Assert.IsTrue(members.Children[0].Selected);
                    Assert.AreEqual(2, scene.Notifications.Count); Assert.AreEqual(scene.Pane.Handle, scene.Notifications[0].Item1);
                    Assert.AreEqual(new IntPtr(42 | (1 << 16)), scene.Notifications[0].Item3);
                }
        }
        [TestMethod]
        public void BrowserMemberTimeoutAndFinalSelectionReadbackNeverClaimSuccess()
        {
            using (var scene = new ObjectBrowserScene())
            {
                var classes = BrowserList("Classes Alpha"); scene.Add(classes);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { ObjectName = "Alpha", Procedure = "Missing" }));
                Assert.AreEqual(10, scene.Pauses); Assert.AreEqual(1, scene.Notifications.Count);
            }
            using (var scene = new ObjectBrowserScene())
            {
                var classes = BrowserList("Classes Alpha"); scene.Add(classes);
                scene.OnPause = count => classes.Children[0].Selected = false;
                Assert.IsFalse((bool)((dynamic)VbeDebugWindows.SelectObjectBrowser(new Request { ObjectName = "Alpha" })).SelectionObserved);
            }
        }
        [TestMethod]
        public void BrowserClassSelectionRejectsMissingAmbiguousNativeUnavailableAndNotificationFailures()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4, 5, 6 })
                using (var scene = new ObjectBrowserScene())
                {
                    var classes = BrowserList(scenario == 0 ? "Classes Other" : "Classes Alpha");
                    var window = scene.Add(classes);
                    if (scenario == 1) classes.Add(new AutomationNode { Name = "Classes Alpha" }.With(SelectionItemPattern.Pattern));
                    if (scenario == 2) classes.NativeHandle = 0;
                    if (scenario == 3) window.Parent = IntPtr.Zero;
                    if (scenario == 4) classes.Children[0].Patterns.Clear();
                    if (scenario == 5) classes.Children[0].SelectedAction = () => classes.Children[0].Selected = false;
                    if (scenario == 6) scene.PostSucceeded = false;
                    // Keep enumeration rooted at the browser while the native parent boundary refuses the control.
                    if (scenario == 3) { window.Parent = scene.Pane.Handle; VbeDebugWindows.ObjectBrowserParent = handle => handle == window.Handle ? IntPtr.Zero : scene.Native.Find(handle)?.Parent ?? IntPtr.Zero; }
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { ObjectName = "Alpha" }));
                    Assert.AreEqual(scenario == 6 ? 1 : 0, scene.Notifications.Count);
                }
        }
        [TestMethod]
        public void BrowserLibrarySelectionRequiresUniqueExactNativeComboAndSelectableItem()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4, 5, 6, 7 })
                using (var scene = new ObjectBrowserScene())
                {
                    var combo = BrowserLibraries("Libraries", scenario == 5 ? "Other" : "Excel");
                    NativeDebugScene.Window window = null;
                    if (scenario != 0) window = scene.Add(combo, scenario == 3 ? "Wrong" : "ComboBox");
                    if (scenario == 1) scene.Add(BrowserLibraries("Bibliothèques", "Excel"), "ComboBox");
                    if (scenario == 2) combo.NativeHandle = 0;
                    if (scenario == 4) window.Enabled = false;
                    if (scenario == 6) combo.Add(new AutomationNode { Name = "Excel" }.With(SelectionItemPattern.Pattern));
                    if (scenario == 7) combo.Children[0].Patterns.Clear();
                    scene.Add(new AutomationNode { Kind = ControlType.Document });
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { Context = "Excel" }));
                    Assert.AreEqual(0, scene.Notifications.Count);
                }
        }
        [TestMethod]
        public void BrowserLibrarySelectionVerifiesNotificationAndAccessibleReadback()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3 })
                using (var scene = new ObjectBrowserScene())
                {
                    var combo = BrowserLibraries("Libraries", "Excel", "VBA"); scene.Add(combo, "ComboBox");
                    if (scenario == 0) scene.PostSucceeded = false;
                    if (scenario == 1) combo.Patterns.Remove(SelectionPattern.Pattern.Id);
                    if (scenario == 2) combo.Children[0].SelectedAction = () => { combo.Children[0].Selected = false; combo.Children[1].Selected = true; };
                    if (scenario < 3) Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { Context = "Excel" }));
                    else
                    {
                        dynamic result = VbeDebugWindows.SelectObjectBrowser(new Request { Context = "Excel" });
                        Assert.IsTrue((bool)result.SelectionObserved); Assert.AreEqual("Excel", (string)result.Library);
                        Assert.AreEqual(1, combo.Children[0].SelectionCount);
                    }
                    Assert.AreEqual(1, scene.Notifications.Count);
                }
        }
        [TestMethod]
        public void BrowserFinalObservationSeparatesLibraryClassAndMemberReadbacks()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4 })
                using (var scene = new ObjectBrowserScene())
                {
                    var combo = BrowserLibraries("Libraries", "Excel"); scene.Add(combo, "ComboBox");
                    var classes = BrowserList("Classes Alpha"); scene.Add(classes);
                    var members = BrowserList("Members of 'Alpha' Go"); scene.Add(members);
                    scene.OnPause = count =>
                    {
                        if (count == 3)
                        {
                            if (scenario == 0) combo.Name = "other";
                            if (scenario == 1) combo.Patterns.Remove(SelectionPattern.Pattern.Id);
                            if (scenario == 2) combo.Children[0].Selected = false;
                            if (scenario == 3) classes.Children[0].Selected = false;
                            if (scenario == 4) members.Children[0].Selected = false;
                        }
                    };
                    dynamic result = VbeDebugWindows.SelectObjectBrowser(new Request { Context = "Excel", ObjectName = "Alpha", Procedure = "Go" });
                    Assert.IsFalse((bool)result.SelectionObserved);
                }
        }
    }
}
