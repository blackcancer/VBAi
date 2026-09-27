using System;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsCoreBranchTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Control;
            public VbeForms Service;
            public string Version => (string)((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion;
            public Request Request(string path)
            {
                return new Request { Project = Project.Name, Form = Form.Name,
                    ControlPath = path, ExpectedTreeVersion = Version };
            }
        }

        private static Fixture Create(string type)
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture { Project = project, Form = form,
                Control = form.Designer.Controls.AddExisting(type, type + "1"),
                Service = new VbeForms(vbe) };
        }

        [TestMethod]
        public void MultiPageAddAndRemoveUseCanonicalPagePathAndIndex()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.Caption = "First";
            dynamic first = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("Controls/MultiPage1/Pages/PageA", (string)first.AddedPath);
            Assert.AreEqual("First", f.Control.Pages.Single().Caption);
            request.ExpectedTreeVersion = first.Tree.TreeVersion;
            request.NewName = "PageB";
            request.InsertIndex = 0;
            dynamic second = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("PageB", f.Control.Pages.First().Name);
            var remove = f.Request("Controls/MultiPage1/Pages/PageA");
            dynamic removed = f.Service.RemovePageOrTab(remove);
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual("PageB", f.Control.Pages.Single().Name);
            Assert.AreNotEqual((string)second.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void TabStripAddAndRemoveTrackNativeTabCollection()
        {
            var f = Create("TabStrip");
            var request = f.Request("Controls/TabStrip1");
            request.ParentPath = request.ControlPath;
            request.NewName = "TabA";
            dynamic added = f.Service.AddPageOrTab(request, "Tabs");
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Name);
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Caption);
            dynamic removed = f.Service.RemovePageOrTab(f.Request("Controls/TabStrip1/Tabs/TabA"));
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(0, f.Control.Tabs.Count);
            Assert.AreNotEqual((string)added.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void PageAddRejectsStaleVersionDuplicateIndexAndWrongParent()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.ExpectedTreeVersion = f.Version;
            request.InsertIndex = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.InsertIndex = null;
            f.Control.Pages.Add("PageA", "Existing");
            request.ExpectedTreeVersion = f.Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            Assert.AreEqual(1, f.Control.Pages.Count);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Tabs"));
        }

        [TestMethod]
        public void PageAddRollsBackAfterNativeAddFailureAndReportsFailedRollback()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            f.Control.Pages.FailAfterAdd = true;
            var clean = Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(clean.Message, "absent after rollback");
            Assert.AreEqual(0, f.Control.Pages.Count);
            f.Control.Pages.FailRemove = true;
            var incomplete = Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(incomplete.Message, "rollback could not be verified");
            Assert.AreEqual(1, f.Control.Pages.Count);
        }

        [TestMethod]
        public void NodeScalarMutationUsesVersionAndConvertsToDeclaredType()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Left";
            request.Value = "12.5";
            dynamic changed = f.Service.SetNodeProperty(request);
            Assert.AreEqual(12.5d, f.Control.Left);
            request.ExpectedTreeVersion = changed.Tree.TreeVersion;
            request.Property = "Caption";
            request.Value = "Updated";
            dynamic captioned = f.Service.SetNodeProperty(request);
            Assert.AreEqual("Updated", f.Control.Caption);
            Assert.AreNotEqual((string)changed.Tree.TreeVersion, (string)captioned.Tree.TreeVersion);
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
        }

        [TestMethod]
        public void NodeMutationRefusesCrashProneNativeSettersBeforeWriting()
        {
            foreach (var caseInfo in new[] {
                Tuple.Create("ComboBox", "ColumnCount", (object)2),
                Tuple.Create("TextBox", "ScrollBars", (object)2),
                Tuple.Create("SpinButton", "Min", (object)1),
                Tuple.Create("ToggleButton", "Value", (object)true),
                Tuple.Create("Label", "Cancel", (object)true) })
            {
                var f = Create(caseInfo.Item1);
                var request = f.Request("Controls/" + caseInfo.Item1 + "1");
                request.Property = caseInfo.Item2;
                request.Value = caseInfo.Item3;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request),
                    caseInfo.Item1 + "." + caseInfo.Item2);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
        }

        [TestMethod]
        public void NodeMutationRejectsManagedObjectMemberAndUnknownProperty()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Missing";
            request.Value = "x";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            request.Property = "Font.Name";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            Assert.AreEqual("Arial", f.Control.Font.Name);
        }
    }
}
