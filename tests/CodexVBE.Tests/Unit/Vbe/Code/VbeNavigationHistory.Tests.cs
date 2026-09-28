using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeNavigationHistoryTests
    {
        public sealed class NavigationSource
        {
            public string Code { get; set; } = "abc\ndef";
            public string Sha256 { get; set; } = "sha";
        }
        public sealed class NavigationHost
        {
            public List<VbeSessionTests.FakeProject> VBProjects { get; } = new List<VbeSessionTests.FakeProject>();
            public bool FailPane { get; set; }
            public NavigationPane Pane { get; set; }
            public NavigationPane ActiveCodePane { get { if (FailPane) throw new IOException("pane unavailable"); return Pane; } }
        }
        public sealed class NavigationPane
        {
            public NavigationModule CodeModule { get; set; }
            public void GetSelection(ref int line, ref int column, ref int endLine, ref int endColumn)
            { line = endLine = 1; column = endColumn = 1; }
        }
        public sealed class NavigationModule { public NavigationComponent Parent { get; set; } }
        public sealed class NavigationComponent { public string Name { get; set; } = "Active"; public NavigationCollection Collection { get; set; } }
        public sealed class NavigationCollection { public VbeSessionTests.FakeProject Parent { get; set; } }

        private static Request Location(string action = "go", string project = "P", string query = "mark")
        { return new Request { Action = action, Project = project, Query = query, Module = "M", ExpectedSha256 = "SHA", StartLine = 1, StartColumn = 1 }; }
        private static Response Execute(Request request)
        { return Response.Success(request.Command == "read_module" ? (object)new NavigationSource() : new object()); }
        private static IList Stack(VbeNavigationHistory history, string name)
        { return (IList)typeof(VbeNavigationHistory).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(history); }

        [TestMethod]
        public void SessionBookmarksValidateNamesScopeCapacityAndStableUnsavedIdentity()
        {
            var history = new VbeNavigationHistory(null, Execute);
            foreach (string selector in new[] { null, "", " " })
                Assert.ThrowsException<ArgumentException>(() => history.Bookmark(Location("list", selector)));
            foreach (string query in new[] { null, "", " ", new string('x', 101), "bad\nname" })
                Assert.ThrowsException<ArgumentException>(() => history.Bookmark(Location("add", query: query)));
            Assert.ThrowsException<ArgumentException>(() => history.Bookmark(Location("unknown")));
            Assert.ThrowsException<InvalidOperationException>(() => history.Bookmark(Location("go")));
            history.Bookmark(Location("remove"));
            for (int i = 0; i < 200; i++) history.Bookmark(Location("add", query: "mark" + i));
            Assert.ThrowsException<InvalidOperationException>(() => history.Bookmark(Location("add")));
            history.Bookmark(Location("add", query: "mark0"));
            history.Bookmark(Location("go", query: "MARK0"));
            history.Bookmark(Location("list")); history.Bookmark(Location("list", "Other"));
            history.Bookmark(Location("remove", query: "mark0"));

            var host = new NavigationHost();
            var project = new VbeSessionTests.FakeProject { Name = "P", FileName = "" };
            host.VBProjects.Add(project);
            history = new VbeNavigationHistory(host, Execute);
            history.Bookmark(Location("add")); history.Bookmark(Location("list")); history.Bookmark(Location("go"));
            project.ThrowFileName = true;
            history.Bookmark(Location("list")); history.Bookmark(Location("remove"));
            project.ThrowFileName = false; project.FileName = "relative.xlsm";
            history.Bookmark(Location("add")); history.Bookmark(Location("list"));
        }

        [TestMethod]
        public void PersistentBookmarksUseMacroIdentityAndRoundTripAcrossHistoryInstances()
        {
            string root = Path.Combine(Path.GetTempPath(), "CodexNavigation-" + Guid.NewGuid().ToString("N"));
            try
            {
                string database = Path.Combine(root, "bookmarks.db"), projectPath = Path.Combine(root, "macro.xlsm");
                var host = new NavigationHost();
                host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "P", FileName = projectPath });
                var history = new VbeNavigationHistory(host, Execute, database);
                foreach (string query in new[] { null, " ", new string('x', 101), "bad\tname" })
                    Assert.ThrowsException<ArgumentException>(() => history.Bookmark(Location("add", query: query)));
                history.Bookmark(Location("list")); history.Bookmark(Location("remove"));
                Assert.ThrowsException<InvalidOperationException>(() => history.Bookmark(Location("go")));
                Assert.ThrowsException<ArgumentException>(() => history.Bookmark(Location("unknown")));
                history.Bookmark(Location("add"));
                var reopened = new VbeNavigationHistory(null, Execute, database);
                reopened.Bookmark(Location("list", projectPath));
                reopened.Bookmark(Location("go", projectPath, "MARK"));
                reopened.Bookmark(Location("remove", projectPath));
                Assert.ThrowsException<InvalidOperationException>(() => reopened.Bookmark(Location("go", projectPath)));
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        [TestMethod]
        public void NavigationRejectsStaleOutOfRangeAndFailedReadsBeforeSelectingCode()
        {
            var history = new VbeNavigationHistory(null, Execute);
            Assert.ThrowsException<ArgumentException>(() => history.Go(Location(project: " ")));
            Assert.ThrowsException<ArgumentException>(() => history.Go(Location("invalid")));
            Assert.ThrowsException<InvalidOperationException>(() => history.Go(Location("back")));
            Assert.ThrowsException<InvalidOperationException>(() => history.Go(Location("forward")));
            foreach (string sha in new[] { null, "", "changed" })
            {
                var request = Location(); request.ExpectedSha256 = sha;
                Assert.ThrowsException<InvalidOperationException>(() => history.Go(request));
            }
            foreach (var coordinates in new[] { new[] { 0, 1 }, new[] { 3, 1 }, new[] { 1, 0 }, new[] { 1, 5 } })
            {
                var request = Location(); request.StartLine = coordinates[0]; request.StartColumn = coordinates[1];
                Assert.ThrowsException<ArgumentException>(() => history.Go(request));
            }
            history.Go(Location());
            history = new VbeNavigationHistory(null, r => Response.Failure("read failure"));
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => history.Go(Location())).Message, "read failure");
            history = new VbeNavigationHistory(null, r => r.Command == "read_module" ? Execute(r) : Response.Failure("selection failure"));
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => history.Go(Location())).Message, "selection failure");
            history = new VbeNavigationHistory(null, Execute);
            history.Go(Location("go", Path.Combine(Path.GetTempPath(), "never.xlsm")));
        }

        [TestMethod]
        public void LiveNavigationCapturesSelectionsBoundsStacksAndToleratesMissingActivePanes()
        {
            var host = new NavigationHost();
            var project = new VbeSessionTests.FakeProject { Name = "P", FileName = "" };
            host.VBProjects.Add(project);
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "Other", FileName = "" });
            host.Pane = new NavigationPane { CodeModule = new NavigationModule { Parent = new NavigationComponent { Collection = new NavigationCollection { Parent = project } } } };
            var history = new VbeNavigationHistory(host, Execute);
            history.Go(Location());
            Assert.AreEqual(1, Stack(history, "back").Count);
            Assert.ThrowsException<InvalidOperationException>(() => history.Go(Location("back", "Other")));
            history.Go(Location("back")); history.Go(Location("forward"));
            for (int i = 0; i < 102; i++) history.Go(Location());
            Assert.AreEqual(100, Stack(history, "back").Count);
            // Exercise defensive bounds when restoring a pre-existing history state.
            object entry = Stack(history, "back")[0];
            for (int i = 0; i < 100; i++) Stack(history, "forward").Add(entry);
            history.Go(Location("back")); Assert.AreEqual(100, Stack(history, "forward").Count);
            history.Go(Location("forward")); Assert.AreEqual(100, Stack(history, "back").Count);
            host.Pane = null; history.Go(Location()); history.Go(Location("back"));
            host.FailPane = true; history.Go(Location());
            host.FailPane = false;
            host.Pane = new NavigationPane { CodeModule = new NavigationModule { Parent = new NavigationComponent { Collection = new NavigationCollection { Parent = project } } } };
            history = new VbeNavigationHistory(host, r => r.Module == "Active" ? Response.Failure("closed module") : Execute(r));
            history.Go(Location()); Assert.AreEqual(0, Stack(history, "back").Count);
            project.ThrowFileName = true; history = new VbeNavigationHistory(host, Execute);
            history.Go(Location()); history.Bookmark(Location("add"));
            project.ThrowFileName = false; project.FileName = "relative.xlsm";
            history.Go(Location());
            project.FileName = Path.Combine(Path.GetTempPath(), "macro-navigation.xlsm");
            history.Go(Location()); history.Go(Location("back")); history.Go(Location("forward"));
        }
    }
}
