using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using C = VBAi.Tests.Infrastructure.ComInspectionFixture;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeEditorWindowsTests
    {
        private static IDictionary<string, object> InspectionData(object value)
            => (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(value));

        private static VbeEditorWindows OwnedReader(C.Node editor, List<object> released)
            => new VbeEditorWindows(editor) { ReleaseAcquired = released.Add, IndexedCollection = value => value is C.Collection };

        private static object Inspect(VbeEditorWindows reader, string kind)
            => kind == "CodePanes" ? reader.CodePanes() : kind == "AddIns" ? reader.AddIns() : reader.Windows();

        [DataTestMethod]
        [DataRow("Windows")]
        [DataRow("CodePanes")]
        [DataRow("AddIns")]
        [DataRow("EmptyWindows")]
        public void InspectionsBalanceEachCollectionItemAndActiveAliasWithoutReleasingTheEditor(string kind)
        {
            var item = new C.Node(); var collection = new C.Collection();
            if (kind != "EmptyWindows") { collection.Values.Add(item); collection.Values.Add(item); }
            var editor = new C.Node(); string property = kind == "EmptyWindows" ? "Windows" : kind;
            editor.Objects[property] = collection;
            string active = kind == "CodePanes" ? "ActiveCodePane" : "ActiveWindow";
            editor.Objects[active] = kind == "EmptyWindows" ? null : item;
            var released = new List<object>();
            var snapshot = InspectionData(Inspect(OwnedReader(editor, released), kind));
            Assert.AreEqual(kind == "EmptyWindows" ? 0 : 2, ((object[])snapshot[property]).Length);
            int acquisitions = kind == "EmptyWindows" ? 0 : kind == "AddIns" ? 2 : 3;
            CollectionAssert.AreEqual(Enumerable.Repeat<object>(item, acquisitions).Concat(new object[] { collection }).ToArray(), released);
            Assert.IsFalse(released.Contains(editor));
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorReads);
            CollectionAssert.AreEqual(kind == "EmptyWindows" ? new int[0] : new[] { 1, 2 }, collection.ItemReads);
        }

        [DataTestMethod]
        [DataRow("Windows", false)]
        [DataRow("Windows", true)]
        [DataRow("CodePanes", false)]
        [DataRow("CodePanes", true)]
        [DataRow("AddIns", false)]
        [DataRow("AddIns", true)]
        public void FailedInspectionCountOrItemDoesNotRetryAndBalancesEarlierAcquisitions(string kind, bool failItem)
        {
            var item = new C.Node(); var collection = new C.Collection { FailCount = !failItem, FailingIndex = failItem ? 2 : 0 };
            collection.Values.Add(item); collection.Values.Add(new C.Node());
            var editor = new C.Node(); editor.Objects[kind] = collection;
            var released = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() => Inspect(OwnedReader(editor, released), kind));
            CollectionAssert.AreEqual(failItem ? new object[] { item, collection } : new object[] { collection }, released);
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorReads);
            Assert.IsFalse(editor.ObjectReads.Contains("ActiveWindow"));
            Assert.IsFalse(editor.ObjectReads.Contains("ActiveCodePane"));
            Assert.IsFalse(released.Contains(editor));
        }

        [TestMethod]
        public void FailedActiveWindowGetterRetainsItsDiagnosticAndReleasesAllEnumeratedWindows()
        {
            var item = new C.Node(); var collection = new C.Collection(); collection.Values.Add(item);
            var editor = new C.Node(); editor.Objects["Windows"] = collection; editor.FailingGetters.Add("ActiveWindow");
            var released = new List<object>();
            var snapshot = InspectionData(OwnedReader(editor, released).Windows());
            StringAssert.Contains((string)((IDictionary<string, object>)snapshot["ActiveWindow"])["Error"], "ActiveWindow refused");
            CollectionAssert.AreEqual(new object[] { item, collection }, released);
            Assert.AreEqual(0, collection.EnumeratorReads);
        }

        [TestMethod]
        public void PaneInspectionBalancesEveryCodeModuleComponentCollectionProjectAndWindowGetter()
        {
            var editor = new C.Node(); var pane = new C.Node(); var module = new C.Node(); var component = new C.Node();
            var project = new C.Node(); var window = new C.Node(); var components = new C.Collection { Parent = project };
            var panes = new C.Collection(); panes.Values.Add(pane);
            pane.Objects["CodeModule"] = module; pane.Objects["Window"] = window;
            module.Objects["Parent"] = component; component.Objects["Collection"] = components;
            editor.Objects["CodePanes"] = panes; editor.Objects["ActiveCodePane"] = pane;
            var released = new List<object>();
            var data = InspectionData(OwnedReader(editor, released).CodePanes());
            var active = (IDictionary<string, object>)data["ActiveCodePane"];
            Assert.AreEqual(0, ((IDictionary<string, object>)active["Errors"]).Count);
            var fields = (IDictionary<string, object>)active["Properties"];
            Assert.AreEqual(project.FileName, fields["ProjectPath"]);
            Assert.AreEqual(4, ((IDictionary<string, object>)fields["Selection"])["EndColumn"]);
            foreach (var expected in new[] { Tuple.Create((object)pane, 2), Tuple.Create((object)module, 6),
                Tuple.Create((object)component, 6), Tuple.Create((object)components, 4), Tuple.Create((object)project, 4),
                Tuple.Create((object)window, 2), Tuple.Create((object)panes, 1) })
                Assert.AreEqual(expected.Item2, released.Count(item => ReferenceEquals(item, expected.Item1)));
            Assert.AreEqual(25, released.Count); Assert.IsFalse(released.Contains(editor));
            Assert.AreEqual(0, panes.EnumeratorReads); Assert.AreEqual(0, components.EnumeratorReads);
            Assert.IsFalse(component.ObjectReads.Contains("CodePane"), "Inspecting an existing pane must never open a new pane.");
        }

        [DataTestMethod]
        [DataRow("CodeModule")]
        [DataRow("Parent")]
        [DataRow("Collection")]
        public void FailedPaneMetadataStillBalancesEveryPrecedingGetterAndReportsPropertyErrors(string getter)
        {
            var editor = new C.Node(); var pane = new C.Node(); var module = new C.Node(); var component = new C.Node();
            var project = new C.Node(); var window = new C.Node(); var components = new C.Collection { Parent = project };
            var panes = new C.Collection(); panes.Values.Add(pane);
            pane.Objects["CodeModule"] = module; pane.Objects["Window"] = window;
            module.Objects["Parent"] = component; component.Objects["Collection"] = components;
            if (getter == "CodeModule") pane.FailingGetters.Add(getter);
            else if (getter == "Parent") module.FailingGetters.Add(getter);
            else component.FailingGetters.Add(getter);
            editor.Objects["CodePanes"] = panes; editor.Objects["ActiveCodePane"] = pane;
            var released = new List<object>();
            var data = InspectionData(OwnedReader(editor, released).CodePanes());
            var active = (IDictionary<string, object>)data["ActiveCodePane"];
            var errors = (IDictionary<string, object>)active["Errors"];
            Assert.IsTrue(errors.ContainsKey("Project")); Assert.IsTrue(errors.ContainsKey("ProjectPath"));
            Assert.AreEqual(getter == "CodeModule" ? 5 : getter == "Parent" ? 11 : 17, released.Count);
            Assert.IsFalse(released.Contains(project)); Assert.IsFalse(released.Contains(components));
            Assert.AreEqual(0, panes.EnumeratorReads); Assert.IsFalse(released.Contains(editor));
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void EnvironmentCountsBalanceAliasedGetterResultsWithoutEnumeratingOrReleasingTheEditor(bool failedWindowGetter)
        {
            var editor = new C.Node(); var active = new C.Node();
            var projects = new C.Collection(); var shared = new C.Collection(); var addIns = new C.Collection();
            editor.Objects["VBProjects"] = projects; editor.Objects["Windows"] = shared;
            editor.Objects["CodePanes"] = shared; editor.Objects["AddIns"] = addIns; editor.Objects["ActiveVBProject"] = active;
            if (failedWindowGetter) editor.FailingGetters.Add("Windows");
            var released = new List<object>(); var data = InspectionData(OwnedReader(editor, released).Environment());
            Assert.AreEqual(active.Name, ((IDictionary<string, object>)data["Properties"])["ActiveProject"]);
            Assert.AreEqual(failedWindowGetter ? 1 : 0, ((IDictionary<string, object>)data["Errors"]).Count);
            CollectionAssert.AreEqual(failedWindowGetter ? new object[] { active, addIns, shared, projects }
                : new object[] { active, addIns, shared, shared, projects }, released);
            Assert.AreEqual(0, projects.EnumeratorReads + shared.EnumeratorReads + addIns.EnumeratorReads);
            Assert.AreEqual(0, projects.ItemReads.Count + shared.ItemReads.Count + addIns.ItemReads.Count);
            Assert.IsFalse(released.Contains(editor));
        }
    }
}
