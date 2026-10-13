using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;
using C = VBAi.Tests.Infrastructure.ComInspectionFixture;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeSessionTests
    {
        [DataTestMethod]
        [DataRow("empty")]
        [DataRow("valid")]
        [DataRow("broken")]
        public void ReferenceInventoryKeepsItsDynamicAndSerializedContractAfterRelease(string state)
        {
            var reference = new C.Reference
            {
                Broken = state == "broken", FailName = state == "broken", FailPath = state == "broken"
            };
            var collection = new C.Collection();
            if (state != "empty") collection.Values.Add(reference);
            var project = new C.Node(); project.Objects["References"] = collection;
            var released = new List<object>();
            dynamic snapshot = VbeSession.ReadReferenceInventory(() => project, "C:\\Disposable.xlsm",
                released.Add, value => value is C.Collection);

            // Access from the test assembly exercises the dynamic contract used by
            // existing callers, including references whose metadata type is private.
            Assert.AreEqual("C:\\Disposable.xlsm", (string)snapshot.Project);
            string version = (string)snapshot.Version;
            Assert.IsFalse(string.IsNullOrWhiteSpace(version));
            object[] rows = ((System.Collections.IEnumerable)snapshot.References).Cast<object>().ToArray();
            Assert.AreEqual(state == "empty" ? 0 : 1, rows.Length);
            Assert.AreEqual(state == "empty" ? 2 : 3, released.Count);
            Assert.AreEqual(0, collection.EnumeratorReads);
            Assert.AreEqual(state == "valid" ? 1 : 0, reference.NameReads);
            Assert.AreEqual(state == "valid" ? 1 : 0, reference.PathReads);
            Assert.AreEqual(state == "empty" ? 0 : 1, reference.GuidReads);

            string expectedName = state == "broken" ? null : reference.LibraryName;
            string expectedPath = state == "broken" ? null : reference.Path;
            reference.LibraryName = "Changed after release";
            reference.Path = "Changed after release";
            collection.Values.Clear(); project.Objects.Clear();
            var serializer = new JavaScriptSerializer();
            var payload = (IDictionary<string, object>)serializer.DeserializeObject(serializer.Serialize((object)snapshot));
            CollectionAssert.AreEquivalent(new[] { "Project", "Version", "References" }, payload.Keys.ToArray());
            Assert.AreEqual(version, payload["Version"]);
            var detached = (object[])payload["References"];
            Assert.AreEqual(rows.Length, detached.Length);
            if (state != "empty")
            {
                var item = (IDictionary<string, object>)detached.Single();
                Assert.AreEqual(expectedName, item["Name"]);
                Assert.AreEqual(expectedPath, item["FullPath"]);
                Assert.AreEqual(state == "broken", item["IsBroken"]);
                Assert.AreEqual(reference.GuidText, item["Guid"]);
                Assert.AreEqual(reference.Major, item["Major"]);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ReferenceInventoryReleasesEachGetterAndReturnsDetachedMetadata(bool duplicateIdentity)
        {
            var first = new C.Reference();
            var second = duplicateIdentity ? first : new C.Reference { LibraryName = "Other", Minor = 1 };
            var collection = new C.Collection(); collection.Values.Add(first); collection.Values.Add(second);
            var project = new C.Node(); project.Objects["References"] = collection;
            var released = new List<object>(); int resolves = 0;
            var result = VbeSession.ReadReferenceInventory(() => { resolves++; return project; },
                "C:\\Disposable.xlsm", released.Add, value => value is C.Collection);
            var serializer = new JavaScriptSerializer();
            string snapshot = serializer.Serialize(result);
            var metadata = (IDictionary<string, object>)serializer.DeserializeObject(snapshot);
            Assert.AreEqual("C:\\Disposable.xlsm", metadata["Project"]);
            Assert.AreEqual(64, ((string)metadata["Version"]).Length);
            Assert.AreEqual(2, ((object[])metadata["References"]).Length);
            first.LibraryName = "Changed after snapshot";
            Assert.AreEqual(snapshot, serializer.Serialize(result));
            Assert.AreEqual(1, resolves);
            CollectionAssert.AreEqual(new[] { "References" }, project.ObjectReads);
            CollectionAssert.AreEqual(new[] { 1, 2 }, collection.ItemReads);
            CollectionAssert.AreEqual(new object[] { second, first, collection, project }, released);
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorReads);
        }

        [DataTestMethod]
        [DataRow("project")]
        [DataRow("collection")]
        [DataRow("count")]
        [DataRow("negative-count")]
        [DataRow("second-item")]
        [DataRow("second-metadata")]
        public void ReferenceInventoryFailureBalancesOnlyActuallyAcquiredObjectsWithoutFallback(string failure)
        {
            var first = new C.Reference(); var second = new C.Reference { FailGuid = failure == "second-metadata" };
            var collection = new C.Collection { FailCount = failure == "count", FailingIndex = failure == "second-item" ? 2 : 0,
                ReportedCount = failure == "negative-count" ? (int?)-1 : null };
            collection.Values.Add(first); collection.Values.Add(second);
            var project = new C.Node(); project.Objects["References"] = collection;
            if (failure == "collection") project.FailingGetters.Add("References");
            var released = new List<object>(); int resolves = 0;
            Assert.ThrowsException<InvalidOperationException>(() => VbeSession.ReadReferenceInventory(() =>
            {
                resolves++; if (failure == "project") throw new InvalidOperationException("Project refused"); return project;
            }, "Project", released.Add, value => value is C.Collection));
            var expected = new List<object>();
            if (failure == "second-metadata") expected.Add(second);
            if (failure == "second-item" || failure == "second-metadata") expected.Add(first);
            if (failure != "project" && failure != "collection") expected.Add(collection);
            if (failure != "project") expected.Add(project);
            CollectionAssert.AreEqual(expected, released);
            Assert.AreEqual(1, resolves); Assert.AreEqual(0, collection.EnumeratorReads);
            Assert.AreEqual(failure == "second-item" || failure == "second-metadata" ? 2 : 0, collection.ItemReads.Count);
            Assert.AreEqual(failure == "second-metadata" ? 1 : 0, second.GuidReads);
        }

        [DataTestMethod]
        [DataRow("broken")]
        [DataRow("name")]
        [DataRow("path")]
        public void ReferenceInventoryPreservesIdentityWhenOptionalMetadataIsUnavailable(string condition)
        {
            var reference = new C.Reference { Broken = condition == "broken", FailName = condition != "path", FailPath = condition != "name" };
            var collection = new C.Collection(); collection.Values.Add(reference);
            var project = new C.Node(); project.Objects["References"] = collection;
            var released = new List<object>();
            object result = VbeSession.ReadReferenceInventory(() => project, "Project", released.Add, value => value is C.Collection);
            var data = (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(result));
            var item = (IDictionary<string, object>)((object[])data["References"])[0];
            Assert.AreEqual(reference.GuidText, item["Guid"]);
            Assert.AreEqual(condition == "broken", item["IsBroken"]);
            Assert.AreEqual(condition == "broken" || condition == "name" ? null : reference.LibraryName, item["Name"]);
            Assert.AreEqual(condition == "broken" || condition == "path" ? null : reference.Path, item["FullPath"]);
            Assert.AreEqual(condition == "broken" ? 0 : 1, reference.NameReads);
            Assert.AreEqual(condition == "broken" ? 0 : 1, reference.PathReads);
            CollectionAssert.AreEqual(new object[] { reference, collection, project }, released);
        }

        [TestMethod]
        public void ReferenceInventoryRejectsMissingResolverBeforeAcquiringAnything()
        {
            var released = new List<object>();
            Assert.ThrowsException<ArgumentNullException>(() => VbeSession.ReadReferenceInventory(null, "Project", released.Add));
            Assert.AreEqual(0, released.Count);
        }
    }
}
