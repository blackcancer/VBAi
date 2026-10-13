using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading;
using VBAi.Tests.Infrastructure;
using C = VBAi.Tests.Infrastructure.ComInspectionFixture;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorReferenceOwnershipTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void LanguageReferencesReleaseAcquisitionsAndNeverReleaseTheBorrowedProject(bool duplicateIdentity)
        {
            var first = new C.Reference(); var second = duplicateIdentity ? first : new C.Reference { Broken = true, FailPath = true };
            var collection = new C.Collection(); collection.Values.Add(first); collection.Values.Add(second);
            var project = new C.Node(); project.Objects["References"] = collection;
            var released = new List<object>();
            string[] paths = ModernEditorWindow.ReadLanguageReferencePaths(project, CancellationToken.None,
                () => false, released.Add, value => value is C.Collection);
            CollectionAssert.AreEqual(duplicateIdentity ? new[] { first.Path, first.Path } : new[] { first.Path }, paths);
            first.Path = "Changed after capture";
            Assert.AreEqual("C:\\Library.tlb", paths[0]);
            CollectionAssert.AreEqual(new object[] { second, first, collection }, released);
            CollectionAssert.AreEqual(new[] { 1, 2 }, collection.ItemReads);
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorReads);
            Assert.IsFalse(released.Contains(project));
            Assert.AreEqual(duplicateIdentity ? 2 : 0, second.PathReads);
        }

        [DataTestMethod]
        [DataRow("collection")]
        [DataRow("count")]
        [DataRow("item")]
        [DataRow("path")]
        [DataRow("classifier")]
        public void FailedLanguageReadsBalanceEarlierGettersWithoutEnumerationFallback(string failure)
        {
            var first = new C.Reference { FailPath = failure == "path" }; var second = new C.Reference();
            var collection = new C.Collection { FailCount = failure == "count", FailingIndex = failure == "item" ? 2 : 0 };
            collection.Values.Add(first); collection.Values.Add(second);
            var project = new C.Node(); project.Objects["References"] = collection;
            if (failure == "collection") project.FailingGetters.Add("References");
            var released = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() => ModernEditorWindow.ReadLanguageReferencePaths(project,
                CancellationToken.None, () => false, released.Add,
                value => failure == "classifier" ? throw new InvalidOperationException("Classifier refused") : value is C.Collection));
            var expected = new List<object>();
            if (failure == "item" || failure == "path") expected.Add(first);
            if (failure != "collection") expected.Add(collection);
            CollectionAssert.AreEqual(expected, released);
            Assert.AreEqual(0, collection.EnumeratorReads);
            Assert.IsFalse(released.Contains(project));
        }

        [DataTestMethod]
        [DataRow("pre-cancel")]
        [DataRow("pre-busy")]
        [DataRow("count-cancel")]
        [DataRow("flag-cancel")]
        [DataRow("flag-busy")]
        [DataRow("path-cancel")]
        [DataRow("path-busy")]
        [DataRow("empty-count-cancel")]
        public void LanguageReadsRecheckCancellationAndNativeAvailabilityAfterPumpingGetters(string stage)
        {
            using (var cancellation = new CancellationTokenSource())
            {
                bool busy = stage == "pre-busy";
                var reference = new C.Reference(); var collection = new C.Collection();
                if (stage != "empty-count-cancel") collection.Values.Add(reference);
                var project = new C.Node(); project.Objects["References"] = collection;
                var released = new List<object>();
                Action cancelOrBusy = () => { if (stage.EndsWith("busy", StringComparison.Ordinal)) busy = true; else cancellation.Cancel(); };
                if (stage == "pre-cancel") cancellation.Cancel();
                if (stage == "count-cancel" || stage == "empty-count-cancel") collection.BeforeCount = cancelOrBusy;
                if (stage.StartsWith("flag", StringComparison.Ordinal)) reference.BeforeBroken = cancelOrBusy;
                if (stage.StartsWith("path", StringComparison.Ordinal)) reference.BeforePath = cancelOrBusy;
                Assert.ThrowsException<OperationCanceledException>(() => ModernEditorWindow.ReadLanguageReferencePaths(project,
                    cancellation.Token, () => busy, released.Add, value => value is C.Collection));
                if (stage.StartsWith("pre", StringComparison.Ordinal))
                {
                    Assert.AreEqual(0, project.ObjectReads.Count); Assert.AreEqual(0, released.Count);
                }
                else if (stage == "empty-count-cancel") CollectionAssert.AreEqual(new object[] { collection }, released);
                else CollectionAssert.AreEqual(new object[] { reference, collection }, released);
                Assert.AreEqual(stage.StartsWith("path", StringComparison.Ordinal) ? 1 : 0, reference.PathReads);
                Assert.IsFalse(released.Contains(project)); Assert.AreEqual(0, collection.EnumeratorReads);
                Assert.IsTrue(collection.ItemReads.Count <= 1, "No subsequent item may be acquired after cancellation.");
            }
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void LanguageReferencesRejectMissingArgumentsWithoutNativeAcquisition(bool missingProject)
        {
            var project = new C.Node(); var released = new List<object>();
            Assert.ThrowsException<ArgumentNullException>(() => ModernEditorWindow.ReadLanguageReferencePaths(
                missingProject ? null : project, CancellationToken.None, missingProject ? (Func<bool>)(() => false) : null, released.Add));
            Assert.AreEqual(0, released.Count); Assert.AreEqual(0, project.ObjectReads.Count);
        }

        [TestMethod]
        public void EmptyLanguageReferencesReturnAnEmptySnapshotAndReleaseTheirCollection()
        {
            var collection = new C.Collection(); var project = new C.Node(); project.Objects["References"] = collection;
            var released = new List<object>();
            Assert.AreEqual(0, ModernEditorWindow.ReadLanguageReferencePaths(project, CancellationToken.None,
                () => false, released.Add, value => value is C.Collection).Length);
            CollectionAssert.AreEqual(new object[] { collection }, released);
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorReads);
        }
    }
}
