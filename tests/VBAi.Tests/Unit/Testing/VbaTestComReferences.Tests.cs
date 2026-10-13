using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbaTestComReferencesTests
    {
        [DataTestMethod]
        [DataRow(false, false)]
        [DataRow(false, true)]
        [DataRow(true, false)]
        [DataRow(true, true)]
        public void IndexedCollectionsBalanceOnlyAcquiredItemsWithoutCreatingAnEnumerator(bool borrowed, bool early)
        {
            var collection = new IndexedSequence(); var released = new List<object>();
            using (var scope = new VbaTestComReferences(released.Add, _ => true))
                foreach (object item in borrowed ? scope.ItemsFromBorrowed(collection) : scope.Items(collection))
                    if (early) break;
            var expected = new List<object>();
            if (!early) expected.Add(collection.Second);
            expected.Add(collection.First);
            if (!borrowed) expected.Add(collection);
            CollectionAssert.AreEqual(expected, released);
            CollectionAssert.AreEqual(early ? new[] { 1 } : new[] { 1, 2 }, collection.Reads);
            Assert.AreEqual(1, collection.CountReads);
            Assert.AreEqual(0, collection.EnumeratorRequests);
        }

        [DataTestMethod]
        [DataRow(false, 1)]
        [DataRow(false, 2)]
        [DataRow(true, 1)]
        [DataRow(true, 2)]
        public void IndexedGetterFailureReleasesEarlierAcquisitionsWithoutFallingBackOrRetrying(bool borrowed, int failure)
        {
            var collection = new IndexedSequence { FailingIndex = failure }; var released = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() =>
            {
                using (var scope = new VbaTestComReferences(released.Add, _ => true))
                    foreach (object item in borrowed ? scope.ItemsFromBorrowed(collection) : scope.Items(collection)) { }
            });
            var expected = new List<object>();
            if (failure == 2) expected.Add(collection.First);
            if (!borrowed) expected.Add(collection);
            CollectionAssert.AreEqual(expected, released);
            Assert.AreEqual(failure, collection.Reads.Count);
            Assert.AreEqual(0, collection.EnumeratorRequests);
        }

        [DataTestMethod]
        [DataRow("empty")]
        [DataRow("negative")]
        [DataRow("throw")]
        public void IndexedCountReadsAreBoundedAndNeverFallBackToEnumeration(string outcome)
        {
            var collection = new IndexedSequence { ReportedCount = outcome == "empty" ? 0 : -1, ThrowCount = outcome == "throw" };
            var released = new List<object>();
            Action read = () =>
            {
                using (var scope = new VbaTestComReferences(released.Add, _ => true))
                    foreach (object item in scope.Items(collection)) { }
            };
            if (outcome == "empty") read(); else Assert.ThrowsException<InvalidOperationException>(read);
            CollectionAssert.AreEqual(new object[] { collection }, released);
            Assert.AreEqual(1, collection.CountReads); Assert.AreEqual(0, collection.Reads.Count);
            Assert.AreEqual(0, collection.EnumeratorRequests);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void DeferredBorrowedEnumerationRefusesDisposedOrRetainedScopesBeforeAnyGetter(bool retain)
        {
            var collection = new IndexedSequence(); var scope = new VbaTestComReferences(null, _ => true);
            var sequence = scope.ItemsFromBorrowed(collection);
            if (retain) scope.Retain(); else scope.Dispose();
            Assert.ThrowsException<ObjectDisposedException>(() => { foreach (object item in sequence) { } });
            Assert.AreEqual(0, collection.CountReads); Assert.AreEqual(0, collection.Reads.Count);
            Assert.AreEqual(0, collection.EnumeratorRequests);
        }

        [TestMethod]
        public void RepeatedIndexedIdentityBalancesEachGetterAndTransfersOnlyOneAcquisition()
        {
            var collection = new IndexedSequence(); collection.Second = collection.First;
            var released = new List<object>();
            using (var scope = new VbaTestComReferences(released.Add, _ => true))
            {
                foreach (object item in scope.ItemsFromBorrowed(collection)) { }
                scope.Transfer(collection.First);
            }
            CollectionAssert.AreEqual(new[] { collection.First }, released);
            using (var owner = new VbaTestComReferences(released.Add)) owner.Add(collection.First);
            CollectionAssert.AreEqual(new[] { collection.First, collection.First }, released);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ManagedDisposableEnumeratorsAreDisposedOnceAfterFullOrEarlyEnumeration(bool early)
        {
            var collection = new DisposableSequence(); var scope = new VbaTestComReferences();
            foreach (object item in scope.ItemsFromBorrowed(collection)) if (early) break;
            scope.Dispose(); scope.Dispose();
            Assert.AreEqual(1, collection.Iterator.Disposals);
        }

        [TestMethod]
        public void FailedManagedIteratorDisposalIsRetainedAndNeverRetried()
        {
            var collection = new DisposableSequence(); collection.Iterator.ThrowDispose = true;
            var scope = new VbaTestComReferences();
            foreach (object item in scope.ItemsFromBorrowed(collection)) { }
            Assert.ThrowsException<InvalidOperationException>(() => scope.Dispose());
            scope.Dispose();
            Assert.AreEqual(1, collection.Iterator.Disposals);
            Assert.ThrowsException<ObjectDisposedException>(() => scope.Add(new object()));
        }

        [TestMethod]
        public void IndexedClassificationFailureReleasesTheCollectionWithoutTouchingItsGetters()
        {
            var collection = new IndexedSequence(); var released = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() =>
            {
                using (var scope = new VbaTestComReferences(released.Add, _ => throw new InvalidOperationException("Classification refused")))
                    foreach (object item in scope.Items(collection)) { }
            });
            CollectionAssert.AreEqual(new object[] { collection }, released);
            Assert.AreEqual(0, collection.CountReads); Assert.AreEqual(0, collection.EnumeratorRequests);
            Assert.ThrowsException<ArgumentNullException>(() => new VbaTestComReferences(null, null));
        }

        [TestMethod]
        public void IndexedReleaseFailureRetainsTheRemainingAcquisitionsWithoutRetry()
        {
            var collection = new IndexedSequence(); var released = new List<object>();
            var scope = new VbaTestComReferences(value =>
            {
                released.Add(value);
                if (ReferenceEquals(value, collection.Second)) throw new InvalidOperationException("Release refused");
            }, _ => true);
            foreach (object item in scope.Items(collection)) { }
            Assert.ThrowsException<InvalidOperationException>(() => scope.Dispose());
            scope.Dispose();
            CollectionAssert.AreEqual(new[] { collection.Second }, released);
        }

        [TestMethod]
        public void UncertainIndexedOutcomeKeepsTheCollectionAndEveryItemAcquisition()
        {
            var collection = new IndexedSequence(); var released = new List<object>();
            var scope = new VbaTestComReferences(released.Add, _ => true);
            foreach (object item in scope.Items(collection)) { }
            scope.Retain(); scope.Dispose();
            Assert.AreEqual(0, released.Count);
            Assert.AreEqual(0, collection.EnumeratorRequests);
        }

        public sealed class IndexedSequence : IEnumerable
        {
            public object First = new object(), Second = new object();
            public int CountReads, EnumeratorRequests, FailingIndex;
            public int ReportedCount = 2;
            public bool ThrowCount;
            public readonly List<int> Reads = new List<int>();
            public int Count { get { CountReads++; if (ThrowCount) throw new InvalidOperationException("Count refused"); return ReportedCount; } }
            public object this[int index]
            {
                get
                {
                    Reads.Add(index);
                    if (index == FailingIndex) throw new InvalidOperationException("Item refused");
                    if (index != 1 && index != 2) throw new InvalidOperationException("Wrong native index");
                    return index == 1 ? First : Second;
                }
            }
            public IEnumerator GetEnumerator() { EnumeratorRequests++; throw new InvalidOperationException("Native enumeration must not be requested"); }
        }

        private sealed class DisposableSequence : IEnumerable
        {
            internal readonly DisposableIterator Iterator = new DisposableIterator();
            public IEnumerator GetEnumerator() => Iterator;
        }

        private sealed class DisposableIterator : IEnumerator, IDisposable
        {
            private int index;
            internal int Disposals;
            internal bool ThrowDispose;
            public object Current => new object();
            public bool MoveNext() => ++index <= 2;
            public void Reset() => index = 0;
            public void Dispose() { Disposals++; if (ThrowDispose) throw new InvalidOperationException("Iterator release refused"); }
        }

        [TestMethod]
        public void GetterAcquisitionsAreBalancedInReverseOrderWithoutReleasingBorrowedArguments()
        {
            var released = new List<object>();
            object borrowed = new object(), repeated = new object(), child = new object();
            var scope = new VbaTestComReferences(released.Add);
            scope.Add(repeated); scope.Add(repeated); scope.Add(null); scope.Add(child);
            scope.Dispose(); scope.Dispose();
            CollectionAssert.AreEqual(new[] { child, repeated, repeated }, released);
            Assert.IsFalse(released.Contains(borrowed));
            Assert.ThrowsException<ObjectDisposedException>(() => scope.Add(new object()));
        }

        [TestMethod]
        public void TransferRemovesOnlyOneAcquisitionAndItsNewOwnerReleasesIt()
        {
            var released = new List<object>(); object value = new object();
            var scope = new VbaTestComReferences(released.Add);
            scope.Add(value); scope.Add(value); scope.Transfer(value);
            scope.Dispose();
            Assert.AreEqual(1, released.Count);
            var owner = new VbaTestComReferences(released.Add); owner.Add(value); owner.Dispose();
            Assert.AreEqual(2, released.Count);
            Assert.ThrowsException<InvalidOperationException>(() => new VbaTestComReferences().Transfer(value));
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void CollectionAndEnumeratorAreBalancedAfterFullOrEarlyEnumeration(bool stopEarly)
        {
            var collection = new Sequence(); var released = new List<object>();
            using (var scope = new VbaTestComReferences(released.Add))
                foreach (object item in scope.Items(collection)) if (stopEarly) break;
            Assert.AreEqual(stopEarly ? 3 : 4, released.Count);
            Assert.AreSame(collection, released[released.Count - 1]);
            Assert.AreSame(collection.Enumerator, released[released.Count - 2]);
            Assert.AreSame(collection.First, released[stopEarly ? 0 : 1]);
        }

        [TestMethod]
        public void EnumerationFailureReleasesAlreadyAcquiredItemsAndBorrowedCollectionRemainsOwnedByCaller()
        {
            var collection = new Sequence { ThrowAfterFirst = true }; var released = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() =>
            {
                using (var scope = new VbaTestComReferences(released.Add))
                    foreach (object item in scope.ItemsFromBorrowed(collection)) { }
            });
            CollectionAssert.AreEqual(new[] { collection.First, collection.Enumerator }, released);
            Assert.IsFalse(released.Contains(collection));
        }

        [TestMethod]
        public void UncertainNativeOutcomeRetainsTheAcquisitionsAndCannotReleaseOnLaterDisposal()
        {
            var released = new List<object>(); var scope = new VbaTestComReferences(released.Add);
            scope.Add(new object()); scope.Retain(); scope.Retain(); scope.Dispose();
            Assert.AreEqual(0, released.Count);
            Assert.ThrowsException<ObjectDisposedException>(() => scope.Add(new object()));
        }

        [TestMethod]
        public void ReleaseFailureRetainsRemainingObjectsAndDoesNotRetryTheFailedRelease()
        {
            var released = new List<object>(); object first = new object(), failed = new object(), last = new object();
            var scope = new VbaTestComReferences(value =>
            {
                released.Add(value);
                if (ReferenceEquals(value, failed)) throw new InvalidOperationException("Release failed");
            });
            scope.Add(first); scope.Add(failed); scope.Add(last);
            Assert.ThrowsException<InvalidOperationException>(() => scope.Dispose());
            scope.Dispose(); scope.Retain();
            CollectionAssert.AreEqual(new[] { last, failed }, released);
        }

        [TestMethod]
        public void DefaultReleaseAcceptsManagedDoublesWithoutFinalizingBorrowedObjects()
        {
            VbaTestComReferences.Release(null); VbaTestComReferences.Release(new object());
            using (var scope = new VbaTestComReferences()) scope.Add(new object());
        }

        private sealed class Sequence : IEnumerable
        {
            internal readonly object First = new object(), Second = new object();
            internal bool ThrowAfterFirst;
            internal Iterator Enumerator;
            public IEnumerator GetEnumerator() { return Enumerator = new Iterator(this); }
            internal sealed class Iterator : IEnumerator
            {
                private readonly Sequence sequence; private int index;
                internal Iterator(Sequence sequence) { this.sequence = sequence; }
                public object Current => index == 1 ? sequence.First : sequence.Second;
                public bool MoveNext()
                {
                    if (sequence.ThrowAfterFirst && index == 1) throw new InvalidOperationException("Enumerator failed");
                    return ++index <= 2;
                }
                public void Reset() { index = 0; }
            }
        }
    }
}
