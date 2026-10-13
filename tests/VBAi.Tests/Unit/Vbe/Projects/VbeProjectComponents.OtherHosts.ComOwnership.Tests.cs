using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeOtherHostPersistenceTests
    {
        [DataTestMethod]
        [DataRow("none"), DataRow("empty"), DataRow("collection"), DataRow("count")]
        [DataRow("item"), DataRow("changed"), DataRow("hwnd"), DataRow("owner")]
        [DataRow("zero"), DataRow("foreign"), DataRow("limit"), DataRow("negative"), DataRow("alias")]
        public void WordOwnershipReadBalancesEveryTraversalAcquisitionEvenOnRefusal(string fault)
        {
            var application = new OwnedReadApplication { Fault = fault };
            var first = new OwnedReadWindow { Handle = 1 };
            var second = new OwnedReadWindow { Handle = 2, Fail = fault == "hwnd" };
            object[] values = fault == "empty" ? new object[0] : fault == "limit" ? Enumerable.Repeat<object>(first, 1001).ToArray() :
                fault == "alias" ? new object[] { first, first } : new object[] { first, second };
            application.Sequence = new OwnedReadSequence(values, fault);
            var released = new List<object>();
            var native = OwnedProbe("Word", released.Add);
            native.ReadOwner = handle =>
            {
                if (fault == "owner") throw new InvalidOperationException("owner read");
                return handle.ToInt64() == 2 && fault == "zero" ? 0u : handle.ToInt64() == 2 && fault == "foreign" ? 88u : 77u;
            };
            Action read = () => Assert.AreEqual(fault == "empty" || fault == "zero" || fault == "foreign" ? 0u : 77u,
                native.ApplicationProcessId(application));
            if (new[] { "collection", "count", "item", "changed", "hwnd", "owner", "limit", "negative" }.Contains(fault))
                Assert.ThrowsException<InvalidOperationException>(read);
            else read();
            CollectionAssert.AreEqual(application.Acquired.Concat(application.Sequence.Acquired).Reverse().ToArray(), released);
            Assert.IsFalse(released.Contains(application));
        }

        [DataTestMethod, DataRow("none"), DataRow("getter"), DataRow("hwnd"), DataRow("owner"), DataRow("zero")]
        public void PublisherOwnershipReadReleasesItsWindowWithoutReleasingTheBorrowedApplication(string fault)
        {
            var application = new OwnedReadApplication { Fault = fault, Window = new OwnedReadWindow { Handle = 1, Fail = fault == "hwnd" } };
            var released = new List<object>(); var native = OwnedProbe("Publisher", released.Add);
            native.ReadOwner = handle => fault == "owner" ? throw new InvalidOperationException("owner read") : fault == "zero" ? 0u : 77u;
            Action read = () => Assert.AreEqual(fault == "zero" ? 0u : 77u, native.ApplicationProcessId(application));
            if (new[] { "getter", "hwnd", "owner" }.Contains(fault)) Assert.ThrowsException<InvalidOperationException>(read);
            else read();
            CollectionAssert.AreEqual(application.Acquired.AsEnumerable().Reverse().ToArray(), released);
            Assert.IsFalse(released.Contains(application));
        }

        [DataTestMethod]
        [DataRow("Word", "none"), DataRow("Word", "empty"), DataRow("Word", "alias")]
        [DataRow("Word", "collection"), DataRow("Word", "count"), DataRow("Word", "item"), DataRow("Word", "changed"), DataRow("Word", "limit")]
        [DataRow("Word", "negative")]
        [DataRow("PowerPoint", "none"), DataRow("PowerPoint", "empty"), DataRow("PowerPoint", "alias")]
        [DataRow("PowerPoint", "collection"), DataRow("PowerPoint", "count"), DataRow("PowerPoint", "item"), DataRow("PowerPoint", "changed"), DataRow("PowerPoint", "limit")]
        [DataRow("PowerPoint", "negative")]
        public void CatalogueTransfersItemsOnlyAfterCompleteTraversalAndReleasesUnreturnedItems(string kind, string fault)
        {
            object first = new object(), second = new object();
            object[] values = fault == "empty" ? new object[0] : fault == "limit" ? Enumerable.Repeat(first, 1001).ToArray() :
                fault == "alias" ? new[] { first, first } : new[] { first, second };
            var application = new OwnedReadApplication { Fault = fault, Sequence = new OwnedReadSequence(values, fault) };
            var released = new List<object>(); var native = OwnedProbe(kind, released.Add);
            IList<object> returned = null;
            Action read = () => returned = native.Documents(application);
            bool failed = new[] { "collection", "count", "item", "changed", "limit", "negative" }.Contains(fault);
            if (failed) Assert.ThrowsException<InvalidOperationException>(read);
            else { read(); CollectionAssert.AreEqual(values, returned.ToArray()); }
            var acquired = application.Acquired.Concat(application.Sequence.Acquired).ToArray();
            var traversal = acquired.Where(item => !values.Any(value => ReferenceEquals(value, item))).Reverse().ToArray();
            var expected = failed ? application.Sequence.Acquired.Where(item => values.Any(value => ReferenceEquals(value, item))).Reverse().Concat(traversal).ToArray() : traversal;
            CollectionAssert.AreEqual(expected, released);
            Assert.IsFalse(released.Contains(application));
            if (!failed) Assert.IsFalse(released.Any(item => values.Contains(item)), "Returned document acquisitions belong to the caller.");
        }

        [TestMethod]
        public void FailedTraversalReleaseDoesNotTransferDocumentReferencesOrRepeatTheFailedRelease()
        {
            var application = new OwnedReadApplication { Sequence = new OwnedReadSequence(new[] { new object(), new object() }, "none") };
            var released = new List<object>();
            var native = OwnedProbe("Word", value =>
            {
                released.Add(value);
                if (ReferenceEquals(value, application.Sequence)) throw new InvalidOperationException("collection release");
            });
            Assert.ThrowsException<InvalidOperationException>(() => native.Documents(application));
            Assert.AreEqual(1, released.Count(item => ReferenceEquals(item, application.Sequence)));
            Assert.AreEqual(2, released.Count(item => application.Sequence.Values.Contains(item)));
            Assert.IsFalse(released.Contains(application));
            Assert.AreEqual(1, released.Count(item => ReferenceEquals(item, application.Sequence)), "The failed release is not retried.");
        }

        private static VbeProjectComponents.NativeOtherHostProbe OwnedProbe(string host, Action<object> release) =>
            new VbeProjectComponents.NativeOtherHostProbe { ReadHostKind = () => host, ReleaseAcquiredReference = release };

        public sealed class OwnedReadApplication
        {
            public string Fault;
            public OwnedReadSequence Sequence;
            public OwnedReadWindow Window;
            public readonly List<object> Acquired = new List<object>();
            public IEnumerable Windows => ReadCollection();
            public IEnumerable Documents => ReadCollection();
            public IEnumerable Presentations => ReadCollection();
            private IEnumerable ReadCollection()
            {
                if (Fault == "collection") throw new InvalidOperationException("collection getter");
                Acquired.Add(Sequence); return Sequence;
            }
            public object ActiveWindow
            {
                get { if (Fault == "getter") throw new InvalidOperationException("window getter"); Acquired.Add(Window); return Window; }
            }
        }

        public sealed class OwnedReadWindow
        {
            public long Handle;
            public bool Fail;
            public long Hwnd => Fail ? throw new InvalidOperationException("window handle read") : Handle;
            public long hWnd => Hwnd;
        }

        public sealed class OwnedReadSequence : IEnumerable
        {
            public readonly object[] Values;
            public readonly List<object> Acquired = new List<object>();
            public readonly string Fault;
            private int countReads;
            public OwnedReadSequence(object[] values, string fault) { Values = values; Fault = fault; }
            public IEnumerator GetEnumerator()
            {
                throw new AssertFailedException("Office reads must use explicit Count/Item instead of an implicit COM enumerator.");
            }
            public int Count
            {
                get
                {
                    if (Fault == "count") throw new InvalidOperationException("collection count");
                    if (Fault == "negative") return -1;
                    return Values.Length + (Fault == "changed" && ++countReads > 1 ? 1 : 0);
                }
            }
            public object Item(int index)
            {
                if (Fault == "item" && index == 2) throw new InvalidOperationException("collection item");
                object value = Values[index - 1]; Acquired.Add(value); return value;
            }
        }
    }
}
