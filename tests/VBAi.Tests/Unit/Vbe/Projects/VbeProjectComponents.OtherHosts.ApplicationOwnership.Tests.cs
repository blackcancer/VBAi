using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeOtherHostPersistenceTests
    {
        /// <summary>Managed NativeOM double whose application getter records one acquired reference.</summary>
        public sealed class NativeWindowApplication
        {
            public object Value;
            public bool Fail;
            public OwnedReadSequence Sequence;
            public object Application => Fail ? throw new InvalidOperationException("application getter refused") : Value;
            public IEnumerable Windows => Sequence;
        }

        [DataTestMethod]
        [DataRow("success"), DataRow("lookup failure"), DataRow("null native"), DataRow("getter failure")]
        [DataRow("foreign"), DataRow("owner failure"), DataRow("empty windows"), DataRow("alias")]
        public void NativeWindowApplicationTransfersOnlyTheVerifiedApplicationAcquisition(string fault)
        {
            var window = new OwnedReadWindow { Handle = 1 };
            var sequence = new OwnedReadSequence(fault == "empty windows" ? new object[0] : new object[] { window }, "none");
            var application = new OwnedReadApplication { Sequence = sequence };
            var nativeWindow = new NativeWindowApplication { Value = application, Sequence = sequence, Fail = fault == "getter failure" };
            if (fault == "alias") nativeWindow.Value = nativeWindow;
            var released = new List<object>();
            var probe = OwnedProbe("Word", released.Add);
            probe.ReadOwner = handle => fault == "owner failure" ? throw new InvalidOperationException("owner refused") :
                (uint)probe.CurrentProcessId + (fault == "foreign" ? 1u : 0u);
            object owned = fault == "null native" ? null : nativeWindow;
            object result = null;
            Action read = () => result = probe.ReadNativeWindowApplication(fault == "lookup failure" ? -1 : 0, owned);
            if (fault == "getter failure" || fault == "owner failure") Assert.ThrowsException<InvalidOperationException>(read);
            else read();
            if (fault == "null native") Assert.AreEqual(0, released.Count);
            else Assert.AreEqual(1, released.Count(value => ReferenceEquals(value, nativeWindow)),
                "The NativeOM acquisition is always balanced, including when the application aliases it.");
            bool success = fault == "success" || fault == "alias";
            if (success) Assert.AreSame(nativeWindow.Value, result);
            else Assert.IsNull(result);
            int expectedApplicationReleases = new[] { "foreign", "owner failure", "empty windows" }.Contains(fault) ? 1 : 0;
            Assert.AreEqual(expectedApplicationReleases, released.Count(value => ReferenceEquals(value, application)));
            Assert.AreEqual(sequence.Acquired.Count, released.Count(value => ReferenceEquals(value, window)));
            if (fault == "lookup failure") Assert.AreEqual(0, sequence.Acquired.Count);
        }

        [DataTestMethod]
        [DataRow("success"), DataRow("foreign"), DataRow("window failure")]
        public void RotLookupRetainsOnlyASuccessfullyTransferredApplication(string fault)
        {
            var application = new OwnedReadApplication { Window = new OwnedReadWindow { Handle = 1 },
                Fault = fault == "window failure" ? "getter" : "none" };
            var released = new List<object>(); var probe = OwnedProbe("Publisher", released.Add);
            int lookups = 0;
            probe.ReadActiveApplication = id => { lookups++; return application; };
            probe.ReadOwner = handle => (uint)probe.CurrentProcessId + (fault == "foreign" ? 1u : 0u);
            if (fault == "success")
            {
                Assert.AreSame(application, probe.Application());
                Assert.AreEqual(0, released.Count(value => ReferenceEquals(value, application)));
            }
            else
            {
                Assert.ThrowsException<InvalidOperationException>(() => probe.Application());
                Assert.AreEqual(1, released.Count(value => ReferenceEquals(value, application)));
            }
            Assert.AreEqual(1, lookups, "Ownership refusal must not activate another application.");
        }
    }
}
