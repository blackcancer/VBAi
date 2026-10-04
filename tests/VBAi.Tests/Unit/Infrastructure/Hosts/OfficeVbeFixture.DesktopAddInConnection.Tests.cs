using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureDesktopAddInConnectionTests
    {
        private const string Registration = "CurrentUser|DWord|3;LocalMachine|KEY_ABSENT";

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ExactConnectionBooleanIsObservedOnceAndOnlyFalsePermitsOneSetter(bool initial)
        {
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            var rows = new List<IDictionary<string, object>>();
            int reads = 0, setters = 0, guards = 0;
            bool connected = initial;
            gate.Run("VBAi.AddIn", () => guards++, () => { reads++; return connected; }, () => {
                Assert.IsTrue(rows.Any(row => State(row) == "SETTER_CLAIMED"), "A durable mutation claim must precede the sole setter.");
                setters++; connected = true;
            }, () => Registration, rows.Add);
            Assert.IsTrue(gate.Consumed); Assert.IsFalse(gate.DeliveryUncertain);
            Assert.AreEqual(initial ? 1 : 2, reads); Assert.AreEqual(initial ? 0 : 1, setters);
            Assert.AreEqual(setters, gate.SetterEntries);
            Assert.AreEqual(initial ? 2 : 5, guards);
            Assert.AreEqual(initial ? "NOOP_ALREADY_CONNECTED" : "CONNECTED_VERIFIED", State(rows.Last()));
            Assert.IsTrue(rows.Where(row => row.ContainsKey("RegistryWrites")).All(row => Equals(row["RegistryWrites"], 0)));
            Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => guards++, () => { reads++; return false; },
                () => setters++, () => Registration, rows.Add));
            Assert.AreEqual(initial ? 1 : 2, reads); Assert.AreEqual(initial ? 0 : 1, setters);
        }

        [DataTestMethod, DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5)]
        public void ChangedOriginalOwnershipAtAnyConnectionBoundaryRefusesFurtherDispatch(int failedGuard)
        {
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            int guards = 0, setters = 0, reads = 0;
            var original = new InvalidOperationException("synthetic changed original owner/window/desktop");
            var error = Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => {
                if (++guards == failedGuard) throw original;
            }, () => { reads++; return setters != 0; }, () => setters++, () => Registration, row => { }));
            Assert.AreSame(original, error);
            Assert.AreEqual(failedGuard <= 3 ? 0 : 1, setters);
            Assert.IsTrue(gate.Consumed);
            Assert.AreEqual(setters != 0, gate.DeliveryUncertain);
            int priorReads = reads;
            Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => { }, () => { reads++; return true; },
                () => setters++, () => Registration, row => { }));
            Assert.AreEqual(priorReads, reads);
        }

        [TestMethod]
        public void WrongProgIdUnknownBooleanAndFailedClaimCannotEnterASetter()
        {
            foreach (string progId in new[] { "Other.AddIn", "vbai.addin", "", null })
            {
                int reads = 0, setters = 0;
                var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
                Assert.ThrowsException<InvalidOperationException>(() => gate.Run(progId, () => { }, () => { reads++; return false; },
                    () => setters++, () => Registration, row => { }));
                Assert.AreEqual(0, reads); Assert.AreEqual(0, setters); Assert.IsTrue(gate.Consumed);
            }
            foreach (object unknown in new object[] { null, 0, "False" })
            {
                int setters = 0;
                var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
                Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => { }, () => unknown,
                    () => setters++, () => Registration, row => { }));
                Assert.AreEqual(0, setters);
            }
            var failedClaim = new InvalidOperationException("synthetic durable claim failure");
            var claimGate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            int claimSetters = 0;
            Assert.AreSame(failedClaim, Assert.ThrowsException<InvalidOperationException>(() => claimGate.Run("VBAi.AddIn", () => { },
                () => false, () => claimSetters++, () => Registration, row => { if (State(row) == "SETTER_CLAIMED") throw failedClaim; })));
            Assert.AreEqual(0, claimSetters); Assert.AreEqual(0, claimGate.SetterEntries); Assert.IsTrue(claimGate.Consumed);
        }

        [TestMethod]
        public void SetterFailureRetainsOriginalErrorAndUncertainConnectionWithoutReadbackOrRetry()
        {
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            var original = new COMException("synthetic connection failure", unchecked((int)0x80010001));
            int reads = 0, setters = 0;
            var error = Assert.ThrowsException<COMException>(() => gate.Run("VBAi.AddIn", () => { }, () => { reads++; return false; },
                () => { setters++; throw original; }, () => Registration, row => { }));
            Assert.AreSame(original, error); Assert.AreEqual(1, setters); Assert.AreEqual(1, reads);
            Assert.IsTrue(gate.DeliveryUncertain); Assert.IsTrue(gate.Consumed);
        }

        [TestMethod]
        public void GetterFailureAndFailedPostSetterReadbackAreNeverRetried()
        {
            var original = new COMException("synthetic getter failure", unchecked((int)0x80010001));
            int reads = 0, setters = 0;
            var getterGate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => getterGate.Run("VBAi.AddIn", () => { },
                () => { reads++; throw original; }, () => setters++, () => Registration, row => { })));
            Assert.AreEqual(1, reads); Assert.AreEqual(0, setters);
            reads = 0;
            var readbackGate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => readbackGate.Run("VBAi.AddIn", () => { }, () => {
                if (++reads == 1) return false; throw original;
            }, () => setters++, () => Registration, row => { })));
            Assert.AreEqual(2, reads); Assert.AreEqual(1, setters); Assert.IsTrue(readbackGate.DeliveryUncertain);
            foreach (object after in new object[] { false, null, 1, "True" })
            {
                reads = 0; setters = 0;
                var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
                Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => { },
                    () => ++reads == 1 ? (object)false : after, () => setters++, () => Registration, row => { }));
                Assert.AreEqual(2, reads); Assert.AreEqual(1, setters); Assert.IsTrue(gate.DeliveryUncertain);
            }
        }

        [DataTestMethod, DataRow(2), DataRow(3)]
        public void ChangedLoadBehaviorRefusesWithoutRegistryRepairOrSecondConnection(int changedRead)
        {
            int registrationReads = 0, setters = 0, reads = 0;
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            Assert.ThrowsException<InvalidOperationException>(() => gate.Run("VBAi.AddIn", () => { }, () => ++reads != 1,
                () => setters++, () => ++registrationReads == changedRead ? "CurrentUser|DWord|2;LocalMachine|KEY_ABSENT" : Registration, row => { }));
            Assert.AreEqual(changedRead == 2 ? 0 : 1, setters);
            Assert.AreEqual(setters != 0, gate.DeliveryUncertain); Assert.IsTrue(gate.Consumed);
        }

        [TestMethod]
        public void FailedEvidencePreservesOriginalErrorOrRefusesVerifiedConnectionDelivery()
        {
            var original = new COMException("synthetic setter failure");
            var evidence = new InvalidOperationException("synthetic failed receipt");
            var gate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            var aggregate = Assert.ThrowsException<AggregateException>(() => gate.Run("VBAi.AddIn", () => { }, () => false,
                () => { throw original; }, () => Registration, row => { if (State(row) == "FAILED_NO_RETRY") throw evidence; }));
            Assert.AreSame(original, aggregate.InnerExceptions[0]); Assert.AreSame(evidence, aggregate.InnerExceptions[1]);
            int reads = 0, setters = 0;
            var verifiedGate = new OfficeVbeFixture.PrivateAccessAddInConnectionGate();
            Assert.AreSame(evidence, Assert.ThrowsException<InvalidOperationException>(() => verifiedGate.Run("VBAi.AddIn", () => { },
                () => ++reads != 1, () => setters++, () => Registration, row => { if (State(row) == "CONNECTED_VERIFIED") throw evidence; })));
            Assert.AreEqual(1, setters); Assert.IsTrue(verifiedGate.DeliveryUncertain); Assert.IsTrue(verifiedGate.Consumed);
        }

        [TestMethod]
        public void FixtureStaAndExactVbeWindowThreadOwnershipAreRequired()
        {
            OfficeVbeFixture.RequirePrivateAccessAddInEntry("VBAi.AddIn", typeof(VBAi.AddIn).GUID.ToString("B"));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInEntry("Other.AddIn", typeof(VBAi.AddIn).GUID.ToString("D")));
            foreach (string identity in new[] { "malformed", Guid.Empty.ToString("D"), null })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInEntry("VBAi.AddIn", identity));
            OfficeVbeFixture.RequirePrivateAccessAddInFixtureThread(10, 10, ApartmentState.STA);
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInFixtureThread(10, 11, ApartmentState.STA));
            foreach (var apartment in new[] { ApartmentState.MTA, ApartmentState.Unknown })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInFixtureThread(10, 10, apartment));
            var window = new IntPtr(500);
            OfficeVbeFixture.RequirePrivateAccessAddInWindow(100, window, 20, window, 100, 20);
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInWindow(100, window, 20, window, 100, 20, false));
            foreach (var changed in new[] { new Tuple<IntPtr, uint, uint>(IntPtr.Zero, 100, 20), Tuple.Create(new IntPtr(501), 100u, 20u),
                Tuple.Create(window, 101u, 20u), Tuple.Create(window, 100u, 21u), Tuple.Create(window, 100u, 0u) })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateAccessAddInWindow(100, window, 20, changed.Item1, changed.Item2, changed.Item3));
        }

        [TestMethod]
        public void VisibleOwnedDialogClassOrDisabledVbeRefusesWhileRegularPrivateFormsAreAllowed()
        {
            var regular = new OfficeVbeFixture.PrivateAccessConnectionWindow { Window = 1, ThreadId = 10, Class = "WindowsForms10.Window", Visible = true, Enabled = true };
            OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(new[] { regular }, true);
            foreach (string kind in new[] { "#32770", "bosa_sdm_Microsoft Office" })
            {
                var modal = new OfficeVbeFixture.PrivateAccessConnectionWindow { Window = 2, ThreadId = 10, Class = kind, Visible = true };
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(new[] { regular, modal }, true));
                modal.Visible = false; OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(new[] { regular, modal }, true);
            }
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(new[] { regular }, false));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(null, true));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireNoPrivateAccessConnectionModal(new OfficeVbeFixture.PrivateAccessConnectionWindow[] { null }, true));
        }

        private static string State(IDictionary<string, object> row) => row.TryGetValue("State", out object state) ? Convert.ToString(state) : null;
    }
}
