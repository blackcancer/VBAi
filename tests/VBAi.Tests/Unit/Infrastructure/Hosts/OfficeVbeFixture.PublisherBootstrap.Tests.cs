using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePublisherBootstrapTests
    {
        private sealed class Harness
        {
            internal OfficeVbeFixture.PublisherBootstrapSnapshot Snapshot = Exact();
            internal readonly OfficeVbeFixture.PublisherBootstrapGate Gate;
            internal readonly List<IDictionary<string, object>> Rows = new List<IDictionary<string, object>>();
            internal long Identity = 700;
            internal int Queries, Mutations, Proofs, Saves, Releases;
            internal bool Supported = true;
            internal Action DuringMutation;
            internal Action<IDictionary<string, object>> Record;
            internal Harness()
            {
                Gate = new OfficeVbeFixture.PublisherBootstrapGate(true, Snapshot); Record = Rows.Add;
            }
            internal void Bind() => Gate.Bind(() => Snapshot, () => Identity, () => { Queries++; return Supported; }, Record);
            internal void Recheck() => Gate.Recheck(() => Snapshot, () => Identity);
            internal object Invoke(string action = "NewDocument") => Gate.Invoke(action, () => Snapshot, () => Identity,
                () => { Mutations++; DuringMutation?.Invoke(); return new object(); }, Record);
            internal void Confirm(Action proof = null) => Gate.ConfirmNativePublication(() => { Proofs++; proof?.Invoke(); });
            internal void Save() { Gate.RequireNativePublication(); Recheck(); Saves++; }
        }
        private static OfficeVbeFixture.PublisherBootstrapSnapshot Exact() => new OfficeVbeFixture.PublisherBootstrapSnapshot {
            ProcessId = 100, Handle = 300, StartedUtc = "2026-10-03T09:00:00.0000000Z", Image = @"C:\Office\MSPUB.EXE",
            SessionId = 2, OwnerThread = 42, Alive = true, PrivateWindowsVerified = true, NoVisibleModal = true, Sta = true,
            PublisherProcessIds = new[] { 100 } };

        private static void Change(OfficeVbeFixture.PublisherBootstrapSnapshot state, string change)
        {
            switch (change)
            {
                case "Process": state.ProcessId = 101; state.PublisherProcessIds = new[] { 101 }; break;
                case "Handle": state.Handle++; break;
                case "Started": state.StartedUtc = "2026-10-03T09:00:00.0000001Z"; break;
                case "Image": state.Image = @"C:\Other\MSPUB.EXE"; break;
                case "Session": state.SessionId++; break;
                case "OwnerThread": state.OwnerThread++; break;
                case "Exited": state.Alive = false; break;
                case "PrivateInventory": state.PrivateWindowsVerified = false; break;
                case "VisibleModal": state.NoVisibleModal = false; break;
                case "Apartment": state.Sta = false; break;
                case "ExtraProcess": state.PublisherProcessIds = new[] { 100, 101 }; break;
                case "NoProcess": state.PublisherProcessIds = new int[0]; break;
                case "IncompleteInventory": state.PublisherProcessIds = null; break;
                default: throw new ArgumentException(change);
            }
        }

        [TestMethod]
        public void EmptyPrelaunchInventoryIsRequiredBeforeAnyBinding()
        {
            Assert.ThrowsException<InvalidOperationException>(() => new OfficeVbeFixture.PublisherBootstrapGate(false, Exact()));
        }

        [DataTestMethod]
        [DataRow("Process")]
        [DataRow("Handle")]
        [DataRow("Started")]
        [DataRow("Image")]
        [DataRow("Session")]
        [DataRow("OwnerThread")]
        [DataRow("Exited")]
        [DataRow("PrivateInventory")]
        [DataRow("VisibleModal")]
        [DataRow("Apartment")]
        [DataRow("ExtraProcess")]
        [DataRow("NoProcess")]
        [DataRow("IncompleteInventory")]
        public void ChangedOriginalProcessOrIsolationRefusesBeforePublication(string change)
        {
            var h = new Harness(); h.Bind(); h.Snapshot = Exact(); Change(h.Snapshot, change);
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.AreEqual(0, h.Mutations); Assert.AreEqual(1, h.Queries);
            Assert.IsFalse(h.Rows.Any(r => Equals(r["PublisherBootstrap"], "PublicationMutationPending")));
        }

        [TestMethod]
        public void KnownUnsupportedApplicationInterfaceNeverCreatesPublicationOrRetriesBinding()
        {
            var h = new Harness { Supported = false };
            Assert.ThrowsException<InvalidOperationException>(() => h.Bind());
            h.Supported = true;
            Assert.ThrowsException<InvalidOperationException>(() => h.Bind());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.AreEqual(1, h.Queries); Assert.AreEqual(0, h.Mutations);
        }

        [TestMethod]
        public void UnknownApplicationInterfaceFailureNeverCreatesPublicationOrRetriesBinding()
        {
            var h = new Harness();
            Assert.ThrowsException<InvalidOperationException>(() => h.Gate.Bind(() => h.Snapshot, () => h.Identity,
                () => { h.Queries++; throw new InvalidOperationException("QI uncertain"); }, h.Record));
            Assert.ThrowsException<InvalidOperationException>(() => h.Bind());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.AreEqual(1, h.Queries); Assert.AreEqual(0, h.Mutations);
        }

        [TestMethod]
        public void CandidateChangedDuringInterfaceQueryCannotBind()
        {
            var h = new Harness();
            Assert.ThrowsException<InvalidOperationException>(() => h.Gate.Bind(() => h.Snapshot, () => h.Identity,
                () => { h.Queries++; h.Identity++; return true; }, h.Record));
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke()); Assert.AreEqual(0, h.Mutations);
        }

        [TestMethod]
        public void ZeroCanonicalIdentityNeverQueriesApplicationOrMutates()
        {
            var h = new Harness { Identity = 0 };
            Assert.ThrowsException<InvalidOperationException>(() => h.Bind());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.AreEqual(0, h.Queries); Assert.AreEqual(0, h.Mutations);
        }
        [DataTestMethod, DataRow(0L), DataRow(701L)]
        public void MissingOrChangedCandidateIdentityRefusesBeforePublication(long identity)
        {
            var h = new Harness(); h.Bind(); h.Identity = identity;
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke()); Assert.AreEqual(0, h.Mutations);
        }

        [DataTestMethod, DataRow("NewDocument"), DataRow("Open")]
        public void OriginalBootstrapMutationIsClaimedOnceAndRequiresFinalNativeProof(string action)
        {
            var h = new Harness(); h.Bind();
            Assert.ThrowsException<InvalidOperationException>(() => h.Save());
            h.DuringMutation = () => Assert.IsTrue(h.Rows.Any(r => Equals(r["PublisherBootstrap"], "PublicationMutationPending") && Equals(r["MutationInvoked"], true)));
            Assert.IsNotNull(h.Invoke(action));
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke(action));
            Assert.ThrowsException<InvalidOperationException>(() => h.Save());
            h.Confirm(); h.Save();
            Assert.AreEqual(1, h.Queries); Assert.AreEqual(1, h.Mutations); Assert.AreEqual(1, h.Proofs); Assert.AreEqual(1, h.Saves);
            Assert.IsTrue(h.Rows.All(r => Equals(r["NativeWindowAssociationVerified"], false)), "Provisional process inference must not be called native ownership.");
        }

        [TestMethod]
        public void MutationThrowPreservesUncertaintyAndForbidsProofSaveAndReplay()
        {
            var h = new Harness(); h.Bind(); h.DuringMutation = () => { throw new InvalidOperationException("Native outcome uncertain"); };
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.ThrowsException<InvalidOperationException>(() => h.Confirm());
            Assert.ThrowsException<InvalidOperationException>(() => h.Save());
            Assert.AreEqual(1, h.Mutations); Assert.IsTrue(h.Rows.Any(r => Equals(r["PublisherBootstrap"], "PublicationMutationUncertain")));
        }

        [TestMethod]
        public void FailedClaimReceiptNeverEntersNativeAndConsumesMutation()
        {
            var h = new Harness(); h.Bind(); h.Record = row => { throw new InvalidOperationException("Evidence unavailable"); };
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            h.Record = h.Rows.Add; Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.AreEqual(0, h.Mutations);
        }

        [TestMethod]
        public void OwnerChangedAfterDurableClaimStillCannotEnterMutationOrReplay()
        {
            var h = new Harness(); h.Bind();
            h.Record = row => { h.Rows.Add(row); if (Equals(row["PublisherBootstrap"], "PublicationMutationClaimed")) h.Snapshot.SessionId++; };
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            h.Snapshot = Exact(); Assert.ThrowsException<InvalidOperationException>(() => h.Invoke()); Assert.AreEqual(0, h.Mutations);
        }

        [TestMethod]
        public void ChangedCandidateAfterReturnedMutationRefusesSaveAndReplay()
        {
            var h = new Harness(); h.Bind(); h.DuringMutation = () => h.Identity++;
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.ThrowsException<InvalidOperationException>(() => h.Save()); Assert.AreEqual(1, h.Mutations);
        }

        [TestMethod]
        public void WrongNativePublicationProofCannotEnableSave()
        {
            var h = new Harness(); h.Bind(); h.Invoke();
            Assert.ThrowsException<InvalidOperationException>(() => h.Confirm(() => { throw new InvalidOperationException("Wrong native PID/root or nonmatching document"); }));
            Assert.ThrowsException<InvalidOperationException>(() => h.Save()); Assert.AreEqual(0, h.Saves);
        }

        [DataTestMethod]
        [DataRow("", "A", "A")]
        [DataRow("A", "B", "A")]
        [DataRow("A", "A", "B")]
        [DataRow("A", null, "A")]
        public void CanonicalDocumentMismatchRefusesDespiteSharedWindowOrPath(string retained, string active, string sole)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherDocumentIdentities(retained, active, sole));
        }

        [TestMethod]
        public void CanonicalReturnedActiveAndSoleDocumentMayMatch()
        { OfficeVbeFixture.RequirePublisherDocumentIdentities("ABC", "ABC", "ABC"); }

        [TestMethod]
        public void ClosedBootstrapCannotRecheckCreateSaveOrReacquire()
        {
            var h = new Harness(); h.Bind(); h.Invoke(); h.Confirm();
            h.Gate.Close(() => h.Releases++);
            Assert.IsTrue(h.Gate.Closed);
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            Assert.ThrowsException<InvalidOperationException>(() => h.Invoke());
            Assert.ThrowsException<InvalidOperationException>(() => h.Save());
            Assert.ThrowsException<InvalidOperationException>(() => h.Bind());
            Assert.ThrowsException<InvalidOperationException>(() => h.Gate.Close(() => h.Releases++));
            Assert.AreEqual(1, h.Releases); Assert.AreEqual(1, h.Queries); Assert.AreEqual(1, h.Mutations);
        }

        [TestMethod]
        public void FailedReferenceReleaseCannotBeReplayed()
        {
            var h = new Harness(); h.Bind();
            Assert.ThrowsException<InvalidOperationException>(() => h.Gate.Close(() => { h.Releases++; throw new InvalidOperationException("Release uncertain"); }));
            Assert.ThrowsException<InvalidOperationException>(() => h.Gate.Close(() => h.Releases++)); Assert.AreEqual(1, h.Releases);
        }

        private sealed class SecurityHarness
        {
            internal readonly Harness Bootstrap = new Harness();
            internal readonly OfficeVbeFixture.PublisherOpenSecurityGate Security = new OfficeVbeFixture.PublisherOpenSecurityGate();
            internal readonly List<int> WrittenValues = new List<int>();
            internal int Value, Reads, Opens;
            internal Func<int, int> Read;
            internal Action Write;
            internal Action<IDictionary<string, object>> Record;
            internal SecurityHarness(int initial)
            {
                Value = initial; Bootstrap.Bind();
                Record = Bootstrap.Rows.Add;
            }
            internal object Open()
            {
                Bootstrap.Gate.RequireAvailablePublication("Open");
                Security.Prepare(Bootstrap.Recheck, () => { Reads++; return Read == null ? Value : Read(Reads); },
                    value => { WrittenValues.Add(value); Write?.Invoke(); Value = value; }, Record);
                return Bootstrap.Gate.Invoke("Open", () => Bootstrap.Snapshot, () => Bootstrap.Identity,
                    () => { Opens++; return new object(); }, Record);
            }
        }

        [DataTestMethod, DataRow(1, 1), DataRow(2, 1), DataRow(3, 0)]
        public void PublisherOpenRequiresVerifiedForceDisableAndNeverEnablesMacros(int initial, int writes)
        {
            var h = new SecurityHarness(initial);
            h.Write = () => Assert.IsTrue(h.Bootstrap.Rows.Any(row => row.TryGetValue("PublisherOpenSecurity", out var state) && Equals(state, "ForceDisablePending")));
            Assert.IsNotNull(h.Open());
            Assert.AreEqual(2, h.Reads); Assert.AreEqual(writes, h.WrittenValues.Count);
            Assert.IsTrue(h.WrittenValues.All(value => value == 3)); Assert.AreEqual(3, h.Value);
            Assert.AreEqual(1, h.Opens); Assert.IsTrue(h.Security.Verified); Assert.IsTrue(h.Security.Consumed);
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(writes, h.WrittenValues.Count); Assert.AreEqual(1, h.Opens);
        }

        [DataTestMethod, DataRow(-1), DataRow(0), DataRow(4)]
        public void UnknownPublisherAutomationSecurityRefusesBeforeSetterAndOpen(int initial)
        {
            var h = new SecurityHarness(initial);
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            h.Value = 3; Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(1, h.Reads); Assert.AreEqual(0, h.WrittenValues.Count); Assert.AreEqual(0, h.Opens);
            Assert.IsTrue(h.Security.Consumed); Assert.IsFalse(h.Security.Verified);
        }

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2), DataRow(4)]
        public void UnverifiedPublisherSecurityReadbackNeverOpensOrRepeatsSetter(int final)
        {
            var h = new SecurityHarness(1) { Read = number => number == 1 ? 1 : final };
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            h.Read = null; Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(2, h.Reads); CollectionAssert.AreEqual(new[] { 3 }, h.WrittenValues.ToArray());
            Assert.AreEqual(0, h.Opens); Assert.IsFalse(h.Security.Verified);
        }

        [DataTestMethod, DataRow(1), DataRow(2)]
        public void FailedPublisherSecurityGetterPreservesFailureAndConsumesPreparation(int failedRead)
        {
            var error = new InvalidOperationException("Original security getter failed");
            var h = new SecurityHarness(1) { Read = number => number == failedRead ? throw error : 1 };
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => h.Open()));
            h.Read = null; Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(failedRead, h.Reads); Assert.AreEqual(failedRead == 1 ? 0 : 1, h.WrittenValues.Count);
            Assert.AreEqual(0, h.Opens); Assert.IsTrue(h.Security.Consumed); Assert.IsFalse(h.Security.Verified);
        }

        [TestMethod]
        public void UncertainPublisherForceDisableSetterCannotBeRepeatedOrFollowedByOpen()
        {
            var error = new InvalidOperationException("Original native security setter outcome unknown");
            var h = new SecurityHarness(2) { Write = () => { throw error; } };
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => h.Open()));
            h.Write = null; Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(1, h.Reads); CollectionAssert.AreEqual(new[] { 3 }, h.WrittenValues.ToArray());
            Assert.AreEqual(0, h.Opens); Assert.IsFalse(h.Security.Verified);
            Assert.IsTrue(h.Bootstrap.Rows.Any(row => row.TryGetValue("PublisherOpenSecurity", out var state) && Equals(state, "ForceDisableUncertain")));
        }

        [TestMethod]
        public void ReturnedPublisherSecuritySetterRecordsMutationBeforeChangedOwnerRefusesOpen()
        {
            var h = new SecurityHarness(2);
            long originalIdentity = h.Bootstrap.Identity;
            h.Write = () => h.Bootstrap.Identity++;
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            var returned = h.Bootstrap.Rows.Single(row => row.TryGetValue("PublisherOpenSecurity", out var state) && Equals(state, "ForceDisableReturned"));
            Assert.AreEqual(true, returned["MutationInvoked"]);
            Assert.AreEqual(1, h.Reads); CollectionAssert.AreEqual(new[] { 3 }, h.WrittenValues.ToArray());
            Assert.AreEqual(0, h.Opens); Assert.IsTrue(h.Security.Consumed); Assert.IsFalse(h.Security.Verified);
            h.Bootstrap.Identity = originalIdentity; h.Write = null;
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(1, h.Reads); Assert.AreEqual(1, h.WrittenValues.Count); Assert.AreEqual(0, h.Opens);
        }

        [DataTestMethod]
        [DataRow("PreparationClaimed", 0)]
        [DataRow("InitialReadPending", 0)]
        [DataRow("InitialReadReturned", 0)]
        [DataRow("ForceDisableClaimed", 0)]
        [DataRow("ForceDisablePending", 0)]
        [DataRow("ForceDisableReturned", 1)]
        [DataRow("FinalReadPending", 1)]
        [DataRow("FinalReadReturned", 1)]
        [DataRow("ForceDisableVerified", 1)]
        public void PublisherOwnerDriftAtSecurityStagePreventsOpenAndAnyPreparationReplay(string stage, int writes)
        {
            var h = new SecurityHarness(1);
            h.Record = row => { h.Bootstrap.Rows.Add(row); if (row.TryGetValue("PublisherOpenSecurity", out var state) && Equals(state, stage)) h.Bootstrap.Identity++; };
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            h.Bootstrap.Identity = 700; h.Record = h.Bootstrap.Rows.Add;
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(writes, h.WrittenValues.Count); Assert.AreEqual(0, h.Opens); Assert.IsFalse(h.Security.Verified);
        }

        [DataTestMethod, DataRow("ForceDisableClaimed"), DataRow("ForceDisablePending")]
        public void FailedPublisherSecurityMutationReceiptPreventsSetterAndOpen(string failedStage)
        {
            var h = new SecurityHarness(1);
            h.Record = row => { if (row.TryGetValue("PublisherOpenSecurity", out var state) && Equals(state, failedStage)) throw new InvalidOperationException("No durable mutation evidence"); };
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            h.Record = h.Bootstrap.Rows.Add; Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(0, h.WrittenValues.Count); Assert.AreEqual(0, h.Opens); Assert.IsTrue(h.Security.Consumed);
        }

        [TestMethod]
        public void AlreadyForceDisabledPublisherStillRequiresFreshFinalReadback()
        {
            var h = new SecurityHarness(3) { Read = number => number == 1 ? 3 : 2 };
            Assert.ThrowsException<InvalidOperationException>(() => h.Open());
            Assert.AreEqual(2, h.Reads); Assert.AreEqual(0, h.WrittenValues.Count); Assert.AreEqual(0, h.Opens);
            Assert.IsFalse(h.Security.Verified);
        }

        [TestMethod]
        public void ClosedOrPreviouslyUsedPublisherBootstrapNeverPreparesOpenSecurity()
        {
            var closed = new SecurityHarness(1); closed.Bootstrap.Gate.Close(() => { });
            Assert.ThrowsException<InvalidOperationException>(() => closed.Open());
            Assert.AreEqual(0, closed.Reads); Assert.AreEqual(0, closed.WrittenValues.Count);
            var used = new SecurityHarness(1); used.Bootstrap.Invoke("NewDocument");
            Assert.ThrowsException<InvalidOperationException>(() => used.Open());
            Assert.AreEqual(0, used.Reads); Assert.AreEqual(0, used.WrittenValues.Count);
        }

        [TestMethod]
        public void PublisherSecurityPreparationIsIndependentForEachNewOriginalGeneration()
        {
            var first = new SecurityHarness(2); first.Open();
            var second = new SecurityHarness(1); second.Open();
            Assert.AreEqual(1, first.WrittenValues.Count); Assert.AreEqual(1, second.WrittenValues.Count);
            Assert.AreEqual(1, first.Opens); Assert.AreEqual(1, second.Opens);
        }
    }
}
