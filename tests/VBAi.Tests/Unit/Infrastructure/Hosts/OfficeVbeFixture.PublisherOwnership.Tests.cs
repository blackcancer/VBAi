using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePublisherOwnershipTests
    {
        private sealed class Harness
        {
            internal readonly OfficeVbeFixture.PublisherOwnershipProbe Probe = new OfficeVbeFixture.PublisherOwnershipProbe(100, 300);
            internal readonly List<IDictionary<string, object>> Rows = new List<IDictionary<string, object>>();
            internal int Queries, GetWindows, BarInventories, NativeOms, Identities, Guards;
            internal int QueryResult;
            internal long Identity = 700, BarIdentity = 700;
            internal OfficeVbeFixture.PublisherOwnerWindow Window = Exact();
            internal OfficeVbeFixture.PublisherOwnerWindow[] Bars;
            internal Func<IntPtr, OfficeVbeFixture.PublisherOwnerWindow> ReadWindow;
            internal Func<OfficeVbeFixture.PublisherOwnerWindow, long> ReadBarIdentity;
            internal Action Guard;
            internal Action<IDictionary<string, object>> Record;
            internal Harness()
            {
                Bars = new[] { Window };
                ReadWindow = handle => Copy(Bars.FirstOrDefault(w => w != null && w.Window == handle) ?? Window);
                ReadBarIdentity = window => BarIdentity; Guard = () => Guards++; Record = Rows.Add;
            }
            internal void Bind()
            {
                Probe.Bind(() => { Identities++; return Identity; }, Guard, () => { Queries++; return QueryResult; },
                    () => { GetWindows++; return Window; }, () => { BarInventories++; return Bars; },
                    window => { NativeOms++; return ReadBarIdentity(window); }, ReadWindow, Record);
            }
            internal void Recheck(uint pid = 100, long handle = 300)
            { Probe.Recheck(pid, handle, () => { Identities++; return Identity; }, Guard, ReadWindow); }
        }

        private static OfficeVbeFixture.PublisherOwnerWindow Exact() => new OfficeVbeFixture.PublisherOwnerWindow
        {
            Window = new IntPtr(11),
            Root = new IntPtr(10),
            Process = 100,
            RootProcess = 100,
            Thread = 200,
            RootThread = 201,
            Class = "MsoCommandBar"
        };

        private static OfficeVbeFixture.PublisherOwnerWindow Copy(OfficeVbeFixture.PublisherOwnerWindow w) =>
            w == null ? null : new OfficeVbeFixture.PublisherOwnerWindow
            {
                Window = w.Window,
                Root = w.Root,
                Process = w.Process,
                RootProcess = w.RootProcess,
                Thread = w.Thread,
                RootThread = w.RootThread,
                Class = w.Class
            };

        private static OfficeVbeFixture.PublisherOwnerWindow Bar(long hwnd)
        { var value = Exact(); value.Window = new IntPtr(hwnd); return value; }

        private sealed class NativeBarHarness
        {
            internal readonly OfficeVbeFixture.PublisherCommandBarAcquisition Acquisition = new OfficeVbeFixture.PublisherCommandBarAcquisition();
            internal readonly OfficeVbeFixture.PublisherOwnershipProbe Probe = new OfficeVbeFixture.PublisherOwnershipProbe(100, 300);
            internal readonly List<IDictionary<string, object>> Rows = new List<IDictionary<string, object>>();
            internal readonly List<object> ReleasedObjects = new List<object>();
            internal readonly object Bar = new object(), Container = new object();
            internal int Acquires, Wraps, Getters, Identities, RawReleases, Guards;
            internal Action<OfficeVbeFixture.PublisherNativeOmResult> Acquire;
            internal Func<IntPtr, object> Wrap;
            internal Func<object, bool> Typed = value => true;
            internal Func<object, object> Getter;
            internal Func<object, long> Identity;
            internal Action<object> ReleaseObject;
            internal Action Guard;
            internal Action<IDictionary<string, object>> Record;

            internal NativeBarHarness()
            {
                Acquire = result => { result.HResult = 0; result.Pointer = new IntPtr(900); };
                Wrap = pointer => Bar; Getter = bar => Container; Identity = container => 700;
                ReleaseObject = ReleasedObjects.Add; Guard = () => Guards++; Record = Rows.Add;
            }
            internal long Read() => Acquisition.Read(Probe, "Bar11", Guard,
                result => { Acquires++; Acquire(result); }, pointer => { Wraps++; return Wrap(pointer); }, Typed,
                bar => { Getters++; return Getter(bar); }, container => { Identities++; return Identity(container); },
                pointer => { Assert.AreEqual(new IntPtr(900), pointer); RawReleases++; }, ReleaseObject, Record);
        }

        [TestMethod]
        public void TypedNativeOmAcquiresReadsAndBalancesEachReferenceOnce()
        {
            var h = new NativeBarHarness(); Assert.AreEqual(700L, h.Read());
            Assert.AreEqual(1, h.Acquires); Assert.AreEqual(1, h.Wraps); Assert.AreEqual(1, h.Getters);
            Assert.AreEqual(1, h.Identities); Assert.AreEqual(1, h.RawReleases);
            CollectionAssert.AreEqual(new[] { h.Container, h.Bar }, h.ReleasedObjects.ToArray());
            var native = (IDictionary<string, object>)h.Rows.Single(r => Equals(r["PublisherOwnershipRead"], "Bar11.NativeOM") && Equals(r["State"], "RETURNED"))["Value"];
            Assert.AreEqual("0x00000000", native["HResult"]); Assert.AreEqual(900L, native["Pointer"]);
            Assert.AreEqual("000c0304-0000-0000-c000-000000000046", native["RequestedInterfaceId"]);
            var typed = (IDictionary<string, object>)h.Rows.Single(r => Equals(r["PublisherOwnershipRead"], "Bar11.TypedInterface") && Equals(r["State"], "RETURNED"))["Value"];
            Assert.AreEqual("0x60020000", typed["ApplicationDispId"]); Assert.AreEqual(true, typed["TypedInterfaceSupported"]);
            Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Acquires);
        }

        [TestMethod]
        public void FailedNativeOmOrNullPointerNeverMarshalOrReadApplication()
        {
            foreach (var outcome in new[] { Tuple.Create(unchecked((int)0x80004002), 0L),
                Tuple.Create(unchecked((int)0x80004005), 900L), Tuple.Create(0, 0L), Tuple.Create(1, 900L) })
            {
                var h = new NativeBarHarness();
                h.Acquire = result => { result.HResult = outcome.Item1; result.Pointer = new IntPtr(outcome.Item2); };
                Assert.ThrowsException<COMException>(() => h.Read());
                Assert.AreEqual(0, h.Wraps); Assert.AreEqual(0, h.Getters); Assert.AreEqual(0, h.Identities);
                Assert.AreEqual(outcome.Item2 == 0 ? 0 : 1, h.RawReleases); Assert.AreEqual(0, h.ReleasedObjects.Count);
                Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Acquires);
            }
        }

        [TestMethod]
        public void NativeAcquireExceptionBalancesReturnedPointerWithoutFallback()
        {
            var h = new NativeBarHarness(); h.Acquire = result => { result.Pointer = new IntPtr(900); throw new COMException("Acquire failed"); };
            Assert.ThrowsException<COMException>(() => h.Read()); Assert.AreEqual(1, h.RawReleases);
            Assert.AreEqual(0, h.Wraps); Assert.AreEqual(0, h.Getters);
            Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Acquires);
        }

        [TestMethod]
        public void MissingOrUnsupportedTypedWrapperRefusesBeforeApplicationGetter()
        {
            foreach (int outcome in new[] { 0, 1, 2, 3 })
            {
                var h = new NativeBarHarness();
                if (outcome == 0) h.Wrap = pointer => null;
                if (outcome == 1) h.Wrap = pointer => { throw new InvalidCastException("Cannot wrap"); };
                if (outcome == 2) h.Typed = value => false;
                if (outcome == 3) h.Typed = value => { throw new InvalidCastException("Missing exact interface"); };
                try { h.Read(); Assert.Fail("Unsupported wrapper must refuse."); }
                catch (Exception error) { Assert.IsTrue(error is InvalidOperationException || error is InvalidCastException); }
                Assert.AreEqual(0, h.Getters); Assert.AreEqual(1, h.RawReleases);
                Assert.AreEqual(outcome >= 2 ? 1 : 0, h.ReleasedObjects.Count);
                Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Acquires);
            }
        }

        [TestMethod]
        public void TypedApplicationFailureOrNullContainerNeverReadsIdentity()
        {
            foreach (bool throws in new[] { false, true })
            {
                var h = new NativeBarHarness(); h.Getter = bar => { if (throws) throw new COMException("Getter failed"); return null; };
                if (throws) Assert.ThrowsException<COMException>(() => h.Read());
                else Assert.ThrowsException<InvalidOperationException>(() => h.Read());
                Assert.AreEqual(1, h.Getters); Assert.AreEqual(0, h.Identities); Assert.AreEqual(1, h.RawReleases);
                CollectionAssert.AreEqual(new[] { h.Bar }, h.ReleasedObjects.ToArray());
                Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Getters);
            }
        }

        [TestMethod]
        public void IdentityFailureOrMissingIdentityBalancesBothReturnedObjects()
        {
            foreach (bool throws in new[] { false, true })
            {
                var h = new NativeBarHarness(); h.Identity = container => { if (throws) throw new InvalidCastException("Not COM"); return 0; };
                if (throws) Assert.ThrowsException<InvalidCastException>(() => h.Read());
                else Assert.ThrowsException<InvalidOperationException>(() => h.Read());
                CollectionAssert.AreEqual(new[] { h.Container, h.Bar }, h.ReleasedObjects.ToArray());
                Assert.AreEqual(1, h.RawReleases); Assert.ThrowsException<InvalidOperationException>(() => h.Read());
            }
        }

        [TestMethod]
        public void ChangedOwnerGuardStopsNextNativeStageAndStillBalancesReferences()
        {
            foreach (int blockedGuard in new[] { 1, 2, 3, 4, 5 })
            {
                var h = new NativeBarHarness(); h.Guard = () => { if (++h.Guards == blockedGuard) throw new InvalidOperationException("Owner changed"); };
                Assert.ThrowsException<InvalidOperationException>(() => h.Read());
                Assert.AreEqual(blockedGuard > 1 ? 1 : 0, h.Acquires);
                Assert.AreEqual(blockedGuard > 2 ? 1 : 0, h.Wraps);
                Assert.AreEqual(blockedGuard > 3 ? 1 : 0, h.Getters);
                Assert.AreEqual(blockedGuard > 4 ? 1 : 0, h.Identities);
                Assert.AreEqual(blockedGuard > 1 ? 1 : 0, h.RawReleases);
                Assert.AreEqual((blockedGuard > 2 ? 1 : 0) + (blockedGuard > 3 ? 1 : 0), h.ReleasedObjects.Count);
                Assert.ThrowsException<InvalidOperationException>(() => h.Read());
            }
        }

        [TestMethod]
        public void DurableReturnedEvidenceFailureCannotLeakAcquiredReferencesOrReplay()
        {
            foreach (string stage in new[] { "NativeOM", "TypedInterface", "Application", "ApplicationIUnknown" })
            {
                var h = new NativeBarHarness(); h.Record = row =>
                {
                    if (Equals(row["PublisherOwnershipRead"], "Bar11." + stage) && Equals(row["State"], "RETURNED"))
                        throw new InvalidOperationException("Evidence failed");
                    h.Rows.Add(row);
                };
                Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.RawReleases);
                Assert.AreEqual(stage == "NativeOM" ? 0 : stage == "TypedInterface" ? 1 : 2, h.ReleasedObjects.Count);
                Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.Acquires);
            }
        }

        [TestMethod]
        public void FailingReferenceReleaseStillBalancesOthersAndPreservesOriginalFailure()
        {
            var h = new NativeBarHarness(); h.Identity = value => { throw new COMException("Identity failed"); };
            h.ReleaseObject = value => { h.ReleasedObjects.Add(value); if (ReferenceEquals(value, h.Container)) throw new InvalidOperationException("Release failed"); };
            var failure = Assert.ThrowsException<AggregateException>(() => h.Read());
            Assert.AreEqual(2, failure.InnerExceptions.Count); Assert.IsInstanceOfType(failure.InnerExceptions[0], typeof(COMException));
            CollectionAssert.AreEqual(new[] { h.Container, h.Bar }, h.ReleasedObjects.ToArray()); Assert.AreEqual(1, h.RawReleases);
            Assert.ThrowsException<InvalidOperationException>(() => h.Read()); Assert.AreEqual(1, h.RawReleases);
        }

        [TestMethod]
        public void ObservedOleWindowBindsOnceAndRechecksWithoutRediscovery()
        {
            var h = new Harness(); h.Bind(); h.Recheck(); h.Recheck();
            Assert.AreEqual(1, h.Queries); Assert.AreEqual(1, h.GetWindows);
            Assert.AreEqual(0, h.BarInventories); Assert.AreEqual(0, h.NativeOms);
            Assert.AreEqual(4, h.Identities);
            var row = h.Rows.Last(); Assert.AreEqual("VERIFIED", row["PublisherPrePublicationOwnership"]);
            Assert.AreEqual("ObservedIOleWindow", row["Route"]); Assert.AreEqual(200u, row["NativeThreadId"]);
            var returned = h.Rows.Single(r => r.ContainsKey("PublisherOwnershipRead") && Equals(r["PublisherOwnershipRead"], "IOleWindow.GetWindow") && Equals(r["State"], "RETURNED"));
            var window = (IDictionary<string, object>)returned["Value"];
            Assert.AreEqual(11L, window["Hwnd"]); Assert.AreEqual(100u, window["ProcessId"]);
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.Queries);
        }

        [TestMethod]
        public void OnlyKnownNoInterfaceAllowsExactCommandBarApplicationChain()
        {
            var h = new Harness { QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface };
            h.Bind(); h.Recheck(); Assert.AreEqual(1, h.Queries); Assert.AreEqual(0, h.GetWindows);
            Assert.AreEqual(1, h.BarInventories); Assert.AreEqual(1, h.NativeOms);
            Assert.AreEqual("OwnedMsoCommandBarNativeOM", h.Rows.Last()["Route"]);
            Assert.AreEqual(1, h.Rows.Count(r => r.ContainsKey("PublisherOwnershipCapability") && Equals(r["State"], "KNOWN_UNSUPPORTED")));
        }

        [TestMethod]
        public void UnknownQueryFailureNeverEntersNativeOmOrReplaysDiscovery()
        {
            foreach (int result in new[] { unchecked((int)0x80004005), unchecked((int)0x80010001), 1 })
            {
                var h = new Harness { QueryResult = result };
                Assert.ThrowsException<COMException>(h.Bind);
                Assert.AreEqual(0, h.GetWindows); Assert.AreEqual(0, h.BarInventories); Assert.AreEqual(0, h.NativeOms);
                Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.Queries);
                Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            }
        }

        [TestMethod]
        public void OleWindowFailureNeverFallsBackEvenWhenACommandBarExists()
        {
            var h = new Harness(); h.Window = null;
            Assert.ThrowsException<InvalidOperationException>(h.Bind);
            Assert.AreEqual(1, h.GetWindows); Assert.AreEqual(0, h.BarInventories);
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.GetWindows);
        }

        [TestMethod]
        public void MissingDuplicateUndocumentedOrUnboundedBarsRefuseNativeOm()
        {
            foreach (var bars in new[] { null, new OfficeVbeFixture.PublisherOwnerWindow[0], new[] { Exact(), Exact() },
                Enumerable.Range(1, OfficeVbeFixture.PublisherOwnershipProbe.MaximumCommandBars + 1).Select(i => Bar(i)).ToArray(),
                new[] { new OfficeVbeFixture.PublisherOwnerWindow { Window = new IntPtr(11), Root = new IntPtr(10),
                    Process = 100, RootProcess = 100, Thread = 200, RootThread = 201, Class = "PublisherMainWindow" } } })
            {
                var h = new Harness { QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface, Bars = bars };
                Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(0, h.NativeOms);
                Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            }
        }

        [TestMethod]
        public void WrongWindowRootPidOrMissingNativeThreadRefusesBinding()
        {
            foreach (var change in new Action<OfficeVbeFixture.PublisherOwnerWindow>[] {
                w => w.Window = IntPtr.Zero, w => w.Root = IntPtr.Zero, w => w.Process = 101,
                w => w.RootProcess = 101, w => w.Thread = 0, w => w.RootThread = 0 })
            {
                var h = new Harness(); change(h.Window);
                Assert.ThrowsException<InvalidOperationException>(h.Bind);
                Assert.AreEqual(0, h.BarInventories); Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            }
        }

        [TestMethod]
        public void ChangedBarBeforeNativeOmAndForeignContainerNeverBind()
        {
            var changed = new Harness { QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface };
            changed.ReadWindow = handle => { var w = Exact(); w.RootThread++; return w; };
            Assert.ThrowsException<InvalidOperationException>(changed.Bind); Assert.AreEqual(0, changed.NativeOms);
            var foreign = new Harness { QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface, BarIdentity = 701 };
            Assert.ThrowsException<InvalidOperationException>(foreign.Bind); Assert.AreEqual(1, foreign.NativeOms);
            Assert.ThrowsException<InvalidOperationException>(() => foreign.Recheck());
        }

        [TestMethod]
        public void WindowChangingAfterGetterRefusesAndCannotBeRediscovered()
        {
            var h = new Harness(); h.ReadWindow = handle => { var w = Exact(); w.Thread++; return w; };
            Assert.ThrowsException<InvalidOperationException>(h.Bind);
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.Queries);
        }

        [TestMethod]
        public void RecheckRefusesChangedOriginalProcessApplicationOrBoundWindow()
        {
            var h = new Harness(); h.Bind();
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck(101));
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck(100, 301));
            h.Identity = 701; Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            h.Identity = 700; h.ReadWindow = handle => { var w = Exact(); w.Class = "ChangedClass"; return w; };
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            Assert.AreEqual(1, h.Queries); Assert.AreEqual(1, h.GetWindows);
        }

        [TestMethod]
        public void OwnerThreadOrDesktopGuardFailureConsumesProbeWithoutQuery()
        {
            var h = new Harness(); h.Guard = () => { throw new InvalidOperationException("Acquiring STA/private desktop mismatch"); };
            Assert.ThrowsException<InvalidOperationException>(h.Bind);
            Assert.AreEqual(0, h.Identities); Assert.AreEqual(0, h.Queries);
            h.Guard = () => { }; Assert.ThrowsException<InvalidOperationException>(h.Bind);
            Assert.AreEqual(0, h.Queries);
        }

        [TestMethod]
        public void MissingIdentityOrDurableIntentFailureCannotEnterQuery()
        {
            var missing = new Harness { Identity = 0 };
            Assert.ThrowsException<InvalidOperationException>(missing.Bind); Assert.AreEqual(0, missing.Queries);
            var evidence = new Harness(); evidence.Record = row => { throw new InvalidOperationException("Evidence unavailable"); };
            Assert.ThrowsException<InvalidOperationException>(evidence.Bind); Assert.AreEqual(0, evidence.Identities);
            Assert.AreEqual(0, evidence.Queries); Assert.ThrowsException<InvalidOperationException>(evidence.Bind);
        }

        [TestMethod]
        public void ReadAndFailureEvidenceErrorsRemainTogetherAndForbidReplay()
        {
            var h = new Harness(); int writes = 0;
            h.Record = row => { if (++writes == 4) throw new InvalidOperationException("Failure evidence unavailable"); h.Rows.Add(row); };
            // Application identity writes consume entries 1/2; query intent is entry 3.
            h.Guard = () => { };
            var probe = h.Probe;
            var error = Assert.ThrowsException<AggregateException>(() => probe.Bind(() => 700, () => { },
                () => { throw new COMException("Query rejected", unchecked((int)0x80010001)); }, Exact,
                () => { Assert.Fail("No fallback"); return new[] { Exact() }; }, w => 700, w => Exact(), h.Record));
            Assert.AreEqual(2, error.InnerExceptions.Count);
            Assert.ThrowsException<InvalidOperationException>(h.Bind);
        }

        [TestMethod]
        public void ClosedOwnershipReleasesOnceAndForbidsRecheckOrReacquisition()
        {
            var h = new Harness(); h.Bind(); int releases = 0;
            h.Probe.Close(() => releases++);
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            Assert.ThrowsException<InvalidOperationException>(h.Bind);
            Assert.ThrowsException<InvalidOperationException>(() => h.Probe.Close(() => releases++));
            Assert.AreEqual(1, releases); Assert.AreEqual(1, h.Queries); Assert.AreEqual(1, h.GetWindows);
        }

        [TestMethod]
        public void UncertainReferenceReleaseClosesProbeBeforeFailureAndCannotReplay()
        {
            var h = new Harness(); h.Bind(); int releases = 0;
            Assert.ThrowsException<InvalidOperationException>(() => h.Probe.Close(() => { releases++; throw new InvalidOperationException("Release failed"); }));
            Assert.ThrowsException<InvalidOperationException>(() => h.Probe.Close(() => releases++));
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            Assert.AreEqual(1, releases);
        }

        [TestMethod]
        public void MultipleDistinctBarsPermitKnownMismatchThenExactApplicationMatch()
        {
            var h = new Harness
            {
                QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface,
                Bars = new[] { Bar(11), Bar(12), Bar(13) }
            };
            var observed = new List<IntPtr>();
            h.ReadBarIdentity = window => { observed.Add(window.Window); return window.Window.ToInt64() == 12 ? 700 : 701; };
            h.Bind(); h.Recheck();
            CollectionAssert.AreEqual(new[] { new IntPtr(11), new IntPtr(12) }, observed.ToArray());
            Assert.AreEqual(2, h.NativeOms); Assert.AreEqual(1, h.Queries); Assert.AreEqual(0, h.GetWindows);
            Assert.AreEqual(12L, h.Rows.Last()["Hwnd"]);
            var associations = h.Rows.Where(row => row.ContainsKey("PublisherCommandBarAssociation")).ToArray();
            Assert.AreEqual("KNOWN_MISMATCH", associations[0]["PublisherCommandBarAssociation"]);
            Assert.AreEqual("EXACT_MATCH", associations[1]["PublisherCommandBarAssociation"]);
        }

        [TestMethod]
        public void InvalidLaterInventoryWindowRefusesBeforeAnyNativeObjectModelCall()
        {
            foreach (var change in new Action<OfficeVbeFixture.PublisherOwnerWindow>[] {
                w => w.Window = new IntPtr(11), w => w.Root = IntPtr.Zero, w => w.Process = 101,
                w => w.RootProcess = 101, w => w.Thread = 0, w => w.RootThread = 0, w => w.Class = "OtherClass" })
            {
                var later = Bar(12); change(later);
                var h = new Harness
                {
                    QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface,
                    Bars = new[] { Bar(11), later }
                };
                Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(0, h.NativeOms);
            }
        }

        [TestMethod]
        public void FirstBarUnknownFailureNeverAttemptsAnotherWindowOrReplays()
        {
            var h = new Harness
            {
                QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface,
                Bars = new[] { Bar(11), Bar(12) }
            };
            h.ReadBarIdentity = window => { throw new COMException("NativeOM or container getter failed", unchecked((int)0x80004005)); };
            Assert.ThrowsException<COMException>(h.Bind); Assert.AreEqual(1, h.NativeOms);
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.NativeOms);
            Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
        }

        [TestMethod]
        public void AllKnownContainerMismatchesRemainBlockedWithoutApplicationAdoption()
        {
            var h = new Harness
            {
                QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface,
                Bars = new[] { Bar(11), Bar(12) },
                BarIdentity = 701
            };
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(2, h.NativeOms);
            Assert.AreEqual(2, h.Rows.Count(row => row.ContainsKey("PublisherCommandBarAssociation") &&
                Equals(row["PublisherCommandBarAssociation"], "KNOWN_MISMATCH")));
            Assert.AreEqual(700L, h.Identity); Assert.ThrowsException<InvalidOperationException>(() => h.Recheck());
            Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(2, h.NativeOms);
        }

        [TestMethod]
        public void MissingContainerOrChangedCanonicalApplicationStopsBeforeNextBar()
        {
            foreach (bool changedApplication in new[] { false, true })
            {
                var h = new Harness
                {
                    QueryResult = OfficeVbeFixture.PublisherOwnershipProbe.NoInterface,
                    Bars = new[] { Bar(11), Bar(12) }
                };
                h.ReadBarIdentity = window => { if (changedApplication) h.Identity = 702; return changedApplication ? 701 : 0; };
                Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.NativeOms);
                Assert.ThrowsException<InvalidOperationException>(h.Bind); Assert.AreEqual(1, h.NativeOms);
            }
        }
    }
}
