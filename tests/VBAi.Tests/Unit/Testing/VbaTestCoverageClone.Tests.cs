using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestCoverageCloneTests
    {
        private sealed class Fixture : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-Excel-Coverage-" + Guid.NewGuid().ToString("N"));
            internal readonly Application Application = new Application();
            internal readonly Book Source;
            internal readonly VbaTestCoverageClone.HostBoundary Boundary = new VbaTestCoverageClone.HostBoundary();
            internal Fixture()
            {
                Source = new Book { FullName = Path.Combine(Folder, "Original.xlsm"), Application = Application };
                Application.Workbooks.Add(Source);
                Boundary.ReadProcessName = () => "EXCEL";
                Boundary.ReadProcessId = () => 123;
                Boundary.ResolveExcel = id => { Assert.AreEqual(123, id); return Application; };
                Boundary.ReadWindowOwner = hwnd => { Assert.AreEqual(new IntPtr(99), hwnd); return 123; };
                Boundary.SameIdentity = ReferenceEquals;
            }
            internal VbaTestCoverageClone Create() => VbaTestCoverageClone.CreateOwned(Source.VBProject, Source.FullName, Path.Combine(Folder, "Copy"), Boundary);
            public void Dispose() { if (Directory.Exists(Folder)) Directory.Delete(Folder, true); }
        }
        public sealed class Application
        {
            public int Hwnd => 99;
            public readonly Books Workbooks = new Books();
            public bool IgnoreEventsWrite;
            public Action<bool> OnEventsWrite;
            private bool events = true;
            public bool EnableEvents { get => events; set { OnEventsWrite?.Invoke(value); if (!IgnoreEventsWrite) events = value; } }
        }
        public sealed class Books : List<Book>, IEnumerable
        {
            public bool NativeIndexing, RejectEnumeration;
            public int EnumeratorRequests;
            public readonly List<int> IndexedReads = new List<int>();
            public new Book this[int index]
            {
                get { IndexedReads.Add(index); return base[NativeIndexing ? index - 1 : index]; }
            }
            public Book RawItem(int index) => base[index];
            IEnumerator IEnumerable.GetEnumerator()
            {
                EnumeratorRequests++;
                if (RejectEnumeration) throw new InvalidOperationException("An opaque native enumeration adapter must not be created.");
                return base.GetEnumerator();
            }
            public Func<string, Book> OnOpen;
            public int Opens;
            public Book Open(string path, int UpdateLinks, bool ReadOnly, bool AddToMru)
            {
                Assert.AreEqual(0, UpdateLinks); Assert.IsFalse(ReadOnly); Assert.IsFalse(AddToMru);
                Opens++;
                var book = OnOpen == null ? new Book { FullName = path, Application = base[0].Application } : OnOpen(path);
                Add(book); return book;
            }
        }
        [TestMethod]
        public void RepeatedIndexedCopiesNeverAcquireAHiddenEnumeratorAndCloseOnlyTheirOwnedWorkbook()
        {
            using (var fixture = new Fixture())
            {
                var books = fixture.Application.Workbooks;
                books.NativeIndexing = true; books.RejectEnumeration = true;
                fixture.Boundary.IndexedCollection = _ => true;
                var releases = new List<object>(); fixture.Boundary.ReleaseAcquired = releases.Add;
                for (int cycle = 0; cycle < 3; cycle++)
                {
                    var clone = VbaTestCoverageClone.CreateOwned(fixture.Source.VBProject, fixture.Source.FullName,
                        Path.Combine(fixture.Folder, "Indexed-" + cycle), fixture.Boundary);
                    var copy = books.RawItem(1);
                    clone.Dispose(); clone.Dispose();
                    Assert.AreEqual(1, books.Count); Assert.AreSame(fixture.Source, books.RawItem(0));
                    Assert.AreEqual(1, copy.Closes); Assert.AreEqual(0, fixture.Source.Closes);
                    Assert.AreEqual(1, releases.FindAll(value => ReferenceEquals(value, copy.VBProject)).Count);
                }
                Assert.AreEqual(0, books.EnumeratorRequests);
                CollectionAssert.AreEqual(new[] { 1, 1, 1, 1, 1, 1 }, books.IndexedReads);
                Assert.IsTrue(fixture.Application.EnableEvents);
            }
        }

        [TestMethod]
        public void UncertainIndexedCloseRetainsTheCopyAndNeverFallsBackOrClosesAgain()
        {
            using (var fixture = new Fixture())
            {
                var books = fixture.Application.Workbooks;
                books.NativeIndexing = true; books.RejectEnumeration = true;
                fixture.Boundary.IndexedCollection = _ => true;
                var releases = new List<object>(); fixture.Boundary.ReleaseAcquired = releases.Add;
                var clone = fixture.Create(); var copy = books.RawItem(1); copy.CancelClose = true;
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, copy.Closes); Assert.AreEqual(0, fixture.Source.Closes);
                Assert.IsFalse(releases.Contains(copy)); Assert.IsFalse(releases.Contains(fixture.Application));
                Assert.AreEqual(0, books.EnumeratorRequests); Assert.AreEqual(2, books.Count);
                Assert.IsTrue(fixture.Application.EnableEvents);
            }
        }
        public sealed class Book
        {
            public object VBProject { get; set; } = new object();
            public string FullName { get; set; }
            public Application Application;
            public Action OnSaveCopy, OnClose;
            public bool CancelClose;
            public int Saves, Closes;
            public void SaveCopyAs(string path) { Assert.IsFalse(Application.EnableEvents); Saves++; OnSaveCopy?.Invoke(); File.WriteAllText(path, "Disposable workbook bytes"); }
            public void Close(bool save) { Assert.IsFalse(save); Assert.IsFalse(Application.EnableEvents); Closes++; OnClose?.Invoke(); if (!CancelClose) Application.Workbooks.Remove(this); }
        }

        [TestMethod]
        public void CopyAndCloseRestoreBothOriginalEventStatesAndPreserveOriginalWorkbook()
        {
            foreach (bool events in new[] { false, true })
                foreach (string extension in new[] { ".xlsm", ".xlsb", ".xls" })
                    using (var f = new Fixture())
                    {
                        f.Source.FullName = Path.ChangeExtension(f.Source.FullName, extension);
                        f.Application.EnableEvents = events;
                        var clone = f.Create();
                        Assert.AreEqual(events, f.Application.EnableEvents);
                        Assert.AreNotSame(f.Source.VBProject, clone.Project);
                        Assert.AreEqual(extension, Path.GetExtension(clone.Path));
                        var copy = f.Application.Workbooks[1];
                        clone.Dispose(); clone.Dispose();
                        Assert.AreEqual(1, copy.Closes);
                        Assert.AreEqual(0, f.Source.Closes);
                        Assert.AreEqual(events, f.Application.EnableEvents);
                        Assert.AreEqual(1, f.Application.Workbooks.Count);
                    }
            new VbaTestCoverageClone().Dispose();
        }

        [TestMethod]
        public void HostRoutingCallsOnlyTheMatchingProviderAndDefaultHostRefusesNativeMutation()
        {
            var boundary = new VbaTestCoverageClone.HostBoundary();
            var expected = new VbaTestCoverageClone();
            boundary.ReadProcessName = () => "WINWORD";
            boundary.CreateWord = (project, source, folder) => expected;
            boundary.CreatePowerPoint = (project, source, folder) => expected;
            Assert.AreSame(expected, VbaTestCoverageClone.CreateOwned(null, null, null, boundary));
            boundary.ReadProcessName = () => "POWERPNT";
            Assert.AreSame(expected, VbaTestCoverageClone.CreateOwned(null, null, null, boundary));
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestCoverageClone.CreateOwned(null, null, null));
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestCoverageClone.CreateOwned(null, null, null, null));
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestCoverageClone.CreateExcel(null, null, null));
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestWordCoverageClone.CreateWord(null, null, null));
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestPowerPointCoverageClone.CreatePowerPoint(null, null, null));
            Assert.AreEqual(0u, boundary.ReadWindowOwner(IntPtr.Zero));
        }

        [TestMethod]
        public void VerifiedCloseBalancesOwnedInterfacesAndDoesNotFinalReleaseTheBorrowedProject()
        {
            using (var fixture=new Fixture())
            {
                var releases=new Dictionary<object,int>();
                fixture.Boundary.ReleaseAcquired=value => { releases[value]=releases.TryGetValue(value,out int count) ? count+1 : 1; };
                var clone=fixture.Create(); var copy=fixture.Application.Workbooks[1];
                Assert.AreEqual(1,releases[fixture.Source.VBProject], "Only the source getter acquisition is released; the caller's project is borrowed.");
                Assert.IsFalse(releases.ContainsKey(copy.VBProject)); Assert.IsFalse(releases.ContainsKey(copy));
                clone.Dispose();
                Assert.AreEqual(1,releases[copy.VBProject]); Assert.AreEqual(1,releases[copy]);
                Assert.AreEqual(1,releases[fixture.Application]); Assert.AreEqual(1,releases[fixture.Application.Workbooks]);
                Assert.AreEqual(2,releases[fixture.Source], "Initial enumeration and post-close absence check each acquire one source interface.");
                Assert.AreEqual(1,copy.Closes); Assert.AreEqual(0,fixture.Source.Closes);
                int releaseEntries=releases.Count; clone.Dispose(); Assert.AreEqual(releaseEntries,releases.Count);
            }
        }

        [DataTestMethod,DataRow("Owner"),DataRow("Path"),DataRow("Cancel"),DataRow("Close")]
        public void UncertainCopyCloseRetainsInterfacesAndNeverClosesAChangedTargetTwice(string fault)
        {
            using (var fixture=new Fixture())
            {
                var releases=new List<object>(); fixture.Boundary.ReleaseAcquired=releases.Add;
                var clone=fixture.Create(); var copy=fixture.Application.Workbooks[1];
                if (fault=="Owner") fixture.Boundary.ReadWindowOwner=_ => 777;
                if (fault=="Path") copy.FullName=fixture.Source.FullName;
                if (fault=="Cancel") copy.CancelClose=true;
                if (fault=="Close") copy.OnClose=() => { throw new InvalidOperationException("unknown native close"); };
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                clone.Dispose();
                Assert.AreEqual(fault=="Owner" || fault=="Path" ? 0 : 1,copy.Closes);
                Assert.AreEqual(0,fixture.Source.Closes);
                Assert.IsFalse(releases.Contains(copy));
                Assert.AreEqual(fault=="Owner" || fault=="Path" ? 0 : 1,
                    releases.FindAll(value => ReferenceEquals(value, copy.VBProject)).Count,
                    "The project acquisition is released only after owner/path preflight, before the single close attempt.");
                Assert.IsFalse(releases.Contains(fixture.Application)); Assert.IsTrue(fixture.Application.EnableEvents);
            }
        }

        [TestMethod]
        public void CopyProjectAcquisitionIsReleasedBeforeCloseAndNeverReleasedTwice()
        {
            using (var fixture = new Fixture())
            {
                var releases = new List<object>(); fixture.Boundary.ReleaseAcquired = releases.Add;
                var clone = fixture.Create(); var copy = fixture.Application.Workbooks[1];
                copy.OnClose = () =>
                {
                    Assert.AreEqual(1, releases.FindAll(value => ReferenceEquals(value, copy.VBProject)).Count);
                    Assert.IsFalse(releases.Contains(copy), "The workbook remains owned throughout Close.");
                    Assert.IsFalse(releases.Contains(fixture.Application));
                };
                clone.Dispose(); clone.Dispose();
                Assert.AreEqual(1, copy.Closes);
                Assert.AreEqual(1, releases.FindAll(value => ReferenceEquals(value, copy.VBProject)).Count);
                Assert.AreEqual(0, fixture.Source.Closes);
            }
        }

        [TestMethod]
        public void UncertainProjectReleaseCannotDispatchCloseOrRetryItsAcquisition()
        {
            using (var fixture = new Fixture())
            {
                var releases = new List<object>(); object copyProject = null;
                fixture.Boundary.ReleaseAcquired = value =>
                {
                    releases.Add(value);
                    if (ReferenceEquals(value, copyProject)) throw new InvalidOperationException("Uncertain COM release.");
                };
                var clone = fixture.Create(); var copy = fixture.Application.Workbooks[1]; copyProject = copy.VBProject;
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, releases.FindAll(value => ReferenceEquals(value, copyProject)).Count);
                Assert.AreEqual(0, copy.Closes); Assert.AreEqual(0, fixture.Source.Closes);
                Assert.IsFalse(releases.Contains(copy)); Assert.IsFalse(releases.Contains(fixture.Application));
                Assert.IsTrue(fixture.Application.EnableEvents);
            }
        }

        [TestMethod]
        public void KnownPreflightFailureReleasesAcquisitionsWithoutSaveWhileUnknownOpenRetainsThem()
        {
            using (var fixture=new Fixture())
            {
                var releases=new List<object>(); fixture.Boundary.ReleaseAcquired=releases.Add;
                fixture.Boundary.ReadWindowOwner=_ => 777;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Create());
                CollectionAssert.AreEqual(new object[] {fixture.Application},releases);
                Assert.AreEqual(0,fixture.Source.Saves); Assert.AreEqual(0,fixture.Application.Workbooks.Opens);
            }
            using (var fixture=new Fixture())
            {
                var releases=new List<object>(); fixture.Boundary.ReleaseAcquired=releases.Add;
                fixture.Application.Workbooks.OnOpen=_ => { throw new IOException("unknown open completion"); };
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Create()).Uncertain);
                Assert.AreEqual(2,releases.Count);
                Assert.AreSame(fixture.Source.VBProject,releases[0]);
                Assert.IsInstanceOfType(releases[1],typeof(System.Collections.IEnumerator));
                Assert.IsFalse(releases.Contains(fixture.Source));
                Assert.IsFalse(releases.Contains(fixture.Application.Workbooks));
                Assert.IsFalse(releases.Contains(fixture.Application));
                Assert.AreEqual(1,fixture.Source.Saves);
                Assert.AreEqual(1,fixture.Application.Workbooks.Opens); Assert.AreEqual(0,fixture.Source.Closes);
            }
        }

        [TestMethod]
        public void InvalidOwnerProjectPathFormatAndOccupiedOutputNeverSaveOrOpen()
        {
            foreach (string fault in new[] { "pid", "missing", "duplicate", "path", "format", "occupied" })
                using (var f = new Fixture())
                {
                    string expected = f.Source.FullName;
                    if (fault == "pid") f.Boundary.ReadWindowOwner = _ => 777;
                    if (fault == "missing") f.Application.Workbooks.Clear();
                    if (fault == "duplicate") f.Application.Workbooks.Add(new Book { FullName = expected, VBProject = f.Source.VBProject });
                    if (fault == "path") f.Source.FullName = Path.Combine(f.Folder, "Renamed.xlsm");
                    if (fault == "format") expected = f.Source.FullName = Path.ChangeExtension(expected, ".xlsx");
                    if (fault == "occupied") { Directory.CreateDirectory(Path.Combine(f.Folder, "Copy")); File.WriteAllText(Path.Combine(f.Folder, "Copy", "coverage.xlsm"), "Existing"); }
                    Assert.ThrowsException<InvalidOperationException>(() => VbaTestCoverageClone.CreateExcel(f.Source.VBProject, expected, Path.Combine(f.Folder, "Copy"), f.Boundary), fault);
                    Assert.AreEqual(0, f.Source.Saves); Assert.AreEqual(0, f.Application.Workbooks.Opens);
                }
        }

        [TestMethod]
        public void PreparationFailureIsUncertainRestoresEventsAndNeverRetriesOrCloses()
        {
            foreach (string fault in new[] { "disable", "save", "open", "shared" })
                using (var f = new Fixture())
                {
                    if (fault == "disable") f.Application.IgnoreEventsWrite = true;
                    if (fault == "save") f.Source.OnSaveCopy = () => { throw new InvalidOperationException("Save outcome unknown"); };
                    if (fault == "open") f.Application.Workbooks.OnOpen = _ => { throw new InvalidOperationException("Open outcome unknown"); };
                    if (fault == "shared") f.Application.Workbooks.OnOpen = path => new Book { FullName = path, VBProject = f.Source.VBProject };
                    var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Create(), fault);
                    Assert.IsTrue(error.Uncertain);
                    Assert.IsTrue(f.Application.EnableEvents);
                    Assert.AreEqual(0, f.Source.Closes);
                    Assert.IsTrue(f.Source.Saves <= 1);
                }
        }

        [TestMethod]
        public void CloseFailureIsUncertainRestoresEventsAndNeverRetries()
        {
            foreach (string fault in new[] { "disable", "close", "cancel" })
                using (var f = new Fixture())
                {
                    var clone = f.Create(); var copy = f.Application.Workbooks[1];
                    if (fault == "disable") f.Application.IgnoreEventsWrite = true;
                    if (fault == "close") copy.OnClose = () => { throw new InvalidOperationException("Close outcome unknown"); };
                    if (fault == "cancel") copy.CancelClose = true;
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                    clone.Dispose();
                    Assert.IsTrue(copy.Closes <= 1);
                    Assert.IsTrue(f.Application.EnableEvents);
                    Assert.AreEqual(0, f.Source.Closes);
                }
        }

        [TestMethod]
        public void RestorationSetterOrReadbackFailureIsReportedAsUncertain()
        {
            foreach (bool setterThrows in new[] { false, true })
                using (var f = new Fixture())
                {
                    f.Source.OnSaveCopy = () =>
                    {
                        if (setterThrows) f.Application.OnEventsWrite = value => { if (value) throw new InvalidOperationException("Restore outcome unknown"); };
                        else f.Application.IgnoreEventsWrite = true;
                    };
                    var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Create());
                    Assert.IsTrue(error.Uncertain);
                    StringAssert.Contains(error.Message, "restoration could not be verified");
                    Assert.AreEqual(1, f.Source.Saves);
                    Assert.AreEqual(0, f.Source.Closes);
                }
        }
    }
}
