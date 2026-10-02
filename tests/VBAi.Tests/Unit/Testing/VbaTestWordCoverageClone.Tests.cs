using System;
using System.IO;
using System.Collections.Generic;
using System.IO.MemoryMappedFiles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Fixture = VBAi.Tests.Unit.VbaTestWordValuesHostTests.Fixture;
using Document = VBAi.Tests.Unit.VbaTestWordValuesHostTests.Document;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbaTestWordCoverageCloneTests
    {
        [TestMethod]
        public void CopyLeasesBalanceEveryDocumentAcquisitionOnlyAfterVerifiedClose()
        {
            using (var fixture = new Fixture(ownsApplication: true))
            using (var releases = new WordReleaseRecorder())
            {
                var acquired = new Dictionary<object, int>();
                var read = fixture.Host.ReadDocumentItem;
                fixture.Host.ReadDocumentItem = (documents, index) => {
                    object value = read(documents, index);
                    Increment(acquired, value); return value;
                };
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => {
                    copy = new Document { FullName = path, Application = fixture.Application };
                    Increment(acquired, copy); // Documents.Open transfers its own acquisition.
                    return copy;
                };
                var clone = new VbaTestWordCoverageClone { Host = fixture.Host }.Create(
                    fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                Assert.AreEqual(acquired[fixture.Source] - 1, releases.Count(fixture.Source));
                Assert.AreEqual(acquired[copy] - 2, releases.Count(copy));
                Assert.AreEqual(fixture.Reads - 2, releases.Count(fixture.Application));
                clone.Dispose();
                foreach (var item in acquired) Assert.AreEqual(item.Value, releases.Count(item.Key));
                Assert.AreEqual(fixture.Reads, releases.Count(fixture.Application));
                int count = releases.Total;
                clone.Dispose();
                Assert.AreEqual(count, releases.Total);
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void CopyPreflightRefusalBalancesOwnedSourceWithoutReleasingItsBorrowedProject()
        {
            using (var fixture = new Fixture(".docx", true))
            using (var releases = new WordReleaseRecorder())
            {
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestWordCoverageClone { Host = fixture.Host }.Create(
                    fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy")));
                Assert.AreEqual(1, releases.Count(fixture.Source));
                Assert.AreEqual(fixture.Reads, releases.Count(fixture.Application));
                // Only FindDocument's VBProject getter is acquired; the caller's project is borrowed.
                Assert.AreEqual(1, releases.Count(fixture.Source.VBProject));
                Assert.AreEqual(0, fixture.Application.Documents.OpenCalls);
            }
        }

        [TestMethod]
        public void UnverifiedOpenAndCloseKeepLongLivedAcquisitionsWithoutRetry()
        {
            foreach (bool failOpen in new[] { false, true })
            using (var fixture = new Fixture(ownsApplication: true))
            using (var releases = new WordReleaseRecorder())
            {
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => {
                    if (failOpen) throw new InvalidOperationException("Native open uncertain");
                    return copy = new Document { FullName = path, Application = fixture.Application, CancelClose = true };
                };
                var acquired = new Dictionary<object, int>();
                var read = fixture.Host.ReadDocumentItem;
                fixture.Host.ReadDocumentItem = (documents, index) => { var value = read(documents, index); Increment(acquired, value); return value; };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                string folder = Path.Combine(fixture.Folder, "Copy");
                if (failOpen)
                {
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder)).Uncertain);
                    Assert.AreEqual(acquired[fixture.Source] - 1, releases.Count(fixture.Source));
                    Assert.AreEqual(fixture.Reads - 1, releases.Count(fixture.Application));
                }
                else
                {
                    var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder);
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                    int released = releases.Total;
                    clone.Dispose();
                    Assert.AreEqual(released, releases.Total);
                    Assert.AreEqual(1, copy.CloseCalls);
                    Assert.IsTrue(releases.Count(copy) < acquired[copy], "Open, target and failed absence-probe acquisitions must remain leased.");
                    Assert.IsTrue(releases.Count(fixture.Application) < fixture.Reads);
                }
                Assert.AreEqual(1, fixture.Application.Documents.OpenCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void VerifiedCloseReleaseFailurePreservesItsAcquisitionAndNeverRetriesCloseOrRelease()
        {
            using (var fixture = new Fixture(ownsApplication: true))
            using (var releases = new WordReleaseRecorder())
            {
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => copy = new Document { FullName = path, Application = fixture.Application };
                var clone = new VbaTestWordCoverageClone { Host = fixture.Host }.Create(
                    fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                releases.ThrowReference = copy.VBProject;
                releases.ThrowNth = releases.Count(copy.VBProject) + 2; // ValidateTarget getter, then clone-owned getter.
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => clone.Dispose()).Message, "Release failed");
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.IsFalse(fixture.Application.Documents.Contains(copy));
                int count = releases.Total;
                clone.Dispose(); Assert.AreEqual(count, releases.Total);
                Assert.AreEqual(fixture.Reads, releases.Count(fixture.Application));
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }
        [TestMethod]
        public void ProbeOrOpenCollectionReleaseFailureRetainsReferencesWithoutMutationRetry()
        {
            foreach (bool failOpenRelease in new[] { false, true })
            using (var fixture = new Fixture(ownsApplication: true))
            using (var releases = new WordReleaseRecorder())
            {
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => copy = new Document { FullName = path, Application = fixture.Application };
                releases.ThrowReference = fixture.Application.Documents;
                // Resolve source, initial path probe, two validations, then Documents.Open.
                releases.ThrowNth = failOpenRelease ? 5 : 2;
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                string folder = Path.Combine(fixture.Folder, "Copy");
                if (failOpenRelease)
                {
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder)).Uncertain);
                    Assert.AreEqual(0, releases.Count(copy));
                    Assert.IsTrue(fixture.Application.Documents.Contains(copy));
                    Assert.AreEqual(1, fixture.Application.Documents.OpenCalls);
                    Assert.AreEqual(0, copy.CloseCalls);
                }
                else
                {
                    Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                    Assert.AreEqual(0, fixture.Application.Documents.OpenCalls);
                    Assert.IsFalse(Directory.Exists(folder));
                    Assert.AreEqual(fixture.Reads, releases.Count(fixture.Application));
                }
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }
        [TestMethod]
        public void NegativeDocumentCountAfterCloseIsUncertainAndDoesNotRetry()
        {
            using (var fixture = new Fixture())
            {
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => copy = new Document { FullName = path, Application = fixture.Application };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host, CloseCopy = document => {
                    ((Document)document).Close(0); fixture.Application.Documents.CountOverride = -1;
                } };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                StringAssert.Contains(error.Message, "Unexpected Word document count");
                clone.Dispose();
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }
        private static void Increment(Dictionary<object, int> values, object value)
        { values[value] = values.TryGetValue(value, out int count) ? count + 1 : 1; }

        private sealed class WordReleaseRecorder : IDisposable
        {
            private readonly Func<object, bool> priorCheck = VbaTestWordValuesHost.IsComReference;
            private readonly Func<object, int> priorRelease = VbaTestWordValuesHost.ReleaseComReference;
            private readonly Dictionary<object, int> released = new Dictionary<object, int>();
            internal int Total, ThrowNth;
            internal object ThrowReference;
            internal WordReleaseRecorder()
            {
                VbaTestWordValuesHost.IsComReference = value => true;
                VbaTestWordValuesHost.ReleaseComReference = value => { Increment(released, value); Total++; if (ReferenceEquals(value, ThrowReference) && Count(value) == ThrowNth) throw new InvalidOperationException("Release failed"); return 0; };
            }
            internal int Count(object value) => released.TryGetValue(value, out int count) ? count : 0;
            public void Dispose() { VbaTestWordValuesHost.IsComReference = priorCheck; VbaTestWordValuesHost.ReleaseComReference = priorRelease; }
        }
        [TestMethod]
        public void DefaultFileCopyAndOpenPreserveAllFourSavedFormatsAndOriginalBytes()
        {
            foreach (string extension in new[] { ".docm", ".dotm", ".doc", ".dot" })
            using (var fixture = new Fixture(extension))
            {
                fixture.Application.Documents.OnOpen = path => new Document { FullName = path, Application = fixture.Application };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                string original = fixture.Source.FullName, bytes = File.ReadAllText(original);
                var clone = provider.Create(fixture.Source.VBProject, original, Path.Combine(fixture.Folder, "Copy"));
                Assert.AreEqual(extension, Path.GetExtension(clone.Path));
                Assert.AreEqual(bytes, File.ReadAllText(clone.Path));
                Assert.AreEqual(bytes, File.ReadAllText(original));
                Assert.AreEqual(original, fixture.Source.FullName);
                Assert.AreEqual(0, fixture.Source.SaveCalls + fixture.Source.SaveAsCalls);
                Assert.IsFalse(fixture.Application.Documents.LastRecent);
                Assert.IsFalse(fixture.Application.Documents.LastReadOnly);
                Assert.IsFalse(fixture.Application.Documents.LastVisible);
                Assert.IsFalse(fixture.Application.Documents.LastConversions);
                Assert.IsFalse(fixture.Application.Documents.LastRevert);
                Assert.IsFalse(fixture.Application.Documents.LastRepair);
                var copy = fixture.Application.Documents[1];
                clone.Dispose(); clone.Dispose();
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.AreEqual(0, copy.LastSaveChanges);
                Assert.AreEqual(1, fixture.Application.Documents.Count);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void WrongOwnerUnsupportedFormatAndOccupiedCopyPathRefuseWithoutOpen()
        {
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                fixture.Host.ReadWindowOwner = hwnd => 999;
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                fixture.Host.ReadWindowOwner = hwnd => 123;
                File.WriteAllText(Path.Combine(fixture.Folder, "coverage.docm"), "Owned existing file");
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                Assert.AreEqual("Owned existing file", File.ReadAllText(Path.Combine(fixture.Folder, "coverage.docm")));
                fixture.Source.FullName = Path.Combine(fixture.Folder, "coverage.docm");
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Unsupported.docx");
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Other")));
                Assert.AreEqual(0, fixture.Application.Documents.OpenCalls);
            }
        }

        [TestMethod]
        public void ReturnedOriginalSharedProjectOrWrongPathIsNeverClosed()
        {
            foreach (string fault in new[] { "null", "original", "project", "path" })
            using (var fixture = new Fixture())
            {
                Document returned = null;
                fixture.Application.Documents.OnOpen = path => returned = fault == "null" ? null : fault == "original" ? fixture.Source : new Document {
                    FullName = fault == "path" ? fixture.Source.FullName : path, Application = fixture.Application,
                    VBProject = fault == "project" ? fixture.Source.VBProject : new object() };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy")));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, fixture.Application.Documents.OpenCalls, fault);
                Assert.AreEqual(0, returned?.CloseCalls ?? 0, fault);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void PostOpenHostSourceAndCollectionFailuresAreUncertainAndRetainTheCopy()
        {
            foreach (string fault in new[] { "pid", "source", "collection" })
            using (var fixture = new Fixture())
            {
                Document copy = null;
                fixture.Application.Documents.OnOpen = path => copy = new Document { FullName = path, Application = fixture.Application };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host, OpenCopy = (application, path) => {
                    var returned = fixture.Application.Documents.Open(path, false, false, false, false, false, false);
                    if (fault == "pid") fixture.Host.ReadWindowOwner = hwnd => 999;
                    if (fault == "source") fixture.Source.FullName = Path.Combine(fixture.Folder, "Changed.docm");
                    if (fault == "collection") fixture.Application.Documents.Remove(returned);
                    return returned;
                } };
                string folder = Path.Combine(fixture.Folder, "Copy");
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, fixture.Application.Documents.OpenCalls, fault);
                Assert.AreEqual(0, copy.CloseCalls, fault);
                Assert.AreEqual(0, fixture.Source.CloseCalls, fault);
                Assert.IsTrue(File.Exists(Path.Combine(folder, "coverage.docm")), fault);
                StringAssert.Contains(error.Message, "Retained copy:");
            }
        }

        [TestMethod]
        public void UncertainOpenRetainsCopiedFileAndNeverRetriesOrClosesOriginal()
        {
            using (var fixture = new Fixture())
            {
                fixture.Application.Documents.OnOpen = path => { throw new InvalidOperationException("Native open result unknown"); };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                string folder = Path.Combine(fixture.Folder, "Copy");
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                Assert.IsTrue(error.Uncertain);
                Assert.IsTrue(File.Exists(Path.Combine(folder, "coverage.docm")));
                Assert.AreEqual(1, fixture.Application.Documents.OpenCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void CancelledCloseIsUncertainAndNotRetriedOrMisreportedAsClosed()
        {
            using (var fixture = new Fixture())
            {
                fixture.Application.Documents.OnOpen = path => new Document { FullName = path, Application = fixture.Application, CancelClose = true };
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var copy = fixture.Application.Documents[1];
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.AreEqual(2, fixture.Application.Documents.Count);
                Assert.IsTrue(File.Exists(clone.Path));
            }
        }

        [TestMethod]
        public void SavedFileCopyAcceptsAnIdleWordWriterAndBlocksWritesIncludingAppendUntilVerificationEnds()
        {
            using (var f = new Fixture())
            using (var writer = new FileStream(f.Source.FullName, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete, 1))
            {
                byte[] original = System.Text.Encoding.UTF8.GetBytes("Saved Word fixture bytes");
                var provider = new VbaTestWordCoverageClone { Host = f.Host };
                provider.CopySavedFile = (source, destination) => VbaTestWordCoverageClone.CopyFile(source, destination, (input, output) =>
                {
                    writer.Position = 0;
                    Assert.ThrowsException<IOException>(() => writer.Write(new byte[] { 9, 9, 9, 9 }, 0, 4));
                    writer.Position = writer.Length;
                    Assert.ThrowsException<IOException>(() => writer.Write(new byte[] { 9, 9, 9, 9 }, 0, 4));
                    Assert.ThrowsException<IOException>(() => File.Delete(source));
                });
                f.Application.Documents.OnOpen = path => new Document { FullName = path, Application = f.Application };
                var clone = provider.Create(f.Source.VBProject, f.Source.FullName, Path.Combine(f.Folder, "Copy"));
                CollectionAssert.AreEqual(original, File.ReadAllBytes(clone.Path));
                writer.Position = 0;
                writer.Write(original, 0, original.Length); writer.Flush();
                Assert.AreEqual(1, f.Application.Documents.OpenCalls);
                clone.Dispose();
                Assert.AreEqual(0, f.Source.SaveCalls + f.Source.SaveAsCalls + f.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void ConflictingSourceRangeLockRefusesWithoutOpeningOrRetryingTheCopy()
        {
            using (var f = new Fixture())
            using (var writer = new FileStream(f.Source.FullName, FileMode.Open, FileAccess.Write, FileShare.ReadWrite))
            {
                writer.Lock(0, 1);
                try
                {
                    var provider = new VbaTestWordCoverageClone { Host = f.Host };
                    var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, Path.Combine(f.Folder, "Copy")));
                    Assert.IsTrue(error.Uncertain);
                    Assert.AreEqual(0, f.Application.Documents.OpenCalls + f.Source.CloseCalls);
                }
                finally { writer.Unlock(0, 1); }
            }
        }

        [DataTestMethod]
        [DataRow("source-length")]
        [DataRow("output-length")]
        [DataRow("metadata")]
        [DataRow("mapped-source")]
        [DataRow("output-bytes")]
        [DataRow("callback")]
        public void CopyVerificationFailureRetainsDestinationWithoutOpeningOrRetryingWord(string fault)
        {
            using (var f = new Fixture())
            {
                int copies = 0;
                DateTime originalModified = File.GetLastWriteTimeUtc(f.Source.FullName);
                var provider = new VbaTestWordCoverageClone { Host = f.Host };
                provider.CopySavedFile = (source, destination) =>
                {
                    copies++;
                    VbaTestWordCoverageClone.CopyFile(source, destination, (input, output) =>
                    {
                        if (fault == "source-length")
                            using (var writer = new FileStream(source, FileMode.Open, FileAccess.Write, FileShare.ReadWrite)) writer.SetLength(0);
                        if (fault == "output-length") output.SetLength(0);
                        if (fault == "metadata") File.SetLastWriteTimeUtc(source, originalModified.AddMinutes(1));
                        if (fault == "mapped-source")
                        {
                            using (var stream = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                            using (var map = MemoryMappedFile.CreateFromFile(stream, null, 0, MemoryMappedFileAccess.ReadWrite, HandleInheritability.None, false))
                            using (var view = map.CreateViewAccessor()) { view.Write(0, (byte)9); view.Flush(); }
                            File.SetLastWriteTimeUtc(source, originalModified);
                        }
                        if (fault == "output-bytes") { output.Position = 0; output.WriteByte(9); output.Flush(true); }
                        if (fault == "callback") throw new IOException("Synthetic verification failure.");
                    });
                };
                string folder = Path.Combine(f.Folder, "Copy");
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, folder));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, copies, fault);
                StringAssert.Contains(error.Message, "Retained copy:");
                Assert.IsTrue(File.Exists(Path.Combine(folder, "coverage.docm")), fault);
                Assert.AreEqual(0, f.Application.Documents.OpenCalls + f.Source.CloseCalls, fault);
                using (var probe = new FileStream(f.Source.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                { probe.Lock(0, long.MaxValue); probe.Unlock(0, long.MaxValue); }
            }
        }

        [TestMethod]
        public void ExclusiveDestinationCreationPreservesAnExistingFileAndReleasesTheSourceLock()
        {
            using (var f = new Fixture())
            {
                string destination = Path.Combine(f.Folder, "AlreadyExists.docm");
                File.WriteAllText(destination, "Preserve existing bytes");
                Assert.ThrowsException<IOException>(() => VbaTestWordCoverageClone.CopyFile(f.Source.FullName, destination, null));
                Assert.AreEqual("Preserve existing bytes", File.ReadAllText(destination));
                using (var probe = new FileStream(f.Source.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                { probe.Lock(0, long.MaxValue); probe.Unlock(0, long.MaxValue); }
            }
        }
        [TestMethod]
        public void UnsavedOriginalContentRefusesCoverageBeforeCreatingAnyCopyOrOpeningWord()
        {
            using (var fixture = new Fixture())
            {
                fixture.Source.Saved = false;
                string folder = Path.Combine(fixture.Folder, "Copy");
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                Assert.IsFalse(Directory.Exists(folder));
                Assert.AreEqual(0, fixture.Application.Documents.OpenCalls);
                Assert.AreEqual(0, fixture.Source.SaveCalls + fixture.Source.SaveAsCalls + fixture.Source.CloseCalls);
                Assert.IsNotNull(fixture.Resolve(), "Ordinary run target resolution must still allow unsaved content.");
            }
        }

        [TestMethod]
        public void UnsavedContentAfterCopyOrOpenIsUncertainAndNeverDispatchesOrCloses()
        {
            foreach (string phase in new[] { "copy", "open" })
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                Document copy = null;
                if (phase == "copy") provider.CopySavedFile = (source, destination) => {
                    File.Copy(source, destination); fixture.Source.Saved = false;
                };
                fixture.Application.Documents.OnOpen = path => {
                    fixture.Source.Saved = false;
                    return copy = new Document { FullName = path, Application = fixture.Application };
                };
                string folder = Path.Combine(fixture.Folder, "Copy");
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                Assert.IsTrue(error.Uncertain, phase);
                Assert.AreEqual(phase == "copy" ? 0 : 1, fixture.Application.Documents.OpenCalls);
                Assert.IsTrue(File.Exists(Path.Combine(folder, "coverage.docm")));
                Assert.AreEqual(0, copy?.CloseCalls ?? 0);
                Assert.AreEqual(0, fixture.Source.SaveCalls + fixture.Source.SaveAsCalls + fixture.Source.CloseCalls);
            }
        }
        [TestMethod]
        public void ExistingDirectoryAndOpenCopyPathAreRejectedBeforeCopying()
        {
            using (var f = new Fixture())
            {
                var provider = new VbaTestWordCoverageClone { Host = f.Host };
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, f.Folder));
                var folder = Path.Combine(f.Folder, "New");
                f.Application.Documents.Add(new Document { FullName = Path.Combine(folder, "coverage.docm"), Application = f.Application });
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, folder));
                Assert.IsFalse(Directory.Exists(folder)); Assert.AreEqual(0, f.Application.Documents.OpenCalls);
                f.Source.Saved = "not a Boolean";
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, folder));
            }
        }

        [TestMethod]
        public void InconsistentCopyIdentityCannotBeAcceptedOrCloseTheOriginal()
        {
            foreach (string fault in new[] { "verify", "close" })
            using (var f = new Fixture())
            {
                Document copy = null;
                f.Application.Documents.OnOpen = path => copy = new Document { FullName = path, Application = f.Application };
                var provider = new VbaTestWordCoverageClone { Host = f.Host };
                if (fault == "verify")
                {
                    f.Host.SameIdentity = (a, b) => !(ReferenceEquals(a, copy) && ReferenceEquals(b, copy) && copy != null) && ReferenceEquals(a, b);
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(f.Source.VBProject, f.Source.FullName, Path.Combine(f.Folder, "Copy"))).Uncertain);
                }
                else
                {
                    var clone = provider.Create(f.Source.VBProject, f.Source.FullName, Path.Combine(f.Folder, "Copy"));
                    f.Host.SameIdentity = (a, b) => ReferenceEquals(a, b) || (ReferenceEquals(a, copy) && ReferenceEquals(b, f.Source));
                    Assert.ThrowsException<InvalidOperationException>(() => clone.Dispose());
                }
                Assert.AreEqual(0, f.Source.CloseCalls); Assert.AreEqual(0, copy.CloseCalls);
            }
        }

        [TestMethod]
        public void CloseRechecksApplicationAndBoundsRemainingDocuments()
        {
            foreach (bool replaced in new[] { false, true })
            using (var f = new Fixture())
            {
                f.Application.Documents.OnOpen = path => new Document { FullName = path, Application = f.Application };
                var provider = new VbaTestWordCoverageClone { Host = f.Host, CloseCopy = copy => {
                    f.Application.Documents.Remove((Document)copy);
                    if (replaced) f.Host.ReadActiveApplication = _ => new VbaTestWordValuesHostTests.Application();
                    else for (int i = 0; i < 1000; i++) f.Application.Documents.Add(new Document());
                } };
                var clone = provider.Create(f.Source.VBProject, f.Source.FullName, Path.Combine(f.Folder, "Copy"));
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose()).Uncertain);
                Assert.AreEqual(0, f.Source.CloseCalls);
            }
        }

    }
}
