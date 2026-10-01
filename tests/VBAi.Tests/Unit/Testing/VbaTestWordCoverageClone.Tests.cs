using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Fixture = VBAi.Tests.Unit.VbaTestWordValuesHostTests.Fixture;
using Document = VBAi.Tests.Unit.VbaTestWordValuesHostTests.Document;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestWordCoverageCloneTests
    {
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
        public void FileCopyRefusesWritersAndNeverOverwritesOrOpensPartialCopy()
        {
            using (var fixture = new Fixture())
            using (var writer = new FileStream(fixture.Source.FullName, FileMode.Open, FileAccess.Write, FileShare.ReadWrite))
            {
                var provider = new VbaTestWordCoverageClone { Host = fixture.Host };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy")));
                Assert.IsTrue(error.Uncertain);
                Assert.AreEqual(0, fixture.Application.Documents.OpenCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
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
    }
}
