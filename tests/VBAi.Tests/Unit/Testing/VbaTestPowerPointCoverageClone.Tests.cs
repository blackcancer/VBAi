using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestPowerPointCoverageCloneTests
    {
        [TestMethod]
        public void DefaultNativeApiShapeCreatesExplicitPptmWithoutSavingOrRenamingOriginal()
        {
            using (var fixture = new Fixture())
            {
                string original = fixture.Source.FullName;
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                var clone = provider.Create(fixture.Source.VBProject, original, Path.Combine(fixture.Folder, "Run"));
                Assert.AreEqual(1, fixture.Source.SaveCopyCalls);
                Assert.AreEqual(25, fixture.Source.LastFormat);
                Assert.AreEqual(-2, fixture.Source.EmbedFonts);
                Assert.AreEqual(".pptm", Path.GetExtension(clone.Path));
                Assert.AreEqual(original, fixture.Source.FullName);
                Assert.AreEqual(0, fixture.Source.SaveCalls + fixture.Source.SaveAsCalls);
                Assert.AreNotSame(fixture.Source.VBProject, clone.Project);
                Assert.AreEqual(0, fixture.Application.Presentations.ReadOnly);
                Assert.AreEqual(0, fixture.Application.Presentations.Untitled);
                Assert.AreEqual(0, fixture.Application.Presentations.WithWindow);
                var copy = fixture.Application.Presentations[1];
                clone.Dispose(); clone.Dispose();
                Assert.AreEqual(1, copy.CloseCalls);
                Assert.AreEqual(-1, copy.Saved);
                Assert.AreEqual(1, copy.SavedWrites);
                Assert.AreEqual(-1, copy.SavedAtClose);
                Assert.AreEqual(0, copy.SaveCalls + copy.SaveAsCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
                Assert.AreEqual(0, fixture.Source.SavedWrites);
                Assert.AreEqual(0, fixture.Source.Saved);
            }
        }

        [TestMethod]
        public void WrongOwnerUnsupportedAddInAndStalePathNeverSaveCopy()
        {
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                fixture.Host.ReadWindowOwner = hwnd => 777;
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                fixture.Host.ReadWindowOwner = hwnd => 123;
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Original.ppam");
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                Assert.ThrowsException<InvalidOperationException>(() => provider.Create(fixture.Source.VBProject, Path.Combine(fixture.Folder, "Stale.pptm"), fixture.Folder));
                Assert.AreEqual(0, fixture.Source.SaveCopyCalls);
                Assert.AreEqual(0, fixture.Application.Presentations.OpenCalls);
            }
        }

        [TestMethod]
        public void SaveOrOpenExceptionsRemainUncertainWithoutRetryOrAutomaticClose()
        {
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                fixture.Source.OnSaveCopy = () => { throw new InvalidOperationException("SaveCopy unavailable"); };
                var saveError = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, fixture.Folder));
                Assert.IsTrue(saveError.Uncertain);
                Assert.AreEqual(1, fixture.Source.SaveCopyCalls);
                Assert.AreEqual(0, fixture.Application.Presentations.OpenCalls);
                fixture.Source.OnSaveCopy = null;
                fixture.Application.Presentations.OnOpen = path => { throw new InvalidOperationException("Open completion unavailable"); };
                var openError = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Second")));
                Assert.IsTrue(openError.Uncertain);
                Assert.AreEqual(1, fixture.Application.Presentations.OpenCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void UnverifiedCopyIdentityOrPathNeverClosesOriginalOrUnownedPresentation()
        {
            foreach (string fault in new[] { "null", "original", "project", "path", "projectRead", "resolve", "source" })
            using (var fixture = new Fixture())
            {
                Presentation returned = null;
                fixture.Application.Presentations.OnOpen = path => {
                    returned = fault == "null" ? null : fault == "original" ? fixture.Source :
                        new Presentation { FullName = fault == "path" ? Path.Combine(fixture.Folder, "Other.pptm") : path,
                            VBProject = fault == "project" ? fixture.Source.VBProject : new object() };
                    if (fault == "projectRead") returned.OnProjectRead = () => { throw new InvalidOperationException("Project getter failed after Open"); };
                    if (fault == "resolve") fixture.Host.ReadActiveApplication = progId => { throw new InvalidOperationException("Application ownership unavailable after Open"); };
                    if (fault == "source") fixture.Source.FullName = Path.Combine(fixture.Folder, "Renamed.pptm");
                    return returned;
                };
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                string folder = Path.Combine(fixture.Folder, "Copy");
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => provider.Create(fixture.Source.VBProject, fixture.Source.FullName, folder));
                Assert.IsTrue(error.Uncertain, fault);
                StringAssert.Contains(error.Message, Path.Combine(folder, "coverage.pptm"));
                Assert.AreEqual(1, fixture.Source.SaveCopyCalls, fault);
                Assert.AreEqual(1, fixture.Application.Presentations.OpenCalls, fault);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
                if (returned != null) { Assert.AreEqual(0, returned.CloseCalls); Assert.AreEqual(0, returned.SavedWrites); }
                Assert.AreEqual(0, fixture.Source.SavedWrites);
            }
        }

        [TestMethod]
        public void ChangedCopyPathBeforeDisposeIsRefusedInsteadOfClosingAnotherPresentation()
        {
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var copy = fixture.Application.Presentations[1];
                copy.FullName = fixture.Source.FullName;
                Assert.ThrowsException<InvalidOperationException>(() => clone.Dispose());
                Assert.AreEqual(0, copy.CloseCalls);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
                Assert.AreEqual(0, copy.SavedWrites);
                Assert.AreEqual(0, fixture.Source.SavedWrites);
            }
        }

        [TestMethod]
        public void CloseReturningWithPresentationStillOpenIsNotReportedAsVerifiedCompletion()
        {
            using (var fixture = new Fixture())
            {
                int attempts = 0;
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host, CloseCopy = presentation => attempts++ };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, attempts);
                Assert.AreEqual(2, fixture.Application.Presentations.Count);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
                Assert.AreEqual(0, fixture.Source.SavedWrites);
                Assert.AreEqual(-1, fixture.Application.Presentations[1].Saved);
            }
        }

        [TestMethod]
        public void FailedNativeCloseIsUncertainAndNeverAutomaticallyRetried()
        {
            using (var fixture = new Fixture())
            {
                int attempts = 0;
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host,
                    CloseCopy = presentation => { attempts++; throw new InvalidOperationException("Close completion unknown"); } };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, attempts);
                Assert.AreEqual(0, fixture.Source.CloseCalls);
            }
        }

        [TestMethod]
        public void FailedDiscardSetterOrReadbackIsUncertainWithoutClosingOrRetrying()
        {
            foreach (bool throws in new[] { false, true })
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var copy = fixture.Application.Presentations[1];
                copy.IgnoreSavedWrite = true;
                if (throws) copy.OnSavedWrite = () => { throw new InvalidOperationException("Saved completion unknown"); };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                StringAssert.Contains(error.Message, clone.Path);
                clone.Dispose();
                Assert.AreEqual(1, copy.SavedWrites);
                Assert.AreEqual(0, copy.CloseCalls + copy.SaveCalls + copy.SaveAsCalls);
                Assert.AreEqual(2, fixture.Application.Presentations.Count);
                Assert.AreEqual(0, fixture.Source.SavedWrites + fixture.Source.CloseCalls + fixture.Source.SaveCalls + fixture.Source.SaveAsCalls);
                Assert.AreEqual(0, fixture.Source.Saved);
            }
        }

        [TestMethod]
        public void IdentityChangedByDiscardSetterIsRecheckedBeforeClose()
        {
            using (var fixture = new Fixture())
            {
                var provider = new VbaTestPowerPointCoverageClone { Host = fixture.Host };
                var clone = provider.Create(fixture.Source.VBProject, fixture.Source.FullName, Path.Combine(fixture.Folder, "Copy"));
                var copy = fixture.Application.Presentations[1];
                copy.OnSavedWrite = () => copy.VBProject = fixture.Source.VBProject;
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => clone.Dispose());
                Assert.IsTrue(error.Uncertain);
                clone.Dispose();
                Assert.AreEqual(1, copy.SavedWrites);
                Assert.AreEqual(0, copy.CloseCalls);
                Assert.AreEqual(0, fixture.Source.SavedWrites + fixture.Source.CloseCalls);
            }
        }

        private sealed class Fixture : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-PowerPoint-Coverage-" + Guid.NewGuid().ToString("N"));
            internal readonly Application Application = new Application();
            internal readonly Presentation Source;
            internal readonly VbaTestPowerPointValuesHost Host = new VbaTestPowerPointValuesHost();
            internal Fixture()
            {
                Source = new Presentation { FullName = Path.Combine(Folder, "Original.pptm") };
                Application.Presentations.Add(Source);
                Host.ReadProcessName = () => "POWERPNT";
                Host.ReadProcessId = () => 123;
                Host.ReadApplicationWindow = application => { Assert.AreSame(Application, application); return new IntPtr(99); };
                Host.ReadWindowOwner = hwnd => { Assert.AreEqual(new IntPtr(99), hwnd); return 123; };
                Host.ReadActiveApplication = progId => { Assert.AreEqual("PowerPoint.Application", progId); return Application; };
            }
            public void Dispose() { if (Directory.Exists(Folder)) Directory.Delete(Folder, true); }
        }

        public sealed class Application
        {
            public Presentations Presentations { get; } = new Presentations();
        }

        public sealed class Presentations : List<Presentation>
        {
            public int OpenCalls, ReadOnly, Untitled, WithWindow;
            public Func<string, Presentation> OnOpen;
            public Presentation Open(string path, int readOnly, int untitled, int withWindow)
            {
                OpenCalls++; ReadOnly = readOnly; Untitled = untitled; WithWindow = withWindow;
                var copy = OnOpen == null ? new Presentation { FullName = path } : OnOpen(path);
                if (copy != null && !Contains(copy)) Add(copy);
                if (copy != null) copy.OnClose = () => Remove(copy);
                return copy;
            }
        }

        public sealed class Presentation
        {
            private object project = new object();
            public Action OnProjectRead;
            public object VBProject { get { OnProjectRead?.Invoke(); return project; } set { project = value; } }
            public string FullName { get; set; }
            public string Name => System.IO.Path.GetFileName(FullName);
            public string Path => System.IO.Path.GetDirectoryName(FullName);
            public int SaveCalls, SaveAsCalls, SaveCopyCalls, CloseCalls, LastFormat, EmbedFonts, SavedWrites, SavedAtClose;
            public bool IgnoreSavedWrite;
            public Action OnSaveCopy, OnClose, OnSavedWrite;
            private int saved;
            public int Saved
            {
                get => saved;
                set { SavedWrites++; OnSavedWrite?.Invoke(); if (!IgnoreSavedWrite) saved = value; }
            }
            public void SaveCopyAs(string path, int format, int embedFonts)
            { SaveCopyCalls++; LastFormat = format; EmbedFonts = embedFonts; OnSaveCopy?.Invoke(); }
            public void Close() { CloseCalls++; SavedAtClose = Saved; OnClose?.Invoke(); }
        }
    }
}
