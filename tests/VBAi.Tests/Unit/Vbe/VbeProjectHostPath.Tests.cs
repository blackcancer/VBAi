using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeProjectHostPathTests
    {
        public sealed class Project
        {
            public string Name => "Project";
            public bool ThrowFileName;
            public string FileNameValue = @"C:\Temp\~WRL0001.tmp";
            public string FileName => ThrowFileName ? throw new DirectoryNotFoundException() : FileNameValue;
        }
        public sealed class Vbe { public List<object> VBProjects { get; } = new List<object>(); }
        private sealed class Document { internal object Project; internal string Path; }
        private sealed class Probe : VbeProjectComponents.IOtherHostProbe
        {
            internal readonly List<object> Items = new List<object>();
            internal uint Owner = 42;
            internal string Kind = "Word";
            internal Action OnState;
            internal bool Exists = true;
            public string HostKind => Kind;
            public int CurrentProcessId => 42;
            public object Application() => this;
            public uint ApplicationProcessId(object application) => Owner;
            public IList<object> Documents(object application) => Items;
            public object DocumentProject(object document) => ((Document)document).Project;
            public bool SameProject(object a, object b) => ReferenceEquals(a, b);
            public VbeProjectComponents.OtherHostDocumentState State(object document)
            {
                OnState?.Invoke();
                return new VbeProjectComponents.OtherHostDocumentState { Path = ((Document)document).Path, Saved = false, Format = 13 };
            }
            public bool FileExists(string path) => Exists;
            public bool DirectoryExists(string path) => true;
            public long FileLength(string path) => 1;
            public void Save(object document, bool saveAs, string destination, int format) => throw new AssertFailedException("Read-only path resolution must not save.");
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void WordDuplicateNamesResolveByMatchedDocumentPathsDespiteBackingStorage(bool missingBackingPath)
        {
            var first = new Project { ThrowFileName = missingBackingPath };
            var second = new Project { ThrowFileName = missingBackingPath };
            var probe = new Probe();
            probe.Items.Add(new Document { Project = first, Path = @"C:\Owned\First.docm" });
            probe.Items.Add(new Document { Project = second, Path = @"C:\Owned\Second.docm" });
            var vbe = new Vbe(); vbe.VBProjects.Add(first); vbe.VBProjects.Add(second);
            Func<object, string> path = project => VbeProjectHostPath.Read(project, probe);
            Assert.AreEqual(@"C:\Owned\First.docm", path(first));
            Assert.AreSame(second, (object)VbeProjectResolver.Resolve(vbe, @"C:\OWNED\SECOND.DOCM", path));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(vbe, "Project", path));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(vbe, @"C:\Temp\~WRL0001.tmp", path));
        }

        [TestMethod]
        public void WordIdentityPidAndPathFailuresNeverUseBackingStorage()
        {
            var project = new Project(); var probe = new Probe();
            var document = new Document { Project = project, Path = @"C:\Owned\First.docm" };
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            probe.Items.Add(document); probe.Items.Add(document);
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            probe.Items.RemoveAt(1); probe.Owner = 99;
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            probe.Owner = 42; probe.OnState = () => probe.Owner = 99;
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            probe.Owner = 42; probe.OnState = () => document.Project = new Project();
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            probe.OnState = null; document.Project = project; document.Path = "relative.docm";
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.Read(project, probe));
            document.Path = ""; Assert.IsNull(VbeProjectHostPath.Read(project, probe));
            probe.Kind = "PowerPoint";
            string actualPath = VbeProjectHostPath.Read(project, probe);
            Assert.IsTrue(string.Equals(project.FileName, actualPath, StringComparison.OrdinalIgnoreCase),
                "The canonical Windows path must identify the same backing file regardless of directory casing. Actual: " + actualPath);
        }

        [TestMethod]
        public void ExplicitMissingHostPathCannotFallBackToWordTemporaryFile()
        {
            var fields = new Dictionary<string, object> { { "FileName", @"C:\Temp\~WRL0001.tmp" }, { "HostPath", null } };
            Assert.AreEqual("", VbeProjectHostPath.FromFields(fields));
            fields["HostPath"] = @"C:\Owned\First.docm";
            Assert.AreEqual(@"C:\Owned\First.docm", VbeProjectHostPath.FromFields(fields));
            fields.Remove("HostPath"); Assert.AreEqual((string)fields["FileName"], VbeProjectHostPath.FromFields(fields));
        }

        [DataTestMethod]
        [DataRow("Projet 1", false)]
        [DataRow(@"C:\Windows\system32\Projet 1", false)]
        [DataRow(@"H:\Documents\Projet 1", false)]
        [DataRow(@"C:\Owned\Project1", true)]
        [DataRow("relative.otm", true)]
        [DataRow(@"C:\Owned\VbaProject.OTM", false)]
        public void OutlookOpaqueOrUnpersistedStorageKeepsTemporaryProjectIdentity(string reported, bool exists)
        {
            var project = new Project { FileNameValue = reported };
            var probe = new Probe { Kind = "Outlook", Exists = exists };
            Assert.IsNull(VbeProjectHostPath.Read(project, probe));
            var fields = new Dictionary<string, object> { { "FileName", reported },
                { "HostPath", VbeProjectHostPath.Read(project, probe) } };
            Assert.IsTrue(string.IsNullOrEmpty(VbeProjectHostPath.FromFields(fields)),
                "A missing canonical path must not promote the opaque Outlook alias to a saved session.");
        }

        [DataTestMethod]
        [DataRow(@"C:\Owned\VbaProject.OTM")]
        [DataRow(@"C:\Owned\VbaProject.otm")]
        public void OutlookExistingAbsoluteOtmRetainsItsCanonicalPath(string reported)
        {
            var project = new Project { FileNameValue = reported };
            var probe = new Probe { Kind = "Outlook", Exists = true };
            Assert.AreEqual(Path.GetFullPath(reported), VbeProjectHostPath.Read(project, probe));
        }

        [DataTestMethod]
        [DataRow(null, "OUTLOOK", true)]
        [DataRow(null, "outlook", true)]
        [DataRow("Outlook", "testhost", true)]
        [DataRow(null, "EXCEL", false)]
        [DataRow(null, "SLDWORKS", false)]
        [DataRow(null, null, false)]
        [DataRow("Word", "OUTLOOK", false)]
        [DataRow("", "OUTLOOK", false)]
        public void OutlookProcessRecognitionDoesNotExpandTheDocumentAdapterCatalogue(string kind, string process, bool expected)
            => Assert.AreEqual(expected, VbeProjectHostPath.IsOutlookPathHost(kind, process));
    }
}
