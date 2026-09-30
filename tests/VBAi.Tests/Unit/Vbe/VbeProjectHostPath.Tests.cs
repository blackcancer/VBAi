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
            public string FileName => ThrowFileName ? throw new DirectoryNotFoundException() : @"C:\Temp\~WRL0001.tmp";
        }
        public sealed class Vbe { public List<object> VBProjects { get; } = new List<object>(); }
        private sealed class Document { internal object Project; internal string Path; }
        private sealed class Probe : VbeProjectComponents.IOtherHostProbe
        {
            internal readonly List<object> Items = new List<object>();
            internal uint Owner = 42;
            internal string Kind = "Word";
            internal Action OnState;
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
            public bool FileExists(string path) => true;
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
            probe.Kind = "PowerPoint"; Assert.AreEqual(project.FileName, VbeProjectHostPath.Read(project, probe));
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
    }
}
