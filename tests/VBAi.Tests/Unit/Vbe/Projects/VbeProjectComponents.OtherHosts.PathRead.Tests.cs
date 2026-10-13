using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeOtherHostPersistenceTests
    {
        /// <summary>Rejects unrelated state queries while observing the exact native path getters.</summary>
        public sealed class PathOnlyWordDocument
        {
            public string Directory, FullPath;
            public object Project;
            public object VBProject => Project;
            public Action BeforeDirectory, BeforeFullPath;
            public int DirectoryReads, FullPathReads;
            public string Path { get { DirectoryReads++; BeforeDirectory?.Invoke(); return Directory; } }
            public string FullName { get { FullPathReads++; BeforeFullPath?.Invoke(); return FullPath; } }
            public bool ReadOnly => throw new AssertFailedException("Host-path reads must not query ReadOnly.");
            public bool Saved => throw new AssertFailedException("Host-path reads must not query Saved.");
            public int SaveFormat => throw new AssertFailedException("Host-path reads must not query SaveFormat.");
        }

        [DataTestMethod]
        [DataRow("saved"), DataRow("empty"), DataRow("whitespace"), DataRow("null path")]
        [DataRow("directory failure"), DataRow("full path failure"), DataRow("changed path")]
        public void WordPathReadTouchesOnlyFreshDirectoryAndFullName(string scenario)
        {
            var native = new VbeProjectComponents.NativeOtherHostProbe { ReadHostKind = () => "Word" };
            var document = new PathOnlyWordDocument { Directory = @"C:\Owned", FullPath = @"C:\Owned\Current.docm" };
            if (scenario == "empty") document.Directory = "";
            if (scenario == "whitespace") document.Directory = " ";
            if (scenario == "null path") document.Directory = null;
            var failure = new InvalidOperationException("path getter refused");
            if (scenario == "directory failure") document.BeforeDirectory = () => { throw failure; };
            if (scenario == "full path failure") document.BeforeFullPath = () => { throw failure; };
            if (scenario == "changed path") document.BeforeFullPath = () => document.FullPath = @"C:\Owned\Renamed.docm";
            string path = null;
            Action read = () => path = native.ReadWordDocumentPath(document);
            if (scenario.EndsWith("failure")) Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(read));
            else
            {
                read();
                bool unsaved = scenario == "empty" || scenario == "whitespace" || scenario == "null path";
                Assert.AreEqual(unsaved ? "" : document.FullPath, path);
            }
            Assert.AreEqual(1, document.DirectoryReads);
            Assert.AreEqual(scenario == "empty" || scenario == "whitespace" || scenario == "null path" || scenario == "directory failure" ? 0 : 1,
                document.FullPathReads);
        }

        [TestMethod]
        public void WordPathReadRejectsWrongHostAndNullDocumentBeforeReadingItsProperties()
        {
            var document = new PathOnlyWordDocument();
            var native = new VbeProjectComponents.NativeOtherHostProbe { ReadHostKind = () => "Excel" };
            Assert.ThrowsException<InvalidOperationException>(() => native.ReadWordDocumentPath(document));
            native.ReadHostKind = () => "Word";
            Assert.ThrowsException<ArgumentNullException>(() => native.ReadWordDocumentPath(null));
            Assert.AreEqual(0, document.DirectoryReads);
            Assert.AreEqual(0, document.FullPathReads);
        }

        /// <summary>Provides explicit one-based window/document catalogues without activating a host.</summary>
        public sealed class PathOnlyWordApplication
        {
            public OwnedReadSequence WindowItems, DocumentItems;
            public IEnumerable Windows => WindowItems;
            public IEnumerable Documents => DocumentItems;
        }

        [DataTestMethod]
        [DataRow("none"), DataRow("late owner"), DataRow("late project"), DataRow("path failure")]
        public void NativeWordPathSpecializationRetainsFinalProcessAndProjectChecks(string scenario)
        {
            var project = new object();
            var document = new PathOnlyWordDocument { Project = project, Directory = @"C:\Owned", FullPath = @"C:\Owned\Current.docm" };
            var window = new OwnedReadWindow { Handle = 1 };
            var application = new PathOnlyWordApplication
            {
                WindowItems = new OwnedReadSequence(new object[] { window }, "none"),
                DocumentItems = new OwnedReadSequence(new object[] { document }, "none")
            };
            var released = new List<object>(); bool changedOwner = false; int applications = 0;
            var native = new VbeProjectComponents.NativeOtherHostProbe
            {
                ReadHostKind = () => "Word", ReadIdentity = ReferenceEquals,
                ReadActiveApplication = name => { applications++; return application; }, ReleaseAcquiredReference = released.Add
            };
            native.ReadOwner = handle => (uint)native.CurrentProcessId + (changedOwner ? 1u : 0u);
            document.BeforeFullPath = () =>
            {
                if (scenario == "late owner") changedOwner = true;
                if (scenario == "late project") document.Project = new object();
                if (scenario == "path failure") throw new InvalidOperationException("path refused");
            };
            Action read = () => Assert.AreEqual(document.FullPath, VbeProjectHostPath.Read(project, native));
            if (scenario == "none") read();
            else Assert.ThrowsException<InvalidOperationException>(read);
            Assert.AreEqual(1, document.FullPathReads);
            Assert.AreEqual(scenario == "path failure" ? 1 : 2, applications);
            Assert.AreEqual(1, released.FindAll(value => ReferenceEquals(value, document)).Count);
            Assert.AreEqual(applications, released.FindAll(value => ReferenceEquals(value, application)).Count);
        }
    }
}
