using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks getter ownership and fresh Word identity validation without an Office process.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeProjectHostPathComOwnershipTests
    {
        private sealed class Document
        {
            internal object Project;
            internal string Path = @"C:\Owned\Disposable.docm";
        }

        private sealed class Probe : VbeProjectComponents.IOtherHostProbe
        {
            internal readonly object ApplicationValue = new object();
            internal readonly List<object> Acquired = new List<object>();
            internal readonly List<object> Items = new List<object>();
            internal string Fault, Kind = "Word";
            internal uint Owner = 42;
            internal int ApplicationCalls, ProjectCalls, StateCalls;
            public string HostKind => Kind;
            public int CurrentProcessId => 42;
            public object Application()
            {
                ApplicationCalls++;
                if (Fault == "application" || (Fault == "final application" && ApplicationCalls == 2))
                    throw new InvalidOperationException("application refused");
                Acquired.Add(ApplicationValue); return ApplicationValue;
            }
            public uint ApplicationProcessId(object application) => Owner;
            public IList<object> Documents(object application)
            {
                if (Fault == "documents") throw new InvalidOperationException("documents refused");
                if (Fault == "null documents") return null;
                Acquired.AddRange(Items); return Items;
            }
            public object DocumentProject(object document)
            {
                ProjectCalls++;
                if (Fault == "project" || (Fault == "final project" && ProjectCalls == 2))
                    throw new InvalidOperationException("project refused");
                object result = ((Document)document).Project;
                Acquired.Add(result); return result;
            }
            public bool SameProject(object a, object b) => ReferenceEquals(a, b);
            public VbeProjectComponents.OtherHostDocumentState State(object document)
            {
                StateCalls++;
                if (Fault == "state") throw new InvalidOperationException("state refused");
                if (Fault == "late owner") Owner = 99;
                if (Fault == "late project") ((Document)document).Project = new object();
                return new VbeProjectComponents.OtherHostDocumentState { Path = ((Document)document).Path };
            }
            public bool FileExists(string path) => true;
            public bool DirectoryExists(string path) => true;
            public long FileLength(string path) => 1;
            public void Save(object document, bool saveAs, string destination, int format)
                => throw new AssertFailedException("Read-only association must never save.");
        }

        [DataTestMethod]
        [DataRow("none"), DataRow("foreign document"), DataRow("empty path")]
        [DataRow("application"), DataRow("final application"), DataRow("documents"), DataRow("null documents")]
        [DataRow("project"), DataRow("final project"), DataRow("state")]
        [DataRow("wrong owner"), DataRow("late owner"), DataRow("late project")]
        [DataRow("missing"), DataRow("alias"), DataRow("oversized")]
        public void WordAssociationBalancesEveryAcquisitionAndRetainsFreshIdentityChecks(string fault)
        {
            var project = new object();
            var document = new Document { Project = project };
            var probe = new Probe { Fault = fault };
            if (fault != "missing") probe.Items.Add(document);
            if (fault == "foreign document") probe.Items.Add(new Document { Project = new object() });
            if (fault == "alias") probe.Items.Add(document);
            if (fault == "oversized") while (probe.Items.Count <= 1000) probe.Items.Add(document);
            if (fault == "empty path") document.Path = "";
            if (fault == "wrong owner") probe.Owner = 99;
            var released = new List<object>();
            string result = null;
            Action read = () => result = VbeProjectHostPath.ReadWordPath(project, probe, released.Add);
            bool success = new[] { "none", "foreign document", "empty path" }.Contains(fault);
            if (success) read();
            else Assert.ThrowsException<InvalidOperationException>(read);
            CollectionAssert.AreEqual(probe.Acquired.AsEnumerable().Reverse().ToArray(), released.ToArray());
            if (success)
            {
                Assert.AreEqual(document.Path, result);
                Assert.AreEqual(2, probe.ApplicationCalls, "The final application lookup is not cached.");
                Assert.AreEqual(1, probe.StateCalls);
                Assert.AreEqual(2, released.Count(item => ReferenceEquals(item, project)),
                    "Only the two project getter acquisitions are released; the borrowed argument is not acquired.");
            }
            if (fault == "late owner") Assert.AreEqual(1, probe.ProjectCalls,
                "A refused final PID must not cause another project getter.");
        }

        [TestMethod]
        public void WordAssociationRejectsInvalidArgumentsBeforeAcquiringAnApplication()
        {
            var probe = new Probe(); var project = new object();
            Assert.ThrowsException<ArgumentNullException>(() => VbeProjectHostPath.ReadWordPath(null, probe, _ => { }));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.ReadWordPath(project, null, _ => { }));
            Assert.ThrowsException<ArgumentNullException>(() => VbeProjectHostPath.ReadWordPath(project, probe, null));
            probe.Kind = "PowerPoint";
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.ReadWordPath(project, probe, _ => { }));
            Assert.AreEqual(0, probe.ApplicationCalls);
        }

        [TestMethod]
        public void WordAssociationReleaseFailureNeverRetriesTheFailedAcquisition()
        {
            var project = new object(); var probe = new Probe();
            probe.Items.Add(new Document { Project = project });
            int releases = 0;
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectHostPath.ReadWordPath(project, probe,
                value => { releases++; throw new InvalidOperationException("release refused"); }));
            Assert.AreEqual(1, releases);
            Assert.AreEqual(2, probe.ApplicationCalls);
        }
    }
}
