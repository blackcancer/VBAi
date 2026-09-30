namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    // Matrix established before execution:
    // State: empty, duplicate identity, source/mode/protection/saved/path drift, read failure.
    // Create: correct native type, null/missing/stale version, running/protected collection,
    // native exception after mutation, no insertion, two insertions, wrong return, changed prior project.
    // Open: null/relative/extension/absent/empty/oversize/already-open path; exact readback,
    // wrong returned path and native exception; preserve every prior project.
    // Close: null/path/version, host type, protection, running, unsaved, path drift;
    // correct removal, cancelled removal, native exception, changed prior project/read failure.
    // Every uncertain case invokes once and never rolls back or retries.
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProjectLifecycleTests
    {
        [TestMethod]
        public void CollectionVersionTracksSourceSavedModeProtectionAndIdentity()
        {
            var f = Create();
            dynamic initial = f.Service.ProjectCollectionState();
            Assert.AreEqual(64, ((string)initial.Version).Length);
            var project = f.Vbe.VBProjects.Items[0];
            var request = f.CollectionRequest();
            project.Saved = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateStandaloneProject(request));
            request = f.CollectionRequest(); project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateStandaloneProject(request));
            project.Mode = 2; request = f.CollectionRequest();
            project.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("Module", 1));
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateStandaloneProject(request));
            Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
            f.Vbe.VBProjects.Items.Add(new LifecycleProject { Name = project.Name, FileName = project.FileName });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ProjectCollectionState());
        }

        [TestMethod]
        public void CreateUsesNativeStandaloneTypeAndRejectsIncompleteOrUnsafeRequests()
        {
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Service.CreateStandaloneProject(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.CreateStandaloneProject(new Request()));
            foreach (int fault in new[] { 0, 1 })
            {
                f = Create();
                if (fault == 0) f.Vbe.VBProjects.Items[0].Mode = 1;
                else f.Vbe.VBProjects.Items[0].Protection = 1;
                var request = f.CollectionRequest();
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateStandaloneProject(request));
                Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
            }
            f = Create();
            dynamic result = f.Service.CreateStandaloneProject(f.CollectionRequest());
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(101, f.Vbe.VBProjects.LastType);
            Assert.AreEqual(2, f.Vbe.VBProjects.Items.Count);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void CreateReadsBackUnsavedStandalonePathErrorWithoutRetryingTheNativeAdd(bool clrMapped)
        {
            var f = Create();
            var request = f.CollectionRequest();
            f.Vbe.VBProjects.After = () => {
                var added = f.Vbe.VBProjects.Items[1];
                added.Saved = false;
                added.FileNameError = unchecked((int)0x800A004C);
                added.ClrPathNotFound = clrMapped;
            };
            dynamic result = f.Service.CreateStandaloneProject(request);
            Assert.IsTrue((bool)result.Verified);
            Assert.IsFalse((bool)result.Uncertain);
            Assert.AreEqual("Added", (string)result.Project);
            Assert.AreEqual("", (string)result.HostPath);
            Assert.AreEqual(1, f.Vbe.VBProjects.Attempts);
            dynamic reconciled = f.Service.ProjectCollectionState();
            Assert.AreEqual(2, ((System.Collections.ICollection)reconciled.Projects).Count);
        }

        [TestMethod]
        public void PathReadFailuresRemainBlockingForSavedHostAndOtherComErrors()
        {
            foreach (bool clrMapped in new[] { false, true })
            foreach (string scenario in new[] { "saved", "host", "other-hresult" })
            {
                var f = Create();
                var project = f.Vbe.VBProjects.Items[0];
                project.Saved = scenario == "saved";
                project.Type = scenario == "host" ? 100 : 101;
                project.FileNameError = scenario == "other-hresult" ? unchecked((int)0x80004005) : unchecked((int)0x800A004C);
                project.ClrPathNotFound = clrMapped && scenario != "other-hresult";
                if (project.ClrPathNotFound) Assert.ThrowsException<System.IO.DirectoryNotFoundException>(() => f.Service.ProjectCollectionState(), scenario);
                else Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => f.Service.ProjectCollectionState(), scenario);
                Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
            }
        }

        [TestMethod]
        public void CreateReturnsUncertaintyForEveryNativeMutationOrReadbackFaultWithoutRecovery()
        {
            for (int fault = 0; fault < 8; fault++)
            {
                var f = Create(); var request = f.CollectionRequest(); var collection = f.Vbe.VBProjects;
                if (fault == 0) collection.Fail = true;
                if (fault == 1) collection.NoChange = true;
                if (fault == 2) collection.ReturnExisting = true;
                collection.After = () =>
                {
                    if (fault == 3) collection.Items.Add(new LifecycleProject { Name = "Third", FileName = "" });
                    if (fault == 4) collection.Items[0].Saved = false;
                    if (fault == 5) collection.ThrowEnumeration = true;
                    if (fault == 6) collection.Items[1].Type = 100;
                    if (fault == 7) collection.Items[1].Mode = 1;
                };
                dynamic result = f.Service.CreateStandaloneProject(request);
                Assert.IsTrue((bool)result.Uncertain);
                Assert.IsFalse((bool)result.Verified);
                Assert.IsFalse((bool)result.RetryAllowed);
                Assert.AreEqual(1, collection.Attempts);
            }
        }

        [TestMethod]
        public void SolidWorksOpenRefusesBeforeNativeCollectionAccessOrMutation()
        {
            var f = Create();
            var probe = new VbeSolidWorksPersistenceTests.Probe();
            f.Service.SolidWorksSaveProbe = () => probe;
            string path = Path.Combine(Path.GetTempPath(), "VBAi-refused-open-" + Guid.NewGuid().ToString("N") + ".swp");
            try
            {
                File.WriteAllText(path, "owned synthetic fixture");
                var request = f.CollectionRequest(path);
                f.Vbe.VBProjects.ThrowEnumeration = true;
                var error = Assert.ThrowsException<NotSupportedException>(() => f.Service.OpenStandaloneProject(request));
                StringAssert.Contains(error.Message, "SOLIDWORKS");
                Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
                Assert.AreEqual(0, probe.Saves);
                Assert.AreEqual(0, probe.Selections);
                Assert.AreEqual("owned synthetic fixture", File.ReadAllText(path));
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void OpenValidatesFileAndVerifiesExactNativePath()
        {
            string path = Path.Combine(Path.GetTempPath(), "VBAi-lifecycle-" + Guid.NewGuid().ToString("N") + ".swp");
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Service.OpenStandaloneProject(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.OpenStandaloneProject(new Request { Path = "relative.swp" }));
            Assert.ThrowsException<ArgumentException>(() => f.Service.OpenStandaloneProject(new Request { Path = Path.ChangeExtension(path, ".bas") }));
            Assert.ThrowsException<FileNotFoundException>(() => f.Service.OpenStandaloneProject(new Request { Path = path }));
            try
            {
                File.WriteAllBytes(path, new byte[0]);
                Assert.ThrowsException<IOException>(() => f.Service.OpenStandaloneProject(new Request { Path = path }));
                File.WriteAllText(path, "fixture");
                dynamic result = f.Service.OpenStandaloneProject(f.CollectionRequest(path));
                Assert.IsTrue((bool)result.Verified);
                Assert.AreEqual(path, (string)result.HostPath);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.OpenStandaloneProject(f.CollectionRequest(path)));
                f = Create(); f.Vbe.VBProjects.WrongPath = true;
                result = f.Service.OpenStandaloneProject(f.CollectionRequest(path));
                Assert.IsTrue((bool)result.Uncertain);
                Assert.AreEqual(1, f.Vbe.VBProjects.Attempts);
                using (var stream = new FileStream(path, FileMode.Open)) stream.SetLength(64L * 1024 * 1024 + 1);
                Assert.ThrowsException<IOException>(() => f.Service.OpenStandaloneProject(new Request { Path = path }));
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void CloseRefusesHostProtectedRunningUnsavedStaleAndWrongPathProjects()
        {
            for (int fault = 0; fault < 6; fault++)
            {
                var f = Create(); var project = f.Vbe.VBProjects.Items[0];
                project.FileName = Path.Combine(Path.GetTempPath(), "Existing.swp");
                var request = f.CloseRequest(project);
                if (fault == 0) project.Type = 100;
                if (fault == 1) project.Protection = 1;
                if (fault == 2) project.Mode = 1;
                if (fault == 3) project.Saved = false;
                if (fault == 4) request.ExpectedHostPath = Path.Combine(Path.GetTempPath(), "Other.swp");
                if (fault == 5) request.ExpectedProjectVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.CloseStandaloneProject(request));
                Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
            }
            var empty = Create();
            Assert.ThrowsException<ArgumentException>(() => empty.Service.CloseStandaloneProject(null));
            Assert.ThrowsException<ArgumentException>(() => empty.Service.CloseStandaloneProject(new Request()));
        }

        [TestMethod]
        public void CloseChecksExactCollectionRemovalAndDoesNotRetryUncertainNativeCalls()
        {
            for (int fault = 0; fault < 5; fault++)
            {
                var f = Create(); var project = f.Vbe.VBProjects.Items[0];
                project.FileName = Path.Combine(Path.GetTempPath(), "Existing.swp");
                var collection = f.Vbe.VBProjects;
                collection.Items.Add(new LifecycleProject { Name = "Other", FileName = "" });
                var request = f.CloseRequest(project);
                if (fault == 1) collection.NoChange = true;
                if (fault == 2) collection.Fail = true;
                collection.After = () =>
                {
                    if (fault == 3) collection.Items[0].Saved = false;
                    if (fault == 4) collection.ThrowEnumeration = true;
                };
                dynamic result = f.Service.CloseStandaloneProject(request);
                Assert.AreEqual(fault == 0, (bool)result.Verified);
                Assert.AreEqual(fault != 0, (bool)result.Uncertain);
                Assert.AreEqual(1, collection.Attempts);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeProjectLifecycleTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
        public void CollectionFingerprintIncludesNativeReferenceFlagsVersionsAndSource()
        {
            var f = Create(); var project = f.Vbe.VBProjects.Items[0];
            project.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("M", 1));
            var reference = new LifecycleReference { GUID = "owned-guid", Major = 1, Minor = 0, BuiltIn = true };
            project.References.Add(reference);
            dynamic before = f.Service.ProjectCollectionState();
            foreach (int change in new[] { 0, 1, 2, 3, 4 })
            {
                if (change == 0) reference.GUID = "changed-guid";
                if (change == 1) reference.Major++;
                if (change == 2) reference.Minor++;
                if (change == 3) reference.IsBroken = true;
                if (change == 4) reference.BuiltIn = false;
                dynamic after = f.Service.ProjectCollectionState();
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual((string)before.Version, (string)after.Version);
                before = after;
            }
            foreach (var component in project.VBComponents) { component.CodeModule.Source = ""; component.CodeModule.LineCount = 0; }
            dynamic emptySource = f.Service.ProjectCollectionState();
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual((string)before.Version, (string)emptySource.Version);
            before = emptySource;
            project.Protection = 1;
            dynamic protectedState = f.Service.ProjectCollectionState();
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual((string)before.Version, (string)protectedState.Version);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
        public void CloseSelectsTheExactPathAmongProjectsWithTheSameDisplayName()
        {
            var f = Create(); var project = f.Vbe.VBProjects.Items[0];
            project.FileName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "selected.swp");
            var other = new LifecycleProject { Name = project.Name, FileName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "other.swp") };
            f.Vbe.VBProjects.Items.Insert(0, other);
            dynamic metadata = f.Service.ProjectProperties(project.FileName);
            dynamic result = f.Service.CloseStandaloneProject(new Request { Project = project.FileName, ExpectedHostPath = project.FileName, ExpectedProjectVersion = metadata.Version });
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue((bool)result.Verified);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, f.Vbe.VBProjects.Attempts);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, f.Vbe.VBProjects.Items.Count);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(other, f.Vbe.VBProjects.Items[0]);
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class VbeProjectLifecycleTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
        public void CloseRejectsAStandaloneTypeWhoseNativePathIsNotASwpMacro()
        {
            var f = Create(); var project = f.Vbe.VBProjects.Items[0];
            project.FileName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "host-document.xlsm");
            var request = f.CloseRequest(project);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => f.Service.CloseStandaloneProject(request));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, f.Vbe.VBProjects.Attempts);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(project, f.Vbe.VBProjects.Items[0]);
        }
    }
}
