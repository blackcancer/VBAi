using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises reviewed Publisher discard through the real fixture using managed objects and only the testhost query handle.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixturePublisherTestCleanupTests
    {
        private const string ModuleName = "VBAiOfficeModule";
        private const string Source = "Public Sub Synthetic()\r\nEnd Sub\r\n";

        [TestMethod]
        public void ReviewedUnsavedTestProjectQuitsOnceWithoutSaveOrUnsupportedPublisherMembers()
        {
            Assert.IsNull(typeof(FakeDocument).GetProperty("VBProject"));
            Assert.IsNull(typeof(FakeApplication).GetProperty("VBE"));
            Assert.IsNull(typeof(FakeDocument).GetProperty("Saved").GetSetMethod());
            Assert.IsNull(typeof(FakeDocument).GetMethod("Save"));
            Assert.IsNull(typeof(FakeDocument).GetMethod("SaveAs"));
            WithFixture((fixture, app, sources, process) => {
                fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                Assert.AreEqual(0, app.Commands.Count, "Approval must not save, select, execute or dispatch a native command.");
                fixture.Dispose();
                Assert.AreEqual(1, app.QuitCalls);
                Assert.IsFalse(app.Document.Saved);
                Assert.IsFalse((bool)app.Persistence["ProjectSaved"]);
                Assert.IsNull(Field(fixture, "ownedProcess"));
                Assert.IsFalse(fixture.NativeExecutionUnsettled);
                CollectionAssert.AreEquivalent(new[] { "debug_state", "project_persistence_status", "list_modules", "read_module" },
                    app.Commands.Distinct().ToArray());
                var lifecycle = ((OfficeOwnedShutdownEvidence)Field(fixture, "shutdownEvidence")).Record;
                Assert.AreEqual(1, lifecycle["QuitEntries"]); Assert.AreEqual("RETURNED", lifecycle["QuitOutcome"]);
                Assert.AreEqual(true, lifecycle["ProcessExitObserved"]); Assert.AreEqual(0, lifecycle["ExitCode"]);
                fixture.Dispose(); Assert.AreEqual(1, app.QuitCalls);
            });
        }

        [TestMethod]
        public void CleanupApprovalRequiresOwnedSettledSyntheticRowAndOriginalProcessIdentity()
        {
            Action<OfficeVbeFixture, FakeApplication, IDictionary<string, string>>[] changes = {
                (fixture, app, sources) => Property(fixture, "Kind", "Word"),
                (fixture, app, sources) => fixture.NativeExecutionUnsettled = true,
                (fixture, app, sources) => Property(fixture, "ProcessId", -1),
                (fixture, app, sources) => Property(fixture, "Project", null),
                (fixture, app, sources) => sources.Remove(ModuleName),
                (fixture, app, sources) => sources.Remove(VbaTestRuntimeSource.ModuleName),
                (fixture, app, sources) => sources[VbaTestRuntimeSource.ModuleName] = "unreviewed support",
                (fixture, app, sources) => Evidence(fixture)["ProcessStartedUtc"] = "another generation",
                (fixture, app, sources) => Evidence(fixture)["OriginalProcessHandle"] = "another handle",
                (fixture, app, sources) => Evidence(fixture)["ProcessImage"] = "another image"
            };
            foreach (var change in changes)
                WithFixture((fixture, app, sources, process) => {
                    change(fixture, app, sources);
                    Assert.ThrowsException<AssertFailedException>(() => fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources));
                    Assert.AreEqual(0, app.QuitCalls); Assert.AreEqual(0, app.Commands.Count);
                    Assert.IsNull(Field(fixture, "publisherTestCleanupSources"));
                });
            WithFixture((fixture, app, sources, process) => {
                Set(fixture, "owned", false);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources));
                Assert.AreEqual(0, app.QuitCalls); Assert.AreEqual(0, app.Commands.Count);
            });
            WithFixture((fixture, app, sources, process) => {
                fixture.Dispatch = (_, __) => { throw new InvalidOperationException("pending native request"); };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("status"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources));
                Assert.AreEqual(0, app.QuitCalls); Assert.AreEqual(0, app.Commands.Count);
            });
            WithFixture((fixture, app, sources, process) => {
                Assert.ThrowsException<AssertFailedException>(() => fixture.AllowReviewedPublisherTestCleanup("OtherModule", sources));
                Exception refusal = null;
                var worker = new Thread(() => { try { fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources); } catch (Exception error) { refusal = error; } });
                worker.Start(); Assert.IsTrue(worker.Join(5000));
                Assert.IsInstanceOfType(refusal, typeof(AssertFailedException));
                fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                Assert.ThrowsException<AssertFailedException>(() => fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources));
                Assert.AreEqual(0, app.QuitCalls);
            });
        }

        [TestMethod]
        public void ChangedDocumentWindowProjectModeAndPersistenceRefuseQuitAndRetainOwnership()
        {
            Action<OfficeVbeFixture, FakeApplication>[] changes = {
                (fixture, app) => app.Document.PathState += ".foreign",
                (fixture, app) => app.Document.Window.HandleState = 0,
                (fixture, app) => app.ActiveState = new FakeDocument { PathState = app.Document.FullName, Window = app.Document.Window },
                (fixture, app) => app.Documents.CurrentState = new FakeDocument { PathState = app.Document.FullName, Window = app.Document.Window },
                (fixture, app) => app.Documents.CountState = 2,
                (fixture, app) => fixture.ReadPublisherWindowOwner = window => 0,
                (fixture, app) => app.Mode = 1,
                (fixture, app) => app.Persistence["IdentityVerified"] = false,
                (fixture, app) => app.Persistence["HostAvailable"] = false,
                (fixture, app) => app.Persistence["HostPath"] = "C:\\foreign.pub",
                (fixture, app) => app.Persistence["Host"] = "Word",
                (fixture, app) => app.Persistence["OwnerProcessId"] = -1,
                (fixture, app) => app.Persistence["Project"] = "ForeignProject",
                (fixture, app) => Property(fixture, "Project", "ChangedSelector"),
                (fixture, app) => fixture.NativeExecutionUnsettled = true,
                (fixture, app) => Evidence(fixture)["ProcessStartedUtc"] = "replaced generation"
            };
            foreach (var change in changes)
                WithFixture((fixture, app, sources, process) => {
                    fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                    change(fixture, app);
                    RequireRetainedRefusal(fixture, app, process);
                });
        }

        [TestMethod]
        public void FullSourceInventoryAndSecondReadRefuseRevokedReview()
        {
            Action<FakeApplication>[] changes = {
                app => app.Sources.Remove(ModuleName),
                app => app.Sources.Add("UnexpectedModule", "new source"),
                app => app.Sources[ModuleName] += "' unreviewed change",
                app => app.AfterSourceRead = () => app.Sources[ModuleName] += "' changed after first captured response"
            };
            foreach (var change in changes)
                WithFixture((fixture, app, sources, process) => {
                    fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                    change(app);
                    RequireRetainedRefusal(fixture, app, process);
                });
            WithFixture((fixture, app, sources, process) => {
                fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                sources[ModuleName] = "caller mutated approval input";
                fixture.Dispose(); Assert.AreEqual(1, app.QuitCalls, "The approved snapshot must be immutable.");
            });
        }

        [TestMethod]
        public void UnknownBridgeOrQuitAndUnobservedExitCannotBeRetriedOrReleased()
        {
            foreach (string failure in new[] { "bridge", "quit", "exit" })
                WithFixture((fixture, app, sources, process) => {
                    fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources);
                    if (failure == "bridge") app.DispatchError = new InvalidOperationException("bridge outcome unknown");
                    if (failure == "quit") app.QuitError = new InvalidOperationException("Quit outcome unknown");
                    if (failure == "exit") fixture.WaitForOwnedExit = (_, __) => false;
                    if (failure == "bridge") Assert.ThrowsException<AggregateException>(() => fixture.Dispose());
                    else Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                    Assert.AreSame(process, Field(fixture, "ownedProcess"));
                    Assert.IsTrue((bool)Field(fixture, "hostTeardownRefused"));
                    int quits = app.QuitCalls;
                    Assert.AreEqual(failure == "bridge" ? 0 : 1, quits);
                    if (failure == "bridge") Assert.ThrowsException<AggregateException>(() => fixture.Dispose());
                    else Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                    Assert.AreEqual(quits, app.QuitCalls);
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.AllowReviewedPublisherTestCleanup(ModuleName, sources));
                    if (failure != "exit") Assert.AreSame(app, Field(fixture, "application"));
                });
        }

        [TestMethod]
        public void SourceReviewNormalizesOnlyLineEndingsAndRefusesOtherChanges()
        {
            var expected = new Dictionary<string, string> { [ModuleName] = Source };
            OfficeVbeFixture.RequirePublisherTestCleanupSources(expected, new Dictionary<string, string> { [ModuleName] = Source.Replace("\r\n", "\n").TrimEnd('\n') });
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherTestCleanupSources(expected, new Dictionary<string, string>()));
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherTestCleanupSources(expected, new Dictionary<string, string> { ["DifferentModule"] = Source }));
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherTestCleanupSources(expected, new Dictionary<string, string> { [ModuleName] = Source + "' changed" }));
        }

        private static void RequireRetainedRefusal(OfficeVbeFixture fixture, FakeApplication app, Process process)
        {
            Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
            Assert.AreEqual(0, app.QuitCalls);
            Assert.AreSame(process, Field(fixture, "ownedProcess")); Assert.AreSame(app, Field(fixture, "application"));
            Assert.IsTrue((bool)Field(fixture, "hostTeardownRefused"));
            Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose()); Assert.AreEqual(0, app.QuitCalls);
        }

        private static void WithFixture(Action<OfficeVbeFixture, FakeApplication, IDictionary<string, string>, Process> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Publisher-test-cleanup-mirror-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "Disposable.pub"); File.WriteAllText(path, "Managed fixture only");
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            var process = Process.GetCurrentProcess();
            var app = new FakeApplication(); app.Document.PathState = path; app.ActiveState = app.Document;
            app.Documents.CurrentState = app.Document;
            Set(fixture, "application", app); Set(fixture, "document", app.Document); Set(fixture, "owned", true); Set(fixture, "ownedProcess", process);
            Property(fixture, "Kind", "Publisher"); Property(fixture, "Root", root); Property(fixture, "DocumentPath", path);
            Property(fixture, "ProcessId", process.Id); Property(fixture, "Project", "SyntheticPublisher");
            Set(fixture, "shutdownEvidence", new OfficeOwnedShutdownEvidence(process.Id, process.StartTime.ToUniversalTime().ToString("o"),
                ExcelOwnedProcessImage.Read(process.Handle), "0x" + unchecked((ulong)process.Handle.ToInt64()).ToString("X16"),
                typeof(VbeSession).Module.ModuleVersionId.ToString("D")));
            fixture.ReadPublisherWindowOwner = window => (uint)process.Id;
            fixture.ReadPublisherProcessCanary = () => new object[0];
            fixture.WaitForOwnedExit = (observed, timeout) => { Assert.AreSame(process, observed); Assert.AreEqual(5000, timeout); return true; };
            fixture.ReadOwnedExitCode = observed => 0;
            app.Persistence = new Dictionary<string, object> { ["Project"] = "SyntheticPublisher", ["HostAvailable"] = true,
                ["Host"] = "Publisher", ["HostPath"] = path, ["OwnerProcessId"] = process.Id,
                ["IdentityVerified"] = true, ["ProjectSaved"] = false, ["HostSaved"] = false };
            app.Sources = new Dictionary<string, string> { [ModuleName] = Source,
                [VbaTestRuntimeSource.ModuleName] = VbaTestRuntimeSource.Generate(new VbaTestCatalog()) };
            var expected = new Dictionary<string, string>(app.Sources);
            fixture.Dispatch = (pid, request) => {
                if (app.DispatchError != null) throw app.DispatchError;
                var values = (IDictionary<string, object>)request;
                Assert.AreEqual("SyntheticPublisher", values["Project"]); Assert.AreEqual(process.Id, pid);
                string command = (string)values["Command"]; app.Commands.Add(command);
                object data;
                if (command == "project_persistence_status") data = app.Persistence;
                else if (command == "debug_state") data = new Dictionary<string, object> { ["Mode"] = app.Mode };
                else if (command == "list_modules") data = app.Sources.Keys.Select(name => (object)new Dictionary<string, object> { ["Name"] = name }).ToArray();
                else if (command == "read_module")
                {
                    data = new Dictionary<string, object> { ["Code"] = app.Sources[(string)values["Module"]] };
                    var after = app.AfterSourceRead; app.AfterSourceRead = null; after?.Invoke();
                }
                else throw new InvalidOperationException("An unexpected incidental command was emitted: " + command);
                return new Dictionary<string, object> { ["Ok"] = true, ["Error"] = null, ["Data"] = data };
            };
            try { action(fixture, app, expected, process); }
            finally
            {
                var retained = (List<OfficeVbeFixture>)typeof(OfficeVbeFixture).GetField("retainedOfficeFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                retained.Remove(fixture); process.Dispose(); Directory.Delete(root, true);
            }
        }

        private static IDictionary<string, object> Evidence(OfficeVbeFixture fixture) => ((OfficeOwnedShutdownEvidence)Field(fixture, "shutdownEvidence")).Record;
        private static object Field(object target, string name) => typeof(OfficeVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void Set(object target, string name, object value) => typeof(OfficeVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static void Property(object target, string name, object value) => typeof(OfficeVbeFixture).GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);

        // These public members correspond to Publisher's real interfaces: no VBE, VBProject or writable Saved member exists.
        public sealed class FakeApplication
        {
            public FakeDocument Document = new FakeDocument(), ActiveState;
            public FakeDocuments Documents = new FakeDocuments();
            public FakeDocument ActiveDocument => ActiveState;
            public FakeWindow ActiveWindow => Document.ActiveWindow;
            public IDictionary<string, object> Persistence;
            public IDictionary<string, string> Sources;
            public readonly List<string> Commands = new List<string>();
            public int QuitCalls, Mode = 2;
            public Exception DispatchError, QuitError;
            public Action AfterSourceRead;
            public void Quit() { QuitCalls++; if (QuitError != null) throw QuitError; }
        }
        public sealed class FakeDocument
        {
            public string PathState;
            public FakeWindow Window = new FakeWindow();
            public string FullName => PathState;
            public bool Saved => false;
            public bool ReadOnly => false;
            public FakeWindow ActiveWindow => Window;
        }
        public sealed class FakeDocuments
        {
            public int CountState = 1;
            public FakeDocument CurrentState;
            public int Count => CountState;
            public FakeDocument this[int index] => CurrentState;
        }
        public sealed class FakeWindow
        {
            public long HandleState = 123;
            public long Hwnd => HandleState;
            public string Caption => "Disposable.pub";
        }
    }
}
