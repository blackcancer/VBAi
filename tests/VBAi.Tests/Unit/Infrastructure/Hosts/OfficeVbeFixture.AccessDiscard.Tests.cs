using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises discard authority, native-readback races and real fixture teardown using managed fakes only.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureAccessDiscardTests
    {
        private const string Module = "VBAiOfficeModule";

        [TestMethod]
        public void UnsavedModuleMissingFromAllModulesIsCancelledOnceBeforeTheOriginalSingleCloseAndQuit()
        {
            WithFixture((fixture, app, process, root) => {
                fixture.Data("create_module", "Module", Module, "ExpectedMode", 2);
                int clicks = 0, processId = process.Id;
                fixture.ScanAccessDiscard = lease => lease.TryCancel(new IntPtr(100), window => Dialog(process.Id), cancel => {
                    Assert.AreEqual(new IntPtr(102), cancel); Interlocked.Increment(ref clicks); app.CancelObserved.Set(); return true;
                });
                fixture.Dispose();
                Assert.AreEqual(0, app.CurrentProject.AllModules.Count);
                Assert.AreEqual(1, app.CloseEntries); Assert.AreEqual(1, app.QuitEntries); Assert.AreEqual(2, app.QuitOption);
                Assert.AreEqual(1, clicks); Assert.AreEqual(0, app.Commands.CloseEntries);
                var evidence = Read(Path.Combine(root, "access-discard-dialogs.json"));
                Assert.AreEqual(1, evidence["CancelEntries"]); Assert.AreEqual(0, evidence["SaveEntries"]);
                Assert.AreEqual(true, evidence["CloseReturned"]); Assert.AreEqual("RETURNED", evidence["CancelOutcome"]);
                Assert.AreEqual(Module, evidence["Module"]); Assert.AreEqual("RichEdit20W", evidence["EditClass"]); Assert.AreEqual(processId, evidence["ProcessId"]);
                Assert.AreNotEqual(evidence["OwnerThread"], evidence["WorkerThread"]);
                Assert.AreEqual(false, evidence["MacroReplay"]); Assert.AreEqual(false, evidence["QuitReplay"]);
                fixture.Dispose(); Assert.AreEqual(1, app.CloseEntries); Assert.AreEqual(1, app.QuitEntries);
            });
        }

        [TestMethod]
        public void OnlySuccessfulExactProjectAndReturnedCreationNamesAcquireDiscardAuthority()
        {
            WithFixture((fixture, app, process, root) => {
                foreach (string command in new[] { "create_module", "create_class", "create_form" })
                {
                    string key = command == "create_form" ? "Form" : "Module";
                    fixture.Data(command, key, command, "ExpectedMode", 2);
                }
                fixture.Dispatch = (pid, request) => Reply(true, "OtherProject", "Module", Module);
                fixture.Data("create_module", "Module", Module);
                fixture.Dispatch = (pid, request) => Reply(false, "OwnedProject", "Module", Module);
                fixture.Response("create_module", "Module", Module);
                fixture.Dispatch = (pid, request) => Reply(true, "OwnedProject", "Module", "OtherReturnedName");
                fixture.Data("create_module", "Module", Module);
                fixture.Dispatch = (pid, request) => Reply(true, "OwnedProject", "Module", Module);
                fixture.Data("read_module", "Module", Module);
                var tracked = (HashSet<string>)Field(fixture, "createdAccessDiscardObjects");
                CollectionAssert.AreEquivalent(new[] { "create_module", "create_class", "create_form" }, new List<string>(tracked));
            });
        }

        [TestMethod]
        public void ForeignUnknownDisabledOrReplacedNativeDialogsAreNeverClicked()
        {
            Action<OfficeVbeFixture.AccessDiscardDialog>[] changes = {
                value => value.ProcessId++, value => value.EditProcessId++, value => value.CancelProcessId++,
                value => value.ThreadId = 0, value => value.EditThreadId++, value => value.CancelThreadId++,
                value => value.WindowClass = "Other", value => value.EditClass = "Other", value => value.EditClass = "Edit", value => value.CancelClass = "Other",
                value => value.EditId = 1, value => value.CancelId = 1, value => value.Visible = false, value => value.Enabled = false,
                value => value.Title = "Different dialog", value => value.CancelText = "OK", value => value.Name = "UntrackedUserModule",
                value => value.Window = IntPtr.Zero, value => value.Edit = IntPtr.Zero, value => value.Cancel = IntPtr.Zero
            };
            foreach (var change in changes)
            {
                var lease = Lease(); lease.BindWorkerThread(); int clicks = 0;
                lease.TryCancel(new IntPtr(100), window => { var value = Dialog(42); change(value); return value; }, cancel => { clicks++; return true; });
                Assert.AreEqual(0, clicks); Assert.AreEqual(0, lease.Record["CancelEntries"]);
            }
            foreach (var change in new Action<OfficeVbeFixture.AccessDiscardDialog>[] {
                value => value.Edit = new IntPtr(201), value => value.Cancel = new IntPtr(202),
                value => value.Window = new IntPtr(200), value => value.Title = "Save As",
                value => { value.ThreadId++; value.EditThreadId++; value.CancelThreadId++; }
            })
            {
                var lease = Lease(); lease.BindWorkerThread(); int reads = 0, clicks = 0;
                lease.TryCancel(new IntPtr(100), window => { var value = Dialog(42); if (++reads == 2) change(value); return value; }, cancel => { clicks++; return true; });
                Assert.AreEqual(2, reads); Assert.AreEqual(0, clicks);
            }
        }

        [TestMethod]
        public void RevokedOrForeignThreadAuthorityRefusesCancelAndNormalStopDoesNotBecomeAnUncertainMutation()
        {
            bool authority = true;
            var lease = Lease(() => authority); lease.BindWorkerThread(); int reads = 0, clicks = 0;
            Assert.ThrowsException<InvalidOperationException>(() => lease.TryCancel(new IntPtr(100), window => {
                if (++reads == 2) authority = false; return Dialog(42);
            }, cancel => { clicks++; return true; }));
            Assert.AreEqual(0, clicks); Assert.AreEqual(0, lease.Record["CancelEntries"]);
            var stopping = Lease(); stopping.BindWorkerThread(); bool stop = false; stopping.Stopping = () => stop;
            stopping.TryCancel(new IntPtr(100), window => { stop = true; return Dialog(42); }, cancel => { clicks++; return true; });
            Assert.AreEqual(0, clicks);
            var ownThread = Lease(); ownThread.BindWorkerThread(); Exception error = null;
            var thread = new Thread(() => { try { ownThread.TryCancel(new IntPtr(100), window => Dialog(42), cancel => true); } catch (Exception caught) { error = caught; } });
            thread.Start(); Assert.IsTrue(thread.Join(2000)); Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
            Assert.ThrowsException<InvalidOperationException>(() => ownThread.BindWorkerThread());
        }

        [TestMethod]
        public void UnknownCancelRetainsTheGenerationAndDoesNotEmitQuitOrRetryClose()
        {
            WithFixture((fixture, app, process, root) => {
                fixture.Data("create_module", "Module", Module);
                int clicks = 0, processId = process.Id;
                fixture.ScanAccessDiscard = lease => {
                    try { lease.TryCancel(new IntPtr(100), window => Dialog(process.Id), cancel => { Interlocked.Increment(ref clicks); return false; }); }
                    finally { app.CancelObserved.Set(); }
                };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(1, clicks); Assert.AreEqual(1, app.CloseEntries); Assert.AreEqual(0, app.QuitEntries);
                Assert.AreSame(process, Field(fixture, "ownedProcess")); Assert.AreSame(app, Field(fixture, "application"));
                Assert.AreEqual(true, Field(fixture, "hostTeardownRefused"));
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(1, clicks); Assert.AreEqual(1, app.CloseEntries); Assert.AreEqual(0, app.QuitEntries);
                Assert.AreEqual("CALL_FAILED_EFFECT_UNKNOWN", Read(Path.Combine(root, "access-discard-dialogs.json"))["CancelOutcome"]);
            });
            foreach (bool throws in new[] { false, true })
            {
                var lease = Lease(); lease.BindWorkerThread(); int entries = 0;
                Assert.ThrowsException<InvalidOperationException>(() => lease.TryCancel(new IntPtr(100), window => Dialog(42), cancel => {
                    entries++; if (throws) throw new InvalidOperationException("uncertain Cancel"); return false;
                }));
                Assert.ThrowsException<InvalidOperationException>(() => lease.TryCancel(new IntPtr(200), window => Dialog(42), cancel => { entries++; return true; }));
                Assert.AreEqual(1, entries); Assert.AreEqual(1, lease.Record["CancelEntries"]);
            }
        }

        [TestMethod]
        public void PendingExecutionRefusesDiscardPreparationAndVisibleCancelledDialogRefusesQuit()
        {
            WithFixture((fixture, app, process, root) => {
                fixture.NativeExecutionUnsettled = true;
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(0, app.CloseEntries); Assert.AreEqual(0, app.QuitEntries);
                Assert.IsFalse(File.Exists(Path.Combine(root, "access-discard-dialogs.json")));
            });
            WithFixture((fixture, app, process, root) => {
                fixture.Data("create_module", "Module", Module);
                fixture.IsVisibleOwnedAccessDiscard = (window, original, start) => true;
                fixture.ScanAccessDiscard = lease => lease.TryCancel(new IntPtr(100), window => Dialog(process.Id), cancel => { app.CancelObserved.Set(); return true; });
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(1, app.CloseEntries); Assert.AreEqual(0, app.QuitEntries);
                Assert.AreEqual(true, Field(fixture, "hostTeardownRefused"));
            });
        }

        [TestMethod]
        public void AttestedRichEdit20WNameIsReadByBoundedUnicodeWindowMessageOnAnOwnedHiddenControl()
        {
            // This is a hidden Win32 control in the testhost, not an Access instance.
            IntPtr library = LoadLibrary(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "riched20.dll"));
            Assert.AreNotEqual(IntPtr.Zero, library);
            IntPtr richEdit = IntPtr.Zero, ordinaryEdit = IntPtr.Zero;
            try
            {
                richEdit = CreateWindowEx(0, "RichEdit20W", Module, 0, 0, 0, 100, 20, IntPtr.Zero, IntPtr.Zero, library, IntPtr.Zero);
                ordinaryEdit = CreateWindowEx(0, "Edit", Module, 0, 0, 0, 100, 20, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                Assert.AreNotEqual(IntPtr.Zero, richEdit); Assert.AreNotEqual(IntPtr.Zero, ordinaryEdit);
                uint pid; GetWindowThreadProcessId(richEdit, out pid);
                using (var current = Process.GetCurrentProcess()) Assert.AreEqual((uint)current.Id, pid);
                Assert.AreEqual(Module, OfficeVbeFixture.ReadNativeAccessDiscardName(richEdit));
                Assert.IsNull(OfficeVbeFixture.ReadNativeAccessDiscardName(ordinaryEdit));
                Assert.IsNull(OfficeVbeFixture.ReadNativeAccessDiscardName(IntPtr.Zero));
            }
            finally
            {
                if (ordinaryEdit != IntPtr.Zero) Assert.IsTrue(DestroyWindow(ordinaryEdit));
                if (richEdit != IntPtr.Zero) Assert.IsTrue(DestroyWindow(richEdit));
                Assert.IsTrue(FreeLibrary(library));
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr LoadLibrary(string path);
        [DllImport("kernel32.dll")] private static extern bool FreeLibrary(IntPtr module);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateWindowEx(uint extended, string kind, string title, uint style, int x, int y, int width, int height,
            IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        private static OfficeVbeFixture.AccessDiscardLease Lease(Func<bool> authority = null)
        {
            return new OfficeVbeFixture.AccessDiscardLease(42, "start", "handle", "path", "project", Thread.CurrentThread.ManagedThreadId,
                new[] { Module }, authority ?? (() => true), record => { });
        }
        private static OfficeVbeFixture.AccessDiscardDialog Dialog(int pid)
        {
            return new OfficeVbeFixture.AccessDiscardDialog {
                Window = new IntPtr(100), Edit = new IntPtr(101), Cancel = new IntPtr(102),
                ProcessId = pid, EditProcessId = pid, CancelProcessId = pid,
                ThreadId = 3, EditThreadId = 3, CancelThreadId = 3, WindowClass = "#32770", EditClass = "RichEdit20W", CancelClass = "Button",
                EditId = 2020, CancelId = 2, Visible = true, Enabled = true, Title = "Enregistrer sous", Name = Module, CancelText = "Annuler" };
        }
        private static IDictionary<string, object> Reply(bool ok, string project, string key, string name)
        {
            return new Dictionary<string, object> { ["Ok"] = ok, ["Error"] = ok ? null : "synthetic refusal",
                ["Data"] = new Dictionary<string, object> { ["Project"] = project, [key] = name } };
        }
        private static void WithFixture(Action<OfficeVbeFixture, FakeApplication, Process, string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-AccessDiscard-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true); var app = new FakeApplication();
            SetProperty(fixture, "Kind", "Access"); SetProperty(fixture, "Root", root); SetProperty(fixture, "Project", "OwnedProject");
            SetProperty(fixture, "DocumentPath", Path.Combine(root, "ManagedFakeOnly.accdb"));
            using (var current = Process.GetCurrentProcess()) SetProperty(fixture, "ProcessId", current.Id);
            SetField(fixture, "owned", true); SetField(fixture, "application", app);
            typeof(OfficeVbeFixture).GetMethod("CaptureOwnedProcess", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fixture, null);
            var process = (Process)Field(fixture, "ownedProcess");
            fixture.RequireAdapterOnlyCleanup(); fixture.ValidateAccessDiscardDocument = () => { Assert.AreEqual(true, Field(fixture, "owned")); Assert.IsFalse(process.HasExited); };
            fixture.IsVisibleOwnedAccessDiscard = (window, original, start) => false;
            fixture.Dispatch = (pid, request) => {
                var fields = (IDictionary<string, object>)request; string key = (string)fields["Command"] == "create_form" ? "Form" : "Module";
                return Reply(true, "OwnedProject", key, Convert.ToString(fields[key]));
            };
            fixture.WaitForOwnedExit = (original, bound) => { Assert.AreSame(process, original); Assert.AreEqual(15000, bound); return true; };
            fixture.ReadOwnedExitCode = original => 0;
            try { action(fixture, app, process, root); }
            finally
            {
                var worker = Field(fixture, "dialogWorker");
                if (worker != null)
                {
                    SetField(worker, "StopRequested", true);
                    var thread = (Thread)Field(worker, "Thread"); if (thread != null) Assert.IsTrue(thread.Join(2000));
                }
                process.Dispose(); app.CancelObserved.Dispose();
                var retained = (IList)typeof(OfficeVbeFixture).GetField("retainedOfficeFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                lock (retained) retained.Remove(fixture); Directory.Delete(root, true);
            }
        }
        private static object Field(object value, string name) => value.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(value);
        private static void SetField(object value, string name, object content) => value.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(value, content);
        private static void SetProperty(object value, string name, object content) => value.GetType().GetProperty(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(value, content);
        private static IDictionary<string, object> Read(string path) => (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path));
        public sealed class FakeApplication
        {
            public readonly FakeProject CurrentProject = new FakeProject(); public readonly FakeCommands Commands = new FakeCommands();
            public FakeCommands DoCmd => Commands;
            public readonly ManualResetEventSlim CancelObserved = new ManualResetEventSlim();
            public int CloseEntries, QuitEntries, QuitOption;
            public void CloseCurrentDatabase() { CloseEntries++; Assert.IsTrue(CancelObserved.Wait(3000), "Synthetic unsaved module naming prompt was not handled."); }
            public void Quit(int option) { QuitEntries++; QuitOption = option; }
        }
        public sealed class FakeProject { public FakeCollection AllForms => new FakeCollection(); public FakeCollection AllReports => new FakeCollection(); public FakeCollection AllModules => new FakeCollection(); }
        public sealed class FakeCollection { public int Count => 0; public object this[object index] => throw new InvalidOperationException("An empty saved-object inventory must never be indexed."); }
        public sealed class FakeCommands { public int CloseEntries; public void Close(int kind, string name, int save) { CloseEntries++; Assert.Fail("The unsaved module is absent from AllModules."); } }
    }
}