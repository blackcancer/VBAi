using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeSolidWorksMacroCreationTests
    {
        [TestMethod]
        public void ModalOwnerPumpDeliversExactlyOnceAndVerifiesOnlyAfterOriginalReturn()
        {
            var f = new Fixture();
            var task = f.Begin(); f.Scheduler.Dispatch();
            var result = task.GetAwaiter().GetResult();
            Assert.IsTrue(result.Verified); Assert.IsFalse(result.Uncertain);
            Assert.IsTrue(result.OriginalCommandReturned); Assert.IsTrue(result.DialogClosed);
            Assert.AreEqual(1, f.Native.Creates); Assert.AreEqual(1, f.Native.Writes); Assert.AreEqual(1, f.Native.Saves);
            Assert.AreEqual(1, f.Verifications); Assert.AreEqual(1, f.Native.Disposes);
            CollectionAssert.AreEqual(new[] { "command", "filename", "save", "terminal" }, f.Claims.ToArray());
            Assert.IsFalse(result.RetryAllowed); Assert.IsFalse(result.PersistenceReloadVerified);
        }

        [TestMethod]
        public void NestedTimerDuringAuthorizationCannotDeliverTwice()
        {
            var f = new Fixture(); f.Live = () => f.Scheduler.Tick();
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Verified);
            Assert.AreEqual(1, f.Native.Writes); Assert.AreEqual(1, f.Native.Saves);
        }

        [TestMethod]
        public void ExistingModalRefusesBeforeCommandEntry()
        {
            var f = new Fixture(); f.Native.PrepareError = new InvalidOperationException("existing modal");
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsFalse(task.Result.Uncertain); Assert.IsFalse(task.Result.CommandEntered);
            Assert.AreEqual(0, f.Native.Creates); Assert.AreEqual(0, f.Native.Writes); Assert.AreEqual(0, f.Native.Saves);
            Assert.AreEqual(1, f.Native.Disposes);
        }

        [TestMethod]
        public void CommandClaimFailurePreventsNativeEntry()
        {
            var f = new Fixture(); f.ClaimErrorPhase = "command";
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsFalse(task.Result.CommandEntered); Assert.AreEqual(0, f.Native.Creates);
            Assert.AreEqual(0, f.Native.Writes); Assert.AreEqual(0, f.Native.Saves);
        }

        [TestMethod]
        public void DeadlineBeforeCommandPreventsEveryMutation()
        {
            var f = new Fixture(); var task = f.Begin(); f.Scheduler.Time = 100; f.Scheduler.Dispatch();
            Assert.IsFalse(task.Result.CommandEntered); Assert.AreEqual(0, f.Native.Creates);
            Assert.AreEqual(0, f.Native.Saves); Assert.IsFalse(task.Result.Uncertain);
        }

        [TestMethod]
        public void ChangedFilenameReadbackCannotEnqueueSave()
        {
            var f = new Fixture(); f.Native.WrongReadback = true;
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Uncertain); Assert.AreEqual(1, f.Native.Writes);
            Assert.AreEqual(0, f.Native.Saves); Assert.AreEqual(0, f.Verifications);
        }

        [TestMethod]
        public void DeadlineDuringFinalNativeReadPreventsActualSaveDelivery()
        {
            var f = new Fixture(); f.Native.BeforeSaveDelivery = () => f.Scheduler.Time = 100;
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Uncertain); Assert.AreEqual(1, f.Native.Writes);
            Assert.AreEqual(0, f.Native.Saves); Assert.IsFalse(task.Result.SaveQueued);
        }

        [TestMethod]
        public void SaveClaimFailureLeavesPartialDialogWithoutRetryOrCancel()
        {
            var f = new Fixture(); f.ClaimErrorPhase = "save";
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Uncertain); Assert.AreEqual(1, f.Native.Writes); Assert.AreEqual(0, f.Native.Saves);
            f.Scheduler.Tick(); Assert.AreEqual(0, f.Native.Saves);
        }

        [TestMethod]
        public void RevokedAuthorizationAfterFilenameRefusesSave()
        {
            var f = new Fixture(); f.Live = () => { if (f.Native.Writes != 0) throw new InvalidOperationException("approval revoked"); };
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Uncertain); Assert.AreEqual(1, f.Native.Writes); Assert.AreEqual(0, f.Native.Saves);
        }

        [TestMethod]
        public void PendingCommandTimeoutRetainsReferencesAndNeverRepeatsSave()
        {
            var f = new Fixture(); f.Native.RemainPending = true; var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.IsCompleted);
            Assert.IsTrue(task.Result.Uncertain); Assert.IsFalse(task.Result.OriginalCommandReturned);
            Assert.AreEqual(1, f.Native.Saves); Assert.AreEqual(0, f.Native.Disposes);
            f.Scheduler.Tick(); Assert.AreEqual(1, f.Native.Saves);
        }

        [TestMethod]
        public void UnknownDialogRefusesFilenameAndSave()
        {
            var f = new Fixture(); f.Native.CaptureError = new InvalidOperationException("unknown modal");
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsTrue(task.Result.Uncertain); Assert.AreEqual(0, f.Native.Writes); Assert.AreEqual(0, f.Native.Saves);
        }

        [TestMethod]
        public void VerificationFailureDoesNotPromoteCreationOrRepeatSave()
        {
            var f = new Fixture(); f.VerifyError = true;
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsFalse(task.Result.Verified); Assert.IsTrue(task.Result.Uncertain);
            Assert.AreEqual(1, f.Native.Saves); Assert.AreEqual(1, f.Verifications);
            Assert.IsFalse(task.Result.RollbackPerformed);
            Assert.IsFalse(task.Result.DestinationIdentityVerified); Assert.IsFalse(task.Result.DestinationAbsenceProven);
            StringAssert.Contains(task.Result.Limit, "partial file may exist");
        }

        [TestMethod]
        public void CleanupFailureIsPresentBeforeTerminalReceipt()
        {
            var f = new Fixture(); f.Native.DisposeError = true;
            var task = f.Begin(); f.Scheduler.Dispatch();
            Assert.IsFalse(task.Result.Verified); Assert.IsTrue(task.Result.Uncertain);
            Assert.IsFalse(f.TerminalVerified); StringAssert.Contains(task.Result.Error, "cleanup");
        }

        [TestMethod]
        public void FreshPathValidationNeverOverwritesExistingDestination()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-native-path-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory); string path = Path.Combine(directory, "macro.swp");
            try
            {
                Assert.AreEqual(path, VbeProjectComponents.RequireFreshSolidWorksMacroPath(path));
                File.WriteAllText(path, "original");
                Assert.ThrowsException<IOException>(() => VbeProjectComponents.RequireFreshSolidWorksMacroPath(path));
                Assert.AreEqual("original", File.ReadAllText(path));
                Assert.ThrowsException<ArgumentException>(() => VbeProjectComponents.RequireFreshSolidWorksMacroPath("relative.swp"));
                Assert.ThrowsException<ArgumentException>(() => VbeProjectComponents.RequireFreshSolidWorksMacroPath(Path.Combine(directory, "macro.bas")));
            }
            finally { Directory.Delete(directory, true); }
        }

        private sealed class Fixture
        {
            internal readonly Scheduler Scheduler = new Scheduler();
            internal readonly Native Native;
            internal readonly VbeProjectComponents Service;
            internal readonly List<string> Claims = new List<string>();
            internal Action Live = () => { };
            internal string ClaimErrorPhase;
            internal bool VerifyError, TerminalVerified;
            internal int Verifications;
            internal Fixture()
            {
                var vbe = new VbeProjectComponentsTests.FakeVbe(); Service = new VbeProjectComponents(vbe, new VbeForms(vbe));
                Service.SolidWorksMacroCreationTimeoutMilliseconds = 100; Native = new Native(Scheduler);
            }
            internal Task<VbeProjectComponents.SolidWorksMacroCreationResult> Begin() => Service.RunSolidWorksMacroCreationAsync(
                @"C:\synthetic\owned.swp", Native, Scheduler, () => Live(), () => { }, result =>
                {
                    Assert.IsTrue(result.OriginalCommandReturned); Verifications++;
                    if (VerifyError) throw new InvalidOperationException("final readback differs"); result.Verified = true; return result;
                }, result =>
                {
                    string phase = result.Terminal ? "terminal" : result.SaveAttempts != 0 ? "save" : result.FilenameAttempts != 0 ? "filename" : "command";
                    Claims.Add(phase); if (result.Terminal) TerminalVerified = result.Verified;
                    if (phase == ClaimErrorPhase) throw new IOException("durable claim failed");
                });
        }

        internal sealed class Scheduler : VbeProjectGeneralOperation.IScheduler
        {
            internal long Time; internal Action Posted, Polling;
            public long ElapsedMilliseconds => Time;
            public void RequireOwner() { }
            public void Post(Action action) { Posted = action; }
            public IDisposable Poll(Action action) { Polling = action; return new Subscription(this); }
            internal void Dispatch() { var action = Posted; Posted = null; action(); }
            internal void Tick() { Polling?.Invoke(); }
            private sealed class Subscription : IDisposable
            { private readonly Scheduler owner; internal Subscription(Scheduler owner) { this.owner = owner; } public void Dispose() { owner.Polling = null; } }
        }

        internal sealed class Native : VbeProjectComponents.ISolidWorksMacroCreationNative
        {
            private readonly Scheduler scheduler;
            internal int Creates, Writes, Saves, Disposes;
            internal bool WrongReadback, RemainPending, DisposeError;
            internal Exception PrepareError, CaptureError;
            internal Action AfterSave;
            internal Action BeforeSaveDelivery;
            private bool open;
            private string filename = "";
            internal Native(Scheduler scheduler) { this.scheduler = scheduler; }
            public void RequireOwner() { }
            public void Prepare() { if (PrepareError != null) throw PrepareError; }
            public void Create(Action beforeEntry)
            {
                Assert.IsNotNull(scheduler.Polling); beforeEntry(); Creates++; open = true; scheduler.Tick();
                if (RemainPending) { scheduler.Time = 100; scheduler.Tick(); throw new TimeoutException("original command did not return before its deadline"); }
            }
            public VbeProjectComponents.SolidWorksMacroDialog Capture()
            {
                if (CaptureError != null) throw CaptureError;
                return open ? new VbeProjectComponents.SolidWorksMacroDialog { Window = new IntPtr(1), Filename = new IntPtr(2), SaveButton = new IntPtr(3), Thread = 7, FilenameText = filename } : null;
            }
            public void RequireSame(VbeProjectComponents.SolidWorksMacroDialog expected) { if (!expected.Same(Capture())) throw new InvalidOperationException("changed dialog"); }
            public void WriteFilename(VbeProjectComponents.SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery)
            { beforeEntry(); beforeDelivery(); Writes++; filename = WrongReadback ? "other.swp" : path; }
            public void Save(VbeProjectComponents.SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery)
            { beforeEntry(); BeforeSaveDelivery?.Invoke(); beforeDelivery(); Saves++; if (!RemainPending) open = false; AfterSave?.Invoke(); }
            public bool Closed(VbeProjectComponents.SolidWorksMacroDialog expected) => !open;
            public bool SameProject(object first, object second) => ReferenceEquals(first, second);
            public void Dispose() { Disposes++; if (DisposeError) throw new InvalidOperationException("cleanup failed"); }
        }
    }

    public sealed partial class VbeProjectLifecycleTests
    {
        [TestMethod]
        public void NativeCreationVerifiesNewType100ThisLibraryAndPreservesCanonicalOriginal()
        { RunNativeCreationReadback(false, false, false); }

        [TestMethod]
        public void NativeCreationCannotPromoteGenericType101AsHostMacro()
        { RunNativeCreationReadback(true, false, false); }

        [TestMethod]
        public void NativeCreationRequiresActualThisLibraryDocument()
        { RunNativeCreationReadback(false, true, false); }

        [TestMethod]
        public void NativeCreationRejectsSameVersionReplacementOfPreexistingProject()
        { RunNativeCreationReadback(false, false, true); }

        [TestMethod]
        public void NativeCreationStaleCollectionRefusesBeforeNativeFactory()
        {
            var f = Create(); string directory = Path.Combine(Path.GetTempPath(), "VBAi-native-stale-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var request = f.CollectionRequest(Path.Combine(directory, "new.swp")); request.ExpectedMode = 2;
                f.Vbe.VBProjects.Items[0].Description = "changed";
                f.Service.SolidWorksMacroCreationSchedulerFactory = () => new VbeSolidWorksMacroCreationTests.Scheduler();
                int factories = 0; f.Service.SolidWorksMacroCreationNativeFactory = (v, a) => { factories++; throw new InvalidOperationException("must not create"); };
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateSolidWorksMacroAsync(request));
                Assert.AreEqual(0, factories); Assert.AreEqual(1, f.Vbe.VBProjects.Items.Count);
            }
            finally { Directory.Delete(directory, true); }
        }

        [TestMethod]
        public void NativeCreationRefusesPriorSaveOnTheSameOwnerBeforeAnyFactoryOrRead()
        {
            var f = Create(); var pending = typeof(VbeProjectComponents).GetField("solidWorksSavePending", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            bool old = (bool)pending.GetValue(null); int factories = 0;
            f.Service.SolidWorksMacroCreationNativeFactory = (v, a) => { factories++; throw new InvalidOperationException("must not create"); };
            try
            {
                pending.SetValue(null, true); f.Vbe.VBProjects.ThrowEnumeration = true;
                var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.CreateSolidWorksMacroAsync(new Request { ExpectedMode = 2 }));
                StringAssert.Contains(error.Message, "save"); Assert.AreEqual(0, factories);
            }
            finally { pending.SetValue(null, old); }
        }

        private static void RunNativeCreationReadback(bool generic, bool missingDocument, bool replaced)
        {
            var f = Create(); string directory = Path.Combine(Path.GetTempPath(), "VBAi-native-readback-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory); string path = Path.Combine(directory, "native.swp");
            try
            {
                var scheduler = new VbeSolidWorksMacroCreationTests.Scheduler();
                var native = new VbeSolidWorksMacroCreationTests.Native(scheduler);
                f.Service.SolidWorksMacroCreationSchedulerFactory = () => scheduler;
                f.Service.SolidWorksMacroCreationNativeFactory = (v, a) => native;
                var original = f.Vbe.VBProjects.Items[0];
                native.AfterSave = () =>
                {
                    File.WriteAllText(path, "owned synthetic file");
                    if (replaced) f.Vbe.VBProjects.Items[0] = new LifecycleProject { Name = original.Name, FileName = "" };
                    var target = new LifecycleProject { Name = "Native", FileName = path, Type = generic ? 101 : 100 };
                    if (!missingDocument) target.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("ThisLibrary", 100));
                    f.Vbe.VBProjects.Items.Add(target);
                };
                var request = f.CollectionRequest(path); request.ExpectedMode = 2;
                var task = f.Service.CreateSolidWorksMacroAsync(request); scheduler.Dispatch();
                bool expected = !generic && !missingDocument && !replaced;
                Assert.AreEqual(expected, task.Result.Verified); Assert.AreEqual(!expected, task.Result.Uncertain);
                Assert.AreEqual(1, native.Creates); Assert.AreEqual(1, native.Saves);
                Assert.IsTrue(File.Exists(path)); Assert.IsFalse(task.Result.RollbackPerformed);
                Assert.AreEqual(expected, task.Result.DestinationIdentityVerified);
                Assert.IsFalse(task.Result.DestinationAbsenceProven);
                if (expected) { Assert.AreEqual("Native", task.Result.Project); Assert.AreEqual(path, task.Result.HostPath); Assert.IsNotNull(task.Result.CollectionVersion); }
            }
            finally { Directory.Delete(directory, true); }
        }
    }
}
