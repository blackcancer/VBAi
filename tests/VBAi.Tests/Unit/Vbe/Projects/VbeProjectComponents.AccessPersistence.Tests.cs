using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks deferred Access saved notifications with fake native objects on an owned STA.</summary>
    public sealed partial class VbeOtherHostPersistenceTests
    {
        /// <summary>Models Access's native factory getters while keeping every true project/context guard observable.</summary>
        [AccessStaTestMethod]
        [DataRow("unchanged"), DataRow("reopenedproject"), DataRow("mappedproject"), DataRow("application")]
        [DataRow("path"), DataRow("pid"), DataRow("owner"), DataRow("documentmissing"), DataRow("documentduplicate")]
        [DataRow("source"), DataRow("metadata"), DataRow("references"), DataRow("readonly"), DataRow("format")]
        [DataRow("mode"), DataRow("protected"), DataRow("pane")]
        public void AccessFactoryDocumentIdentityUsesApprovedApplicationPathAndMappedProject(string change)
        {
            var f = new AsyncAccessFixture();
            object application = new object(), mappedProject = f.Probe.Project;
            var wrappers = new List<AccessFactoryDocument>();
            var stateReads = new List<AccessFactoryDocument>();
            bool absent = false, duplicate = false;
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            f.Probe.ReadApplication = () => application;
            f.Probe.ReadDocuments = app =>
            {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                var wrapper = new AccessFactoryDocument
                {
                    Project = mappedProject,
                    Path = f.Probe.Observation.Path,
                    Format = f.Probe.Observation.Format,
                    ReadOnly = f.Probe.Observation.ReadOnly
                };
                wrappers.Add(wrapper);
                if (wrappers.Count >= 4) f.Probe.Project.Saved = true;
                return absent ? new List<object>() : duplicate ? new List<object> { wrapper, wrapper } : new List<object> { wrapper };
            };
            f.Probe.ReadDocumentProject = document => ((AccessFactoryDocument)document).Project;
            f.Probe.ReadDocumentState = document =>
            {
                var wrapper = (AccessFactoryDocument)document;
                stateReads.Add(wrapper);
                return new VbeProjectComponents.OtherHostDocumentState
                {
                    Path = wrapper.Path,
                    Format = wrapper.Format,
                    ReadOnly = wrapper.ReadOnly,
                    Saved = null
                };
            };
            f.Probe.AfterInvocation = () => { f.Probe.Observation.Format = 12; f.Probe.Project.Saved = false; };
            var pending = f.SaveAsync(f.Probe.Request());
            Assert.IsFalse(pending.IsCompleted);
            var replacement = new OtherProject { Name = "P", FileName = f.Probe.Project.FileName, Saved = true };
            switch (change)
            {
                case "reopenedproject": f.Editor.VBProjects[0] = replacement; f.Editor.ActiveVBProject = replacement; mappedProject = replacement; break;
                case "mappedproject": mappedProject = replacement; break;
                case "application": application = new object(); break; // Same PID is insufficient to substitute the approved application.
                case "path": f.Probe.Observation.Path = @"C:\fixture\Other.accdb"; break;
                case "pid": f.Probe.ProcessId++; break;
                case "owner": f.Probe.Owner++; break;
                case "documentmissing": absent = true; break;
                case "documentduplicate": duplicate = true; break;
                case "source": f.Probe.Component.CodeModule.Source += "' changed"; break;
                case "metadata": f.Probe.Project.Description = "Changed metadata"; break;
                case "references": f.Probe.Project.References.Add(new AsyncAccessReference()); break;
                case "readonly": f.Probe.Observation.ReadOnly = true; break;
                case "format": f.Probe.Observation.Format = 99; break;
                case "mode": f.Probe.Project.Mode = 1; break;
                case "protected": f.Probe.Project.Protection = 1; break;
                case "pane": f.Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = f.Probe.Component } }; break;
            }
            if (change != "unchanged") f.Probe.Project.Saved = true;
            dynamic result = CompleteAccessSave(pending);
            Assert.AreEqual(1, f.Probe.Attempts, "The pending Save must never be replayed.");
            if (change == "unchanged")
            {
                Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
                Assert.IsTrue(wrappers.Count >= 4, "Exercise at least two deferred factory observations.");
                Assert.AreSame(wrappers[wrappers.Count - 1], stateReads[stateReads.Count - 1], "State must come from the current wrapper, not the retained original.");
                for (int index = 1; index < wrappers.Count; index++) Assert.AreNotSame(wrappers[0], wrappers[index]);
            }
            else
            {
                Assert.IsFalse((bool)result.Verified, change); Assert.IsTrue((bool)result.Uncertain, change);
                Assert.IsFalse(string.IsNullOrWhiteSpace((string)result.Reason));
            }
        }

        /// <summary>A fresh getter result holds a snapshot; consulting the initial result hides later native path/state changes.</summary>
        private sealed class AccessFactoryDocument
        {
            internal object Project;
            internal string Path;
            internal int? Format;
            internal bool ReadOnly;
        }

        /// <summary>Reproduces the single Save returning before Access processes its Saved notification.</summary>
        [AccessStaTestMethod]
        public void AccessSaveYieldsForDelayedOwnerThreadSavedReadbackWithoutInvokingAgain()
        {
            var f = new AsyncAccessFixture();
            int owner = Thread.CurrentThread.ManagedThreadId;
            using (var ui = new Control())
            {
                _ = ui.Handle;
                f.Probe.AfterInvocation = () =>
                {
                    f.Probe.Observation.Format = 12;
                    f.Probe.Observation.Saved = null;
                    f.Probe.Project.Saved = false;
                    ui.BeginInvoke(new Action(() =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        f.Probe.Project.Saved = true;
                    }));
                };
                var pending = f.SaveAsync(f.Probe.Request());
                Assert.IsFalse(pending.IsCompleted, "Access must yield after its single Save before reading the deferred saved flag.");
                dynamic result = CompleteAccessSave(pending);
                Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
                Assert.IsNull((bool?)result.HostSaved);
                Assert.AreEqual(1, f.Probe.Attempts);
                Assert.IsFalse((bool)result.PersistenceReopenVerified);
            }
        }

        /// <summary>Preserves exact context/revision guards throughout the yielded notification window.</summary>
        [AccessStaTestMethod]
        [DataRow("source"), DataRow("metadata"), DataRow("references"), DataRow("path")]
        [DataRow("projectpath"), DataRow("projectidentity"), DataRow("documentidentity"), DataRow("pane")]
        [DataRow("component"), DataRow("activeproject"), DataRow("mode"), DataRow("protected")]
        [DataRow("pid"), DataRow("owner"), DataRow("host"), DataRow("readonly"), DataRow("format")]
        [DataRow("missingfile"), DataRow("emptyfile")]
        public void AccessChangesDuringDeferredVerificationRemainUncertainWithoutAnotherSave(string change)
        {
            var f = new AsyncAccessFixture();
            f.Probe.AfterInvocation = () => { f.Probe.Observation.Format = 12; f.Probe.Observation.Saved = null; f.Probe.Project.Saved = false; };
            var pending = f.SaveAsync(f.Probe.Request());
            Assert.IsFalse(pending.IsCompleted);
            switch (change)
            {
                case "source": f.Probe.Component.CodeModule.Source += "' intervening source"; break;
                case "metadata": f.Probe.Project.Description = "Intervening metadata"; break;
                case "references": f.Probe.Project.References.Add(new AsyncAccessReference()); break;
                case "path": f.Probe.Observation.Path = @"C:\fixture\Other.accdb"; break;
                case "projectpath": f.Probe.Project.FileName = @"C:\fixture\Other.accdb"; break;
                case "projectidentity": f.Editor.VBProjects[0] = new OtherProject { Name = "P" }; break;
                case "documentidentity": f.Probe.Items[0] = new object(); f.Probe.ReadDocumentProject = value => ReferenceEquals(value, f.Probe) ? f.Probe.Project : null; break;
                case "pane": f.Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = f.Probe.Component } }; break;
                case "component": f.Editor.ActiveCodePane.CodeModule.Parent = new object(); break;
                case "activeproject": f.Editor.ActiveVBProject = new object(); break;
                case "mode": f.Probe.Project.Mode = 1; break;
                case "protected": f.Probe.Project.Protection = 1; break;
                case "pid": f.Probe.ProcessId++; break;
                case "owner": f.Probe.Owner++; break;
                case "host": f.Probe.Kind = "Word"; break;
                case "readonly": f.Probe.Observation.ReadOnly = true; break;
                case "format": f.Probe.Observation.Format = 99; break;
                case "missingfile": f.Probe.Exists = false; break;
                case "emptyfile": f.Probe.Bytes = 0; break;
            }
            f.Probe.Project.Saved = true;
            dynamic result = CompleteAccessSave(pending);
            Assert.IsFalse((bool)result.Verified, change); Assert.IsTrue((bool)result.Uncertain, change);
            Assert.AreEqual(1, f.Probe.Attempts, change);
        }

        /// <summary>A failed Saved notification expires read-only, clears ownership, and never repeats native Save.</summary>
        [AccessStaTestMethod]
        public void AccessDeferredTimeoutAndCrossSessionPendingGuardNeverReplayNativeSave()
        {
            var f = new AsyncAccessFixture();
            f.Service.AccessSaveVerificationTimeout = TimeSpan.FromMilliseconds(80);
            f.Probe.AfterInvocation = () =>
            {
                AccessStaTestMethodAttribute.CaptureCurrentContextAtProbe();
                f.Probe.Observation.Format = 12; f.Probe.Project.Saved = false;
            };
            var pending = f.SaveAsync(f.Probe.Request());
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(f.SaveAsync(f.Probe.Request()))).Message,
                "An Access save is already awaiting verification");
            var second = new AsyncAccessFixture();
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(second.SaveAsync(second.Probe.Request()))).Message,
                "An Access save is already awaiting verification");
            Assert.AreEqual(0, second.Probe.Attempts);
            dynamic expired = CompleteAccessSave(pending);
            Assert.IsFalse((bool)expired.Verified); Assert.IsTrue((bool)expired.Uncertain);
            StringAssert.Contains((string)expired.Reason, "timed out");
            Assert.AreEqual(1, f.Probe.Attempts);
            dynamic next = CompleteAccessSave(second.SaveAsync(second.Probe.Request()));
            Assert.IsTrue((bool)next.Verified); Assert.AreEqual(1, second.Probe.Attempts);
        }

        /// <summary>Refuses unsafe approval snapshots before crossing the native mutation boundary.</summary>
        [AccessStaTestMethod]
        [DataRow("version"), DataRow("readonly"), DataRow("path"), DataRow("format")]
        [DataRow("owner"), DataRow("pane"), DataRow("selection"), DataRow("protected"), DataRow("mode")]
        [DataRow("application"), DataRow("applicationfinal")]
        public void AccessAsyncPreflightRefusesUnsafeRequestsBeforeSave(string failure)
        {
            var f = new AsyncAccessFixture(); var request = f.Probe.Request();
            switch (failure)
            {
                case "version": request.ExpectedProjectVersion = "stale"; break;
                case "readonly": f.Probe.Observation.ReadOnly = true; break;
                case "path": request.ExpectedHostPath = @"C:\fixture\Other.accdb"; break;
                case "format": f.Probe.Observation.Format = 99; break;
                case "owner": f.Probe.Owner = 999; break;
                case "pane": f.Editor.ActiveCodePane = null; break;
                case "selection": f.Editor.ActiveVBProject = null; break;
                case "protected": f.Probe.Project.Protection = 1; break;
                case "mode": f.Probe.Project.Mode = 1; break;
                case "application":
                case "applicationfinal":
                    object originalApplication = new object(); int reads = 0;
                    f.Probe.ReadApplication = () => ++reads <= (failure == "application" ? 2 : 3) ? originalApplication : new object();
                    break;
            }
            var refusal = Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(f.SaveAsync(request)), failure);
            Assert.IsFalse(refusal.Message.Contains("already awaiting verification"), "A leftover task must not satisfy a preflight refusal oracle.");
            Assert.AreEqual(0, f.Probe.Attempts, failure);
        }

        /// <summary>Preserves a native error unchanged and refuses the old synchronous mutation route.</summary>
        [AccessStaTestMethod]
        public void AccessSynchronousDispatchAndNativeFailuresCannotBypassDeferredContract()
        {
            var f = new AsyncAccessFixture();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveHostDocument(f.Probe.Request()));
            Assert.AreEqual(0, f.Probe.Attempts);
            f.Probe.Failure = "native error";
            dynamic failed = CompleteAccessSave(f.SaveAsync(f.Probe.Request()));
            Assert.IsFalse((bool)failed.Verified); Assert.IsTrue((bool)failed.Uncertain);
            StringAssert.Contains((string)failed.Reason, "Native save failed after invocation");
            Assert.AreEqual(1, f.Probe.Attempts);
        }

        /// <summary>Checks the final built-in command boundary against the exact approved pane and component.</summary>
        [DataTestMethod, DataRow("unchanged"), DataRow("pane"), DataRow("component"), DataRow("closed"), DataRow("beforeSaveGuard")]
        public void AccessNativeCommandRechecksApprovedSelectionImmediatelyBeforeItsSingleInvocation(string change)
        {
            var editor = new AsyncAccessEditor();
            var project = new PathProject { VBE = editor, FileName = @"C:\fixture\Owned.accdb" };
            object component = new object();
            editor.VBProjects.Add(project); editor.ActiveVBProject = project;
            editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = component } };
            var app = new PathApplication(); var document = new PathDocument { Application = app, FullName = project.FileName };
            app.CurrentProject = document;
            var native = new VbeProjectComponents.NativeOtherHostProbe
            {
                ReadHostKind = () => "Access",
                ReadIdentity = ReferenceEquals,
                ReadOwner = window => (uint)Process.GetCurrentProcess().Id,
                ReadActiveApplication = name => app
            };
            native.BindProject(project); native.BindDocument(document);
            native.BindAccessSaveSelection(editor.ActiveCodePane, component);
            native.PrepareSave(document);
            if (change == "pane") editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = component } };
            if (change == "component") editor.ActiveCodePane.CodeModule.Parent = new object();
            if (change == "closed") editor.ActiveCodePane = null;
            if (change == "beforeSaveGuard") native.AccessBeforeSave = () =>
            {
                Assert.IsFalse(native.SaveInvocationStarted);
                throw new InvalidOperationException("An owned modal or revoked authorization prevents the command.");
            };
            if (change == "unchanged")
            {
                native.Save(document, false, project.FileName, 12);
                Assert.AreEqual(1, editor.CommandBars.Control.Executions);
            }
            else
            {
                Assert.ThrowsException<InvalidOperationException>(() => native.Save(document, false, project.FileName, 12));
                Assert.AreEqual(0, editor.CommandBars.Control.Executions);
                Assert.IsFalse(native.SaveInvocationStarted);
            }
        }

        /// <summary>One synthetic confirmation must yield before the original Save receives verified saved-state evidence.</summary>
        [AccessStaTestMethod]
        public void AccessConfirmationQueuesOnceAndStillRequiresFullOwnerThreadSavedVerification()
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            confirmation.AfterEnqueue = () => f.Probe.Project.Saved = true;
            dynamic result = CompleteAccessSave(f.SaveAsync(f.Probe.Request()));
            Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
            Assert.AreEqual(1, f.Probe.Attempts); Assert.AreEqual(1, confirmation.ConfirmEntries);
            Assert.AreEqual(1, (int)result.ConfirmationAttempts);
            Assert.IsTrue((bool)result.ConfirmationQueued); Assert.IsFalse((bool)result.ConfirmationPending);
            Assert.IsFalse((bool)result.PersistenceReopenVerified);
            Assert.IsTrue(confirmation.ObserveEntries >= 2, "Enqueue must be followed by another read-only observation.");
            Assert.AreEqual(1, confirmation.PrepareEntries);
        }

        /// <summary>A preexisting prompt is refused before the original native Save, without attaching to an earlier uncertain operation.</summary>
        [AccessStaTestMethod]
        public void AccessPreexistingConfirmationRefusesBeforeOriginalSaveOrEnqueue()
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            confirmation.PreparationFailure = new InvalidOperationException("Preexisting owned save prompt");
            Assert.AreSame(confirmation.PreparationFailure,
                Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(f.SaveAsync(f.Probe.Request()))));
            Assert.AreEqual(0, f.Probe.Attempts); Assert.AreEqual(0, confirmation.ObserveEntries);
            Assert.AreEqual(0, confirmation.ConfirmEntries); Assert.AreEqual(0, confirmation.ConfirmationAttempts);
            Assert.IsFalse(confirmation.ConfirmationQueued);
        }

        /// <summary>Changes inside the confirmation boundary must be caught before its one possible native enqueue.</summary>
        [AccessStaTestMethod]
        [DataRow("owner"), DataRow("pid"), DataRow("source"), DataRow("references"), DataRow("metadata")]
        [DataRow("mode"), DataRow("protected"), DataRow("path"), DataRow("projectpath"), DataRow("readonly")]
        [DataRow("pane"), DataRow("activeproject"), DataRow("format"), DataRow("application")]
        [DataRow("authorization")]
        public void AccessConfirmationRevalidatesApprovedContextBeforeAnyEnqueue(string change)
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            object application = f.Probe;
            f.Probe.ReadApplication = () => application;
            var request = f.Probe.Request();
            if (change == "authorization")
                request.RevalidateSaveAuthorization = () =>
                {
                    if (f.Probe.Attempts != 0) throw new InvalidOperationException("Save authorization was revoked before confirmation.");
                };
            confirmation.BeforeContextCheck = () =>
            {
                switch (change)
                {
                    case "owner": f.Probe.Owner++; break;
                    case "pid": f.Probe.ProcessId++; break;
                    case "source": f.Probe.Component.CodeModule.Source += "' intervening source"; break;
                    case "references": f.Probe.Project.References.Add(new AsyncAccessReference()); break;
                    case "metadata": f.Probe.Project.Description = "Intervening metadata"; break;
                    case "mode": f.Probe.Project.Mode = 1; break;
                    case "protected": f.Probe.Project.Protection = 1; break;
                    case "path": f.Probe.Observation.Path = @"C:\fixture\Other.accdb"; break;
                    case "projectpath": f.Probe.Project.FileName = @"C:\fixture\Other.accdb"; break;
                    case "readonly": f.Probe.Observation.ReadOnly = true; break;
                    case "pane": f.Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = f.Probe.Component } }; break;
                    case "activeproject": f.Editor.ActiveVBProject = new object(); break;
                    case "format": f.Probe.Observation.Format = 99; break;
                    case "application": application = new object(); break;
                }
            };
            dynamic result = CompleteAccessSave(f.SaveAsync(request));
            Assert.IsFalse((bool)result.Verified, change); Assert.IsTrue((bool)result.Uncertain, change);
            Assert.AreEqual(1, f.Probe.Attempts, "The original ID3 Save is never replayed.");
            Assert.AreEqual(1, confirmation.ConfirmEntries, "Exercise the immediate confirmation-context guard.");
            Assert.AreEqual(0, confirmation.ConfirmationAttempts, "No native enqueue occurs after context refusal.");
            Assert.AreEqual(0, (int)result.ConfirmationAttempts);
            Assert.IsFalse((bool)result.ConfirmationQueued); Assert.IsTrue((bool)result.ConfirmationPending);
            Assert.IsFalse(string.IsNullOrWhiteSpace((string)result.Reason));
        }

        /// <summary>Queued confirmation is not persistence; a remaining prompt or unsaved project must expire without another action.</summary>
        [AccessStaTestMethod]
        [DataRow(true), DataRow(false)]
        public void AccessQueuedConfirmationWithoutSavedCompletionRemainsUncertain(bool promptRemains)
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            f.Service.AccessSaveVerificationTimeout = TimeSpan.FromMilliseconds(80);
            confirmation.PromptRemainsAfterEnqueue = promptRemains;
            if (promptRemains) confirmation.AfterEnqueue = () => f.Probe.Project.Saved = true;
            dynamic result = CompleteAccessSave(f.SaveAsync(f.Probe.Request()));
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
            Assert.AreEqual(1, f.Probe.Attempts); Assert.AreEqual(1, confirmation.ConfirmEntries);
            Assert.AreEqual(1, (int)result.ConfirmationAttempts);
            Assert.IsTrue((bool)result.ConfirmationQueued);
            Assert.AreEqual(promptRemains, (bool)result.ConfirmationPending);
        }

        /// <summary>A verification deadline already elapsed before confirmation must never enqueue a late native action.</summary>
        [AccessStaTestMethod]
        public void AccessConfirmationDeadlineBeforeConfirmNeverQueuesOrRetries()
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            f.Service.AccessSaveVerificationTimeout = TimeSpan.FromMilliseconds(1);
            dynamic result = CompleteAccessSave(f.SaveAsync(f.Probe.Request()));
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
            Assert.AreEqual(1, f.Probe.Attempts); Assert.AreEqual(0, confirmation.ConfirmEntries);
            Assert.AreEqual(0, (int)result.ConfirmationAttempts);
            Assert.IsFalse((bool)result.ConfirmationQueued); Assert.IsTrue((bool)result.ConfirmationPending);
        }

        /// <summary>A known failure from the initial native call cannot be converted into another native confirmation.</summary>
        [AccessStaTestMethod]
        public void AccessInitialNativeFailureNeverObservesOrConfirmsItsPrompt()
        {
            var f = new AsyncAccessFixture();
            var confirmation = AttachAccessConfirmation(f);
            f.Probe.Failure = "native error";
            dynamic result = CompleteAccessSave(f.SaveAsync(f.Probe.Request()));
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
            StringAssert.Contains((string)result.Reason, "Native save failed after invocation");
            Assert.AreEqual(1, f.Probe.Attempts); Assert.AreEqual(0, confirmation.ObserveEntries);
            Assert.AreEqual(0, confirmation.ConfirmEntries); Assert.AreEqual(0, confirmation.ConfirmationAttempts);
            Assert.IsFalse(confirmation.ConfirmationQueued);
        }

        /// <summary>Injects a purely synthetic modal boundary; none of its numeric window tokens refer to native windows.</summary>
        private static FakeAccessConfirmation AttachAccessConfirmation(AsyncAccessFixture fixture)
        {
            var confirmation = new FakeAccessConfirmation();
            fixture.Service.AccessSaveConfirmationFactory = (window, pid, components) =>
            {
                Assert.AreEqual(new IntPtr(77), window); Assert.AreEqual(fixture.Probe.ProcessId, pid);
                int count = 0;
                foreach (var component in components)
                {
                    Assert.AreEqual(fixture.Probe.Component.Name, component.Name);
                    Assert.AreEqual(1, component.Type); count++;
                }
                Assert.AreEqual(1, count); return confirmation;
            };
            fixture.Probe.AfterInvocation = () =>
            {
                fixture.Probe.Observation.Format = 12; fixture.Probe.Observation.Saved = null;
                fixture.Probe.Project.Saved = false;
            };
            return confirmation;
        }

        /// <summary>Observes real service approval checks while replacing only dialog reads and enqueue with owned-STA counters.</summary>
        private sealed class FakeAccessConfirmation : VbeProjectComponents.IAccessSaveConfirmation
        {
            private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
            private readonly VbeProjectComponents.AccessSaveConfirmationCandidate candidate;
            internal int PrepareEntries, ObserveEntries, ConfirmEntries;
            internal Exception PreparationFailure;
            internal Action BeforeContextCheck, AfterEnqueue;
            internal bool PromptRemainsAfterEnqueue;
            public int ConfirmationAttempts { get; private set; }
            public bool ConfirmationQueued { get; private set; }
            public bool ConfirmationPending { get; private set; }
            internal FakeAccessConfirmation()
            {
                candidate = new VbeProjectComponents.AccessSaveConfirmationCandidate(this,
                    new VbeProjectComponents.AccessSaveDialogSnapshot
                    {
                        Window = new IntPtr(1),
                        ProcessId = 42,
                        ThreadId = 7,
                        Controls = new[] { new VbeProjectComponents.AccessSaveDialogControl { Id = 1, Window = new IntPtr(2) } }
                    }, "synthetic-owner-thread-candidate");
            }
            public void Prepare()
            {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId); PrepareEntries++;
                if (PreparationFailure != null) throw PreparationFailure;
            }
            public void RequireBeforeSave() { Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId); }
            public VbeProjectComponents.AccessSaveConfirmationCandidate Observe()
            {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId); ObserveEntries++;
                ConfirmationPending = !ConfirmationQueued || PromptRemainsAfterEnqueue;
                return ConfirmationPending ? candidate : null;
            }
            public void Confirm(VbeProjectComponents.AccessSaveConfirmationCandidate value, Action revalidateApprovedContext, Action requireDeliveryDeadline = null)
            {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId); Assert.AreSame(candidate, value);
                Assert.AreEqual(0, ConfirmEntries, "No synthetic confirmation call may be repeated."); ConfirmEntries++;
                BeforeContextCheck?.Invoke();
                revalidateApprovedContext();
                requireDeliveryDeadline?.Invoke();
                ConfirmationAttempts = 1; ConfirmationQueued = true;
                AfterEnqueue?.Invoke();
            }
        }

        /// <summary>Pumps only the detached test dispatcher; no Office application or macro is created.</summary>
        private static object CompleteAccessSave(Task<object> pending)
        {
            var watch = Stopwatch.StartNew();
            int doEventsTurns = 0;
            while (!pending.IsCompleted && watch.Elapsed.TotalSeconds < 5)
            { doEventsTurns++; Application.DoEvents(); Thread.Sleep(1); }
            if (!pending.IsCompleted) Assert.Fail("Deferred Access verification did not complete. " + AccessStaTestMethodAttribute.DescribeTimeout(pending, doEventsTurns));
            return pending.GetAwaiter().GetResult();
        }

        /// <summary>Retains real service revision checks while replacing the native Office/COM boundary.</summary>
        private sealed class AsyncAccessFixture
        {
            internal readonly Fixture Probe = new Fixture { Kind = "Access" };
            internal readonly AsyncAccessEditor Editor = new AsyncAccessEditor();
            internal readonly VbeProjectComponents Service;
            internal Task<object> SaveAsync(Request request) => AccessStaTestMethodAttribute.StartSave(() => Service.SaveHostDocumentAsync(request));
            internal AsyncAccessFixture()
            {
                Assert.IsTrue(Application.MessageLoop, "The synthetic Access fixture requires its dedicated WinForms test loop.");
                Probe.Observation.Path = @"C:\fixture\Owned.accdb";
                Probe.Project.FileName = Probe.Observation.Path;
                Probe.Observation.Format = 12; Probe.Observation.Saved = null; Probe.Project.Saved = false;
                Editor.VBProjects.Add(Probe.Project); Editor.ActiveVBProject = Probe.Project;
                Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = Probe.Component } };
                Service = new VbeProjectComponents(Editor, new VbeForms(Editor));
                Service.OtherHostProbe = () => Probe;
                Probe.AfterInvocation = () => { Probe.Observation.Format = 12; Probe.Observation.Saved = null; };
            }
        }

        /// <summary>Simulates selection identity without supplying a different revision implementation.</summary>
        public sealed class AsyncAccessEditor
        {
            public PathWindow MainWindow { get; } = new PathWindow();
            public PathCommands CommandBars { get; } = new PathCommands();
            public List<object> VBProjects { get; } = new List<object>();
            public object ActiveVBProject { get; set; }
            public AsyncAccessPane ActiveCodePane { get; set; }
        }
        /// <summary>Provides only the native pane identity and its selected component association.</summary>
        public sealed class AsyncAccessPane { public AsyncAccessCode CodeModule { get; set; } }
        /// <summary>Identifies the fake native code component selected for Save.</summary>
        public sealed class AsyncAccessCode { public object Parent { get; set; } }

        /// <summary>Changes a native reference revision without evaluating any external library.</summary>
        public sealed class AsyncAccessReference
        {
            public string GUID => "{420B2830-E718-11CF-893D-00A0C9054228}";
            public int Major => 1;
            public int Minor => 0;
            public bool IsBroken => false;
            public bool BuiltIn => false;
        }
    }
}
