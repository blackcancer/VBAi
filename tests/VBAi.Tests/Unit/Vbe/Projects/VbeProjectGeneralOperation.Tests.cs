using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeProjectGeneralOperationTests
    {
        internal sealed class Scheduler : VbeProjectGeneralOperation.IScheduler
        {
            internal Action Posted, Tick; internal bool Active; internal int Disposals; internal Exception DisposeError;
            internal long Time; public long ElapsedMilliseconds => Time;
            public void RequireOwner() { }
            public void Post(Action callback) { Posted = callback; }
            public IDisposable Poll(Action callback) { Tick = callback; Active = true; return new Scope(this); }
            internal void Pump() { if (Active) Tick(); }
            private sealed class Scope : IDisposable
            {
                private readonly Scheduler parent; internal Scope(Scheduler parent) { this.parent = parent; }
                public void Dispose() { parent.Disposals++; parent.Active = false; if (parent.DisposeError != null) throw parent.DisposeError; }
            }
        }
        internal sealed class Native : VbeProjectGeneralOperation.INative
        {
            internal VbeProjectGeneralOperation.Snapshot State = Snapshot();
            internal bool Visible; internal int Captures, Writes, Closes, Owners; internal int Button;
            internal int HelpFileCodePage = 1252, HelpFileRepresentationChecks;
            internal bool HelpFileUnicode = false;
            internal Action OnWrite, OnOwner, OnClose, OnRepresentation, BeforeCloseEntry, BeforeFieldEntry;
            internal Exception WriteError, CaptureError, RepresentationError, CloseError; internal bool RefuseClosure;
            public void RequireOwner() { Owners++; OnOwner?.Invoke(); }
            public void Prepare() { RequireOwner(); if (Visible) throw new InvalidOperationException("preexisting modal"); }
            public VbeProjectGeneralOperation.Snapshot Capture(string name)
            {
                Captures++; if (CaptureError != null) throw CaptureError;
                if (!Visible) return null; if (State.Name != name) throw new InvalidOperationException("unknown project"); return Copy(State);
            }
            public void RequireSame(VbeProjectGeneralOperation.Snapshot expected, bool compareValues)
            {
                if (!Visible || !expected.SameNative(State) || compareValues && expected.OptionsVersion != State.OptionsVersion)
                    throw new InvalidOperationException("changed snapshot");
            }
            public void WriteContext(VbeProjectGeneralOperation.Snapshot expected, int value, Action entry)
            {
                RequireSame(expected, true); BeforeFieldEntry?.Invoke(); entry(); Writes++; State.ContextText = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                OnWrite?.Invoke(); if (WriteError != null) throw WriteError;
            }
            public void WriteHelpFile(VbeProjectGeneralOperation.Snapshot expected, string value, Action entry)
            {
                RequireSame(expected, true); RequireHelpFileRepresentable(expected, value); BeforeFieldEntry?.Invoke(); entry(); Writes++; State.HelpFile = value; OnWrite?.Invoke(); if (WriteError != null) throw WriteError;
            }
            public void RequireHelpFileRepresentable(VbeProjectGeneralOperation.Snapshot expected, string value)
            {
                RequireSame(expected, true); HelpFileRepresentationChecks++;
                OnRepresentation?.Invoke(); if (RepresentationError != null) throw RepresentationError;
                VbeProjectGeneralNative.RequireExactTextRepresentation(value, HelpFileCodePage, HelpFileUnicode);
            }
            public void Close(VbeProjectGeneralOperation.Snapshot expected, int id, Action entry)
            { RequireSame(expected, true); BeforeCloseEntry?.Invoke(); entry(); Closes++; Button = id; OnClose?.Invoke(); if (CloseError != null) throw CloseError; if (!RefuseClosure) Visible = false; }
            public bool Closed(VbeProjectGeneralOperation.Snapshot expected) => !Visible;
        }
        internal static VbeProjectGeneralOperation.Snapshot Snapshot() => new VbeProjectGeneralOperation.Snapshot
        {
            Dialog = new IntPtr(1),
            Page = new IntPtr(2),
            Tab = new IntPtr(3),
            Context = new IntPtr(4),
            NameEdit = new IntPtr(5),
            DescriptionEdit = new IntPtr(6),
            HelpFileEdit = new IntPtr(7),
            CompilationEdit = new IntPtr(8),
            Name = "Disposable",
            Description = "baseline",
            HelpFile = "",
            ContextText = "0",
            Compilation = ""
        };
        internal static VbeProjectGeneralOperation.Snapshot Copy(VbeProjectGeneralOperation.Snapshot s) => new VbeProjectGeneralOperation.Snapshot
        {
            Dialog = s.Dialog,
            Page = s.Page,
            Tab = s.Tab,
            Context = s.Context,
            NameEdit = s.NameEdit,
            DescriptionEdit = s.DescriptionEdit,
            HelpFileEdit = s.HelpFileEdit,
            CompilationEdit = s.CompilationEdit,
            Name = s.Name,
            Description = s.Description,
            HelpFile = s.HelpFile,
            ContextText = s.ContextText,
            Compilation = s.Compilation
        };
        internal sealed class Case
        {
            internal readonly Native Native = new Native(); internal readonly Scheduler Scheduler = new Scheduler();
            internal readonly List<string> Claims = new List<string>(); internal readonly VbeProjectGeneralOperation Operation;
            internal Action Live, Pure, BeforeEntry, AfterEntry, BeforeExecuteReturns; internal Action<VbeProjectGeneralOperation.Result> Claim;
            internal Task<VbeProjectGeneralOperation.Result> Pending;
            internal Case() { Operation = new VbeProjectGeneralOperation(Native, Scheduler); }
            internal void Start(int? value = 321, string file = null, string version = null)
            {
                Pending = Operation.RunAsync("Disposable", value, version ?? Native.State.OptionsVersion, () => Live?.Invoke(), () => Pure?.Invoke(), before =>
                {
                    BeforeEntry?.Invoke(); before(); Native.Visible = true; AfterEntry?.Invoke();
                    Scheduler.Pump(); Scheduler.Pump();
                    BeforeExecuteReturns?.Invoke();
                    if (Pending.IsCompleted) Assert.IsNotNull(Pending.Result.Error, "Only an uncertain/error result may settle before original Execute returns.");
                }, result => { Claims.Add(result.Terminal ? "terminal" : result.OkAttempts == 1 ? "ok" : result.FieldAttempts == 1 ? "field" : result.CancelAttempts == 1 ? "cancel" : "open"); Claim?.Invoke(result); }, file);
            }
            internal VbeProjectGeneralOperation.Result Complete() { Scheduler.Posted(); Scheduler.Pump(); return Pending.GetAwaiter().GetResult(); }
        }
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void OneInitialUiWriteKeepsOtherFieldsAndSettlesOnlyAfterExecuteAndClosure(bool helpFile)
        {
            var c = new Case(); string file = "C:\\qualification-été\\unicode.chm";
            c.Start(helpFile ? (int?)null : 321, helpFile ? file : null); var result = c.Complete();
            Assert.IsTrue(result.Available && result.Terminal && result.DialogClosed && result.OriginalExecuteReturned && result.ControlValueVerified && result.CommittedRequested);
            Assert.IsTrue(result.MutationInvoked); Assert.IsFalse(result.Uncertain || result.PersistenceVerified || result.RetryAllowed);
            Assert.AreEqual(helpFile ? file : "", result.HelpFile); Assert.AreEqual(helpFile ? "0" : "321", result.HelpContextText);
            Assert.AreEqual("Disposable", result.Name); Assert.AreEqual("baseline", result.Description); Assert.AreEqual("", result.ConditionalCompilation);
            Assert.AreEqual(1, c.Native.Writes); Assert.AreEqual(1, c.Native.Closes); Assert.AreEqual(1, c.Native.Button);
            CollectionAssert.AreEqual(new[] { "open", "field", "ok", "terminal" }, c.Claims.ToArray());
            Assert.AreEqual(c.Native.State.OptionsVersion, result.OptionsVersion); Assert.AreEqual(1, c.Scheduler.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => c.Start()); Assert.AreEqual(1, c.Native.Writes);
        }
        [TestMethod]
        public void ReadOnlyCancelsOnceAndNeverWritesOrClaimsPersistence()
        {
            var c = new Case(); c.Start(null); var r = c.Complete();
            Assert.AreEqual(2, c.Native.Button); Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(1, r.CancelAttempts);
            Assert.IsTrue(r.DialogClosed && r.OriginalExecuteReturned && r.Terminal); Assert.IsFalse(r.MutationInvoked || r.Uncertain || r.PersistenceVerified);
            CollectionAssert.AreEqual(new[] { "open", "cancel", "terminal" }, c.Claims.ToArray());
        }
        [TestMethod]
        public void UnrepresentableAnsiHelpFileRefusesBeforeFieldClaimAndCancelsUnchangedModalOnce()
        {
            var c = new Case(); c.Start(null, @"C:\日本.chm"); var result = c.Complete();
            Assert.IsTrue(result.Terminal && result.RefusedBeforeWrite && result.DialogClosed && result.OriginalExecuteReturned);
            Assert.IsFalse(result.Uncertain || result.Available || c.Native.Visible); Assert.IsNotNull(result.Error);
            Assert.IsFalse(result.MutationInvoked || result.ControlValueVerified || result.CommittedRequested || result.RetryAllowed);
            Assert.AreEqual(0, result.FieldAttempts); Assert.AreEqual(0, result.OkAttempts); Assert.AreEqual(1, result.CancelAttempts);
            Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(1, c.Native.Closes); Assert.AreEqual(2, c.Native.Button); Assert.AreEqual(1, c.Native.HelpFileRepresentationChecks);
            CollectionAssert.AreEqual(new[] { "open", "cancel", "terminal" }, c.Claims.ToArray());
            c.Scheduler.Pump(); Assert.AreEqual(1, c.Native.HelpFileRepresentationChecks);
            Assert.ThrowsException<InvalidOperationException>(() => c.Start(null, @"C:\été.chm"));
        }
        [DataTestMethod]
        [DataRow("live")]
        [DataRow("pure")]
        [DataRow("owner")]
        [DataRow("value")]
        [DataRow("handle")]
        [DataRow("claim")]
        [DataRow("policyAfterClaim")]
        [DataRow("deadlineBeforeEnqueue")]
        public void KnownRefusalCannotCancelAfterAuthorizationIdentityClaimOrDeadlineFailure(string fault)
        {
            var c = new Case();
            c.Native.OnRepresentation = () =>
            {
                if (fault == "live") c.Live = () => { throw new InvalidOperationException("live revoked"); };
                if (fault == "pure") c.Pure = () => { throw new InvalidOperationException("policy revoked"); };
                if (fault == "owner") c.Native.OnOwner = () => { throw new InvalidOperationException("owner changed"); };
                if (fault == "value") c.Native.State.Description = "changed";
                if (fault == "handle") c.Native.State.Context = new IntPtr(99);
            };
            c.Claim = receipt =>
            {
                if (receipt.Terminal || receipt.CancelAttempts == 0) return;
                if (fault == "claim") throw new InvalidOperationException("cancel receipt failed");
                if (fault == "policyAfterClaim") c.Pure = () => { throw new InvalidOperationException("policy revoked after claim"); };
            };
            if (fault == "deadlineBeforeEnqueue") c.Native.BeforeCloseEntry = () => c.Scheduler.Time = 20001;
            c.Start(null, @"C:\日本.chm"); var r = c.Complete();
            Assert.IsTrue(r.RefusedBeforeWrite && r.Uncertain && r.Terminal && c.Native.Visible);
            Assert.AreEqual(0, r.FieldAttempts); Assert.AreEqual(0, r.OkAttempts);
            Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
            c.Scheduler.Pump(); Assert.AreEqual(0, c.Native.Closes);
        }
        [TestMethod]
        public void UnknownRepresentationFailureDoesNotEnterKnownRefusalCancelPath()
        {
            var c = new Case(); c.Native.RepresentationError = new InvalidOperationException("native code page unavailable");
            c.Start(null, @"C:\日本.chm"); var r = c.Complete();
            Assert.IsTrue(r.Uncertain && c.Native.Visible); Assert.IsFalse(r.RefusedBeforeWrite);
            Assert.AreEqual(0, r.FieldAttempts); Assert.AreEqual(0, r.CancelAttempts); Assert.AreEqual(0, c.Native.Closes);
        }
        [DataTestMethod]
        [DataRow("post")]
        [DataRow("timeout")]
        [DataRow("terminalReceipt")]
        public void RefusalCancelFailureRemainsUncertainAndNeverClosesAgain(string fault)
        {
            var c = new Case();
            if (fault == "post") c.Native.CloseError = new InvalidOperationException("Cancel enqueue failed");
            if (fault == "timeout") c.Native.RefuseClosure = true;
            if (fault == "terminalReceipt") c.Claim = receipt => { if (receipt.Terminal) throw new InvalidOperationException("terminal receipt failed"); };
            c.Start(null, @"C:\日本.chm"); c.Scheduler.Posted();
            if (fault == "timeout") { c.Scheduler.Time = 20001; c.Scheduler.Pump(); }
            var r = c.Pending.Result;
            Assert.IsTrue(r.RefusedBeforeWrite && r.Uncertain && r.Terminal); Assert.AreEqual(1, r.CancelAttempts);
            Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(1, c.Native.Closes);
            c.Scheduler.Pump(); Assert.AreEqual(1, c.Native.Closes);
        }
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ClosedCancelWithOriginalExecuteStillPendingAtDeadlineRemainsUncertain(bool refusal)
        {
            var c = new Case(); bool observedPendingExecute = false;
            c.BeforeExecuteReturns = () =>
            {
                Assert.IsFalse(c.Pending.IsCompleted, "Closed modal alone cannot settle the original Execute.");
                c.Scheduler.Time = 20001; c.Scheduler.Pump();
                var terminal = c.Pending.Result;
                observedPendingExecute = terminal.DialogClosed && !terminal.OriginalExecuteReturned;
                Assert.IsTrue(terminal.Uncertain && terminal.Terminal);
            };
            c.Start(null, refusal ? @"C:\日本.chm" : null); var r = c.Complete();
            Assert.IsTrue(observedPendingExecute && r.Uncertain); Assert.AreEqual(1, c.Native.Closes); Assert.AreEqual(0, c.Native.Writes);
        }
        [DataTestMethod]
        [DataRow("fieldClaim")]
        [DataRow("fieldEntry")]
        [DataRow("okClaim")]
        [DataRow("okEntry")]
        public void DeadlineCrossedDuringGuardsCannotEnterFieldOrOk(string phase)
        {
            var c = new Case();
            if (phase == "fieldClaim") c.Native.OnRepresentation = () => c.Scheduler.Time = 20001;
            if (phase == "fieldEntry") c.Native.BeforeFieldEntry = () => c.Scheduler.Time = 20001;
            if (phase == "okClaim") c.Native.OnWrite = () => c.Scheduler.Time = 20001;
            if (phase == "okEntry") c.Native.BeforeCloseEntry = () => c.Scheduler.Time = 20001;
            c.Start(null, @"C:\été.chm"); var result = c.Complete();
            bool fieldEntered = phase == "okClaim" || phase == "okEntry";
            Assert.IsTrue(result.Terminal && result.Uncertain && c.Native.Visible);
            Assert.AreEqual(fieldEntered, result.MutationInvoked); Assert.IsFalse(result.CommittedRequested);
            Assert.AreEqual(phase == "fieldClaim" ? 0 : 1, result.FieldAttempts);
            Assert.AreEqual(phase == "okEntry" ? 1 : 0, result.OkAttempts);
            Assert.AreEqual(fieldEntered ? 1 : 0, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
            Assert.AreEqual(0, result.CancelAttempts); c.Scheduler.Pump(); Assert.AreEqual(0, c.Native.Closes);
        }
        [TestMethod]
        public void ChangedAnsiRepresentationAfterFieldClaimIsRecheckedBeforeMutationEntry()
        {
            var c = new Case(); c.Native.HelpFileCodePage = 65001;
            c.Claim = receipt => { if (!receipt.Terminal && receipt.FieldAttempts == 1) c.Native.HelpFileCodePage = 1252; };
            c.Start(null, @"C:\日本.chm"); var result = c.Complete();
            Assert.IsTrue(result.Terminal && result.Uncertain && c.Native.Visible); Assert.AreEqual(1, result.FieldAttempts);
            Assert.AreEqual(2, c.Native.HelpFileRepresentationChecks); Assert.IsFalse(result.MutationInvoked || result.CommittedRequested);
            Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
            CollectionAssert.AreEqual(new[] { "open", "field", "terminal" }, c.Claims.ToArray());
            c.Scheduler.Pump(); Assert.AreEqual(2, c.Native.HelpFileRepresentationChecks);
        }
        [TestMethod]
        public void PumpedUnownedDialogBeforeFinalExecuteGuardIsNeverAdopted()
        {
            var c = new Case(); c.BeforeEntry = () => { c.Native.Visible = true; c.Scheduler.Pump(); };
            c.Start(); c.Scheduler.Pump(); Assert.AreEqual(0, c.Native.Captures);
            c.Scheduler.Posted(); var r = c.Pending.Result;
            Assert.IsFalse(r.CommandEntered || r.MutationInvoked); Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
        }
        [DataTestMethod]
        [DataRow("version")]
        [DataRow("owner")]
        [DataRow("live")]
        [DataRow("pure")]
        [DataRow("unknownModal")]
        public void ChangedApprovalOrModalBeforeFieldNeverWritesAndRetainsEnteredDialog(string fault)
        {
            var c = new Case(); c.AfterEntry = () =>
            {
                if (fault == "owner") c.Native.OnOwner = () => { throw new InvalidOperationException("owner changed"); };
                if (fault == "live") c.Live = () => { throw new InvalidOperationException("revision/privacy/active changed"); };
                if (fault == "pure") c.Pure = () => { throw new InvalidOperationException("selected scope/policy changed"); };
                if (fault == "unknownModal") c.Native.CaptureError = new InvalidOperationException("unknown modal");
            };
            c.Start(version: fault == "version" ? "stale" : null); c.Scheduler.Posted(); var r = c.Pending.Result;
            Assert.IsTrue(r.Uncertain && r.Terminal); Assert.IsFalse(r.MutationInvoked || r.CommittedRequested);
            Assert.AreEqual(0, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
        }
        [DataTestMethod]
        [DataRow("partialError")]
        [DataRow("wrongValue")]
        [DataRow("otherField")]
        [DataRow("changedHandle")]
        [DataRow("policyBeforeOk")]
        public void PartialWriteOrChangedReadbackNeverOkCancelSaveOrRetry(string fault)
        {
            var c = new Case(); c.Native.OnWrite = () =>
            {
                if (fault == "partialError") c.Native.WriteError = new InvalidOperationException("native failed after partial write");
                if (fault == "wrongValue") c.Native.State.ContextText = "32";
                if (fault == "otherField") c.Native.State.HelpFile = "unexpected.chm";
                if (fault == "changedHandle") c.Native.State.Context = new IntPtr(999);
                if (fault == "policyBeforeOk") c.Live = () => { throw new InvalidOperationException("approval revoked"); };
            };
            c.Start(); c.Scheduler.Posted(); var r = c.Pending.Result;
            Assert.IsTrue(r.MutationInvoked && r.Uncertain); Assert.AreEqual(1, c.Native.Writes); Assert.AreEqual(0, c.Native.Closes);
            Assert.IsFalse(r.CommittedRequested || r.DialogClosed || r.PersistenceVerified); c.Scheduler.Pump(); Assert.AreEqual(1, c.Native.Writes);
        }
        [DataTestMethod]
        [DataRow("open")]
        [DataRow("field")]
        [DataRow("ok")]
        [DataRow("terminal")]
        public void DurableReceiptFailureNeverLosesTerminalResultOrReplaysNativeWork(string phase)
        {
            var c = new Case(); c.Claim = claim => { if (c.Claims[c.Claims.Count - 1] == phase) throw new InvalidOperationException("journal unavailable"); };
            c.Start(); c.Scheduler.Posted(); c.Scheduler.Pump(); var r = c.Pending.Result;
            Assert.IsTrue(r.Terminal); Assert.IsNotNull(r.Error); Assert.IsFalse(c.Pending.IsFaulted);
            Assert.AreEqual(phase == "open" || phase == "field" ? 0 : 1, c.Native.Writes);
            Assert.AreEqual(phase == "terminal" ? 1 : 0, c.Native.Closes); if (r.CommandEntered) Assert.IsTrue(r.Uncertain);
        }
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void PendingOkOrCleanupFailureIsUncertainWithOneWriteAndOneClose(bool cleanup)
        {
            var c = new Case(); if (cleanup) c.Scheduler.DisposeError = new InvalidOperationException("scheduler cleanup"); else c.Native.RefuseClosure = true;
            c.Start(); c.Scheduler.Posted(); if (!cleanup) { c.Scheduler.Time = 20001; c.Scheduler.Pump(); } else c.Scheduler.Pump();
            var r = c.Pending.Result; Assert.IsTrue(r.Terminal && r.Uncertain); Assert.AreEqual(1, c.Native.Writes); Assert.AreEqual(1, c.Native.Closes); Assert.IsFalse(c.Pending.IsFaulted);
        }
    }
}
