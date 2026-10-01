using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OptionsCancellationTests
    {
        [DataTestMethod, DataRow(3), DataRow(40)]
        public void EmptyNativeSizeCatalogueRefusesBeforeWriteAndWaitsForItsSingleCancel(int closeAfterPauses)
        {
            var probe = new CancellationProbe(true) { CloseAfterPauses = closeAfterPauses };
            var error = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            StringAssert.Contains(error.Message, "exact native choice is absent");
            Assert.AreEqual(0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
            Assert.AreEqual(1, probe.CancelEntries);
            Assert.AreEqual(closeAfterPauses, probe.CancelPauses, "The original refusal must not return while its queued Cancel is pending.");
            Assert.IsFalse(probe.Inner.Open);
            Assert.AreEqual("10", probe.Inner.Items[0].Value);
        }

        [DataTestMethod, DataRow("stale"), DataRow("absent"), DataRow("capture"), DataRow("readback")]
        public void KnownFailuresCancelOnceAndVerifyClosureWithoutCommit(string phase)
        {
            var probe = new CancellationProbe();
            var primary = new IOException("synthetic primary " + phase);
            if (phase == "stale") probe.Request.ExpectedOptionsVersion = "stale";
            if (phase == "absent") probe.Request.Property = "Unknown";
            if (phase == "capture") probe.CaptureError = primary;
            if (phase == "readback") probe.ReadbackError = primary;
            Exception observed;
            if (phase == "capture" || phase == "readback")
                observed = Assert.ThrowsException<IOException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            else observed = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            if (phase == "capture" || phase == "readback") Assert.AreSame(primary, observed);
            Assert.AreEqual(phase == "readback" ? 1 : 0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
            Assert.AreEqual(1, probe.CancelEntries);
            Assert.IsFalse(probe.Inner.Open);
        }

        [DataTestMethod, DataRow("cancel"), DataRow("lifetime"), DataRow("timeout")]
        public void PrimaryAndCancellationFailuresRemainDistinctWithNoSecondCancel(string phase)
        {
            var probe = new CancellationProbe();
            var primary = new IOException("synthetic inspection failure");
            var cleanup = new InvalidOperationException("synthetic cancellation failure");
            probe.CaptureError = primary;
            if (phase == "cancel") probe.CancelError = cleanup;
            if (phase == "lifetime") probe.LifetimeError = cleanup;
            if (phase == "timeout") probe.CloseAfterPauses = int.MaxValue;
            var error = Assert.ThrowsException<AggregateException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            Assert.AreEqual(2, error.InnerExceptions.Count);
            Assert.AreSame(primary, error.InnerExceptions[0]);
            if (phase != "timeout") Assert.AreSame(cleanup, error.InnerExceptions[1]);
            else { StringAssert.Contains(error.InnerExceptions[1].Message, "bounded observation period"); Assert.AreEqual(40, probe.CancelPauses); }
            Assert.AreEqual(1, probe.CancelEntries);
            Assert.AreEqual(0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
            Assert.IsTrue(probe.Inner.Open);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ReplacedOrHiddenCapturedDialogRefusesCancellationWithoutTouchingAnotherDialog(bool hidden)
        {
            var probe = new CancellationProbe(true) { HiddenDuringCleanup = hidden, ChangedDuringCleanup = !hidden };
            var error = Assert.ThrowsException<AggregateException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            StringAssert.Contains(error.InnerExceptions[0].Message, "exact native choice is absent");
            Assert.AreEqual(0, probe.CancelEntries);
            Assert.AreEqual(0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
            Assert.IsTrue(probe.Inner.Open);
        }

        [TestMethod]
        public void DifferentDialogAfterSingleCancelIsNeverCancelledByFallbackVerification()
        {
            var probe = new CancellationProbe(true) { ReplacementAfterCancel = true, CloseAfterPauses = int.MaxValue };
            var error = Assert.ThrowsException<AggregateException>(() => VbeDebugWindows.SetVbeOption(probe.Request, new WithoutLifetimeProbe(probe)));
            StringAssert.Contains(error.InnerExceptions[0].Message, "exact native choice is absent");
            StringAssert.Contains(error.InnerExceptions[1].Message, "identity changed while verifying Cancel");
            Assert.AreEqual(1, probe.CancelEntries);
            Assert.AreEqual(0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
            Assert.IsTrue(probe.ReplacementAfterCancel);
        }

        [TestMethod]
        public void ExactCapturedClosureDoesNotCancelANewlyOpenedDifferentDialog()
        {
            var probe = new CancellationProbe(true) { ReplacementAfterCancel = true };
            var error = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            StringAssert.Contains(error.Message, "exact native choice is absent");
            Assert.IsFalse(probe.Inner.Open);
            Assert.AreEqual(new IntPtr(72), probe.Dialog());
            Assert.AreEqual(1, probe.CancelEntries);
            Assert.AreEqual(0, probe.Inner.Writes);
            Assert.AreEqual(0, probe.AcceptEntries);
        }

        [DataTestMethod, DataRow("write"), DataRow("accept"), DataRow("commit observation")]
        public void UncertainWriteOrCommitRetainsDialogWithoutCancellationOrReplay(string phase)
        {
            var probe = new CancellationProbe();
            var primary = new IOException("synthetic uncertain " + phase);
            if (phase == "write") probe.WriteError = primary;
            if (phase == "accept") probe.AcceptError = primary;
            if (phase == "commit observation") { probe.KeepCommittedDialog = true; probe.CommitObservationError = primary; }
            var observed = Assert.ThrowsException<IOException>(() => VbeDebugWindows.SetVbeOption(probe.Request, probe));
            Assert.AreSame(primary, observed);
            Assert.AreEqual(1, probe.Inner.Writes);
            Assert.AreEqual(phase == "write" ? 0 : 1, probe.AcceptEntries);
            Assert.AreEqual(0, probe.CancelEntries);
            Assert.IsTrue(probe.Inner.Open);
        }

        [TestMethod]
        public void PendingSuccessfulAcceptReturnsUnclosedWithoutCancel()
        {
            var probe = new CancellationProbe { KeepCommittedDialog = true };
            dynamic result = VbeDebugWindows.SetVbeOption(probe.Request, probe);
            Assert.IsTrue((bool)result.CommitRequested);
            Assert.IsFalse((bool)result.DialogClosed);
            Assert.AreEqual(1, probe.Inner.Writes);
            Assert.AreEqual(1, probe.AcceptEntries);
            Assert.AreEqual(0, probe.CancelEntries);
            Assert.IsTrue(probe.Inner.Open);
        }

        private sealed class CancellationProbe : VbeDebugWindows.IWritableOptionsProbe, VbeDebugWindows.IOptionsDialogLifetimeProbe
        {
            internal readonly WritableOptionsMatrixProbe Inner = new WritableOptionsMatrixProbe();
            internal readonly Request Request;
            internal int CancelEntries, CancelPauses, AcceptEntries, DialogReads;
            internal int CloseAfterPauses;
            internal Exception CaptureError, ReadbackError, CancelError, LifetimeError, WriteError, AcceptError, CommitObservationError;
            internal bool ChangedDuringCleanup, HiddenDuringCleanup, KeepCommittedDialog, ReplacementAfterCancel;
            internal CancellationProbe(bool emptySize = false)
            {
                if (emptySize)
                {
                    Inner.Names[0] = "Format de l'éditeur";
                    Inner.Items[0] = new VbeDebugWindows.OptionsControl { Name = "Taille :", Type = "ControlType.ComboBox", Value = "10", Choices = new string[0] };
                }
                Request = Inner.Request();
                if (emptySize) Request.Value = "12";
            }
            public IntPtr Dialog()
            {
                DialogReads++;
                if (AcceptEntries != 0 && CommitObservationError != null) throw CommitObservationError;
                if (CancelEntries != 0 && ReplacementAfterCancel) return new IntPtr(72);
                if (DialogReads > 1 && HiddenDuringCleanup) return IntPtr.Zero;
                if (DialogReads > 1 && ChangedDuringCleanup) return new IntPtr(72);
                return Inner.Open ? new IntPtr(71) : IntPtr.Zero;
            }
            public IList<string> Tabs(IntPtr dialog) => Inner.Names;
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int index)
            {
                if (CaptureError != null) throw CaptureError;
                if (Inner.Writes != 0 && ReadbackError != null) throw ReadbackError;
                return Inner.Items;
            }
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Write(IntPtr dialog, int index, string name, string type, object value)
            {
                Inner.Write(dialog, index, name, type, value);
                if (WriteError != null) throw WriteError;
            }
            public void Accept(IntPtr dialog)
            {
                AcceptEntries++;
                if (AcceptError != null) throw AcceptError;
                if (!KeepCommittedDialog) Inner.Open = false;
            }
            public void Close(IntPtr dialog)
            {
                Assert.AreEqual(new IntPtr(71), dialog);
                CancelEntries++;
                if (CancelError != null) throw CancelError;
                if (CloseAfterPauses == 0 && LifetimeError == null) Inner.Open = false;
            }
            public bool IsOpen(IntPtr dialog)
            {
                Assert.AreEqual(new IntPtr(71), dialog);
                if (LifetimeError != null) throw LifetimeError;
                return Inner.Open;
            }
            public void Pause(int milliseconds)
            {
                if (CancelEntries != 0 && ++CancelPauses >= CloseAfterPauses) Inner.Open = false;
            }
        }

        private sealed class WithoutLifetimeProbe : VbeDebugWindows.IWritableOptionsProbe
        {
            private readonly CancellationProbe inner;
            internal WithoutLifetimeProbe(CancellationProbe inner) { this.inner = inner; }
            public IntPtr Dialog() => inner.Dialog();
            public IList<string> Tabs(IntPtr dialog) => inner.Tabs(dialog);
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int index) => inner.Controls(dialog, index);
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => inner.ErrorChoices(dialog);
            public void Write(IntPtr dialog, int index, string name, string type, object value) => inner.Write(dialog, index, name, type, value);
            public void Accept(IntPtr dialog) => inner.Accept(dialog);
            public void Close(IntPtr dialog) => inner.Close(dialog);
            public void Pause(int milliseconds) => inner.Pause(milliseconds);
        }
    }

    public sealed partial class VbeDebugWindowsSystemTests
    {
        [DataTestMethod]
        [DataRow("zero")][DataRow("foreign dialog")][DataRow("wrong dialog class")][DataRow("changed dialog")]
        [DataRow("missing button")][DataRow("foreign button")][DataRow("wrong button class")][DataRow("disabled")][DataRow("post failure")]
        [DataRow("zero dialog thread")][DataRow("wrong button thread")][DataRow("zero button thread")][DataRow("foreign cancel lookup")]
        public void NativeOptionsCancelRejectsInvalidIdentityOrButtonWithoutAlternateAction(string scenario)
        {
            var enabled = VbeDebugWindows.OptionsWindowEnabled;
            try
            {
                using (var scene = new SystemScene())
                {
                    var native = Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");
                    var dialog = scene.Add("Options");
                    var button = scene.Add("Annuler", "Button", dialog, 2);
                    VbeDebugWindows.OptionsWindowEnabled = _ => scenario != "disabled";
                    if (scenario == "foreign dialog") dialog.ProcessId++;
                    if (scenario == "zero dialog thread") dialog.ThreadId = 0;
                    if (scenario == "wrong dialog class") dialog.Class = "other";
                    if (scenario == "changed dialog") { var other = scene.Add("Options"); scene.Windows.Remove(other); scene.Windows.Insert(0, other); }
                    if (scenario == "missing button") scene.Windows.Remove(button);
                    if (scenario == "foreign button") button.ProcessId++;
                    if (scenario == "wrong button thread") button.ThreadId++;
                    if (scenario == "zero button thread") button.ThreadId = 0;
                    if (scenario == "wrong button class") button.Class = "Edit";
                    if (scenario == "post failure") scene.PostSucceeds = false;
                    if (scenario == "foreign cancel lookup")
                    {
                        var foreign = scene.Add("Other window"); foreign.ProcessId++;
                        var foreignCancel = scene.Add("Cancel", "Button", foreign, 2);
                        VbeDebugWindows.GetDlgItem = (handle, id) => foreignCancel.Handle;
                    }
                    Assert.ThrowsException<InvalidOperationException>(() => native.Close(scenario == "zero" ? IntPtr.Zero : dialog.Handle));
                    Assert.AreEqual(scenario == "post failure" ? 1 : 0, scene.Messages.Count);
                }
            }
            finally { VbeDebugWindows.OptionsWindowEnabled = enabled; }
        }

        [TestMethod]
        public void NativeOptionsCancelPostsExactlyOneOwnedCancelButtonClick()
        {
            var enabled = VbeDebugWindows.OptionsWindowEnabled;
            try
            {
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Options");
                    var button = scene.Add("Annuler", "Button", dialog, 2);
                    VbeDebugWindows.OptionsWindowEnabled = _ => true;
                    Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe").Close(dialog.Handle);
                    Assert.AreEqual(1, scene.Messages.Count);
                    Assert.AreEqual(button.Handle, scene.Messages[0].Item1);
                    Assert.AreEqual(0xF5, scene.Messages[0].Item2);
                }
            }
            finally { VbeDebugWindows.OptionsWindowEnabled = enabled; }
        }

        [TestMethod]
        public void NativeOptionsCapturedLifetimeObservesHiddenWindowAndRefusesChangedOwnerOrFailedEnumeration()
        {
            using (var scene = new SystemScene())
            {
                var native = Native<VbeDebugWindows.IOptionsDialogLifetimeProbe>("NativeOptionsProbe");
                var dialog = scene.Add("Options");
                dialog.Visible = false;
                Assert.IsTrue(native.IsOpen(dialog.Handle), "Hidden is not destroyed.");
                dialog.ProcessId++;
                Assert.ThrowsException<InvalidOperationException>(() => native.IsOpen(dialog.Handle));
                dialog.ProcessId--;
                scene.Windows.Remove(dialog);
                Assert.IsFalse(native.IsOpen(dialog.Handle));
                VbeDebugWindows.EnumWindows = (callback, argument) => false;
                Assert.ThrowsException<InvalidOperationException>(() => native.IsOpen(dialog.Handle));
                Assert.AreEqual(0, scene.Messages.Count);
            }
        }

        [TestMethod]
        public void PreexistingOptionsDialogRefusalNeverPostsCancel()
        {
            using (var scene = new SystemScene())
            {
                scene.Add("Options");
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoDebugOptionsDialog());
                Assert.AreEqual(0, scene.Messages.Count);
            }
        }

        [TestMethod]
        public void OptionsSceneModelsUnknownHandlesAndInheritedOwnershipAndRestoresEnabledBoundary()
        {
            var previous = VbeDebugWindows.OptionsWindowEnabled;
            using (var scene = new SystemScene())
            {
                Assert.IsFalse(VbeDebugWindows.OptionsWindowEnabled(new IntPtr(999)));
                Assert.AreEqual(0u, VbeDebugWindows.GetWindowThreadProcessId(new IntPtr(999), out uint absentPid));
                Assert.AreEqual(0u, absentPid);
                var dialog = scene.Add("Options"); dialog.ThreadId = 123;
                var cancel = scene.Add("Cancel", "Button", dialog, 2);
                Assert.AreEqual(dialog.ThreadId, cancel.ThreadId);
                Assert.AreEqual(dialog.ProcessId, cancel.ProcessId);
                Assert.IsTrue(VbeDebugWindows.OptionsWindowEnabled(cancel.Handle));
                cancel.Enabled = false;
                Assert.IsFalse(VbeDebugWindows.OptionsWindowEnabled(cancel.Handle));
            }
            Assert.AreSame(previous, VbeDebugWindows.OptionsWindowEnabled);
        }
    }
}
