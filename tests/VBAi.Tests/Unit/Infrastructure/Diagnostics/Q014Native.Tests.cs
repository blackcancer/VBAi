using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Linq;
using Microsoft.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises admission without COM or a running host, and checks the helper's PowerShell compiler boundary.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SolidWorksFileDialogAdmissionTests
    {
        [TestMethod]
        public void QualificationHelperCompilesWithTheWindowsPowerShellLanguageBoundary()
        {
            var directory = new DirectoryInfo(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "tools", "tests", "Q014Native.cs")))
                directory = directory.Parent;
            Assert.IsNotNull(directory, "The checked-out qualification source is required for this compiler contract.");
            using (var provider = new CSharpCodeProvider())
            {
                var parameters = new CompilerParameters
                {
                    GenerateExecutable = false,
                    GenerateInMemory = true,
                    CompilerOptions = "/langversion:5"
                };
                parameters.ReferencedAssemblies.Add("System.dll");
                parameters.ReferencedAssemblies.Add("System.Core.dll");
                var result = provider.CompileAssemblyFromFile(parameters, Path.Combine(directory.FullName, "tools", "tests", "Q014Native.cs"));
                Assert.IsFalse(result.Errors.HasErrors, string.Join(Environment.NewLine,
                    result.Errors.Cast<CompilerError>().Select(error => error.ToString())));
                Assert.IsNotNull(result.CompiledAssembly.GetType("Q014Native+VbeCloseAttempt", true),
                    "The native controller must compile as an actual type without invoking a Windows or host action.");
            }
        }

        private static Q014Native.Window CloseRoot() => new Q014Native.Window
        { Handle = 100, Pid = 42, Thread = 73, Class = "wndclass_desked_gsk", Caption = "VBE disposable macro", Visible = true, Enabled = true };

        [DataTestMethod]
        [DataRow("null")][DataRow("handle")][DataRow("negative handle")][DataRow("pid")][DataRow("large pid")]
        [DataRow("thread")][DataRow("class")][DataRow("caption")][DataRow("hidden")][DataRow("disabled")]
        public void OrdinaryVbeCloseRequiresAnExplicitUsableRoot(string fault)
        {
            var root = CloseRoot();
            if (fault == "null") root = null;
            if (fault == "handle") root.Handle = 0;
            if (fault == "negative handle") root.Handle = -1;
            if (fault == "pid") root.Pid = 0;
            if (fault == "large pid") root.Pid = uint.MaxValue;
            if (fault == "thread") root.Thread = 0;
            if (fault == "class") root.Class = "Another window";
            if (fault == "caption") root.Caption = " ";
            if (fault == "hidden") root.Visible = false;
            if (fault == "disabled") root.Enabled = false;
            Assert.ThrowsException<ArgumentException>(() => new Q014Native.VbeCloseAttempt(root, () => { }));
        }

        [DataTestMethod]
        [DataRow("null")][DataRow("empty")][DataRow("large")][DataRow("null row")][DataRow("handle")]
        [DataRow("foreign pid")][DataRow("thread")][DataRow("class")][DataRow("duplicate")][DataRow("missing root")]
        [DataRow("root thread")][DataRow("root class")][DataRow("caption")][DataRow("hidden")][DataRow("disabled")][DataRow("modal")]
        public void InvalidInventoryAtEitherReadNeverDeliversOrRetriesClose(string fault)
        {
            foreach (int failedRead in new[] { 1, 2 })
            {
                var row = CloseRoot(); var inventory = new[] { row };
                if (fault == "null") inventory = null;
                if (fault == "empty") inventory = new Q014Native.Window[0];
                if (fault == "large") inventory = new Q014Native.Window[8193];
                if (fault == "null row") inventory = new Q014Native.Window[] { null };
                if (fault == "handle") row.Handle = 0;
                if (fault == "foreign pid") row.Pid = 43;
                if (fault == "thread") row.Thread = 0;
                if (fault == "class") row.Class = " ";
                if (fault == "duplicate") inventory = new[] { row, row };
                if (fault == "missing root") row.Handle = 101;
                if (fault == "root thread") row.Thread = 74;
                if (fault == "root class") row.Class = "Other root";
                if (fault == "caption") row.Caption = "Another macro";
                if (fault == "hidden") row.Visible = false;
                if (fault == "disabled") row.Enabled = false;
                if (fault == "modal") inventory = new[] { row, new Q014Native.Window { Handle = 200, Pid = 42, Thread = 73, Class = "#32770", Visible = true } };
                int reads = 0, posts = 0;
                var attempt = new Q014Native.VbeCloseAttempt(CloseRoot(), () => ++reads == failedRead ? inventory : new[] { CloseRoot() }, () => { }, _ => { posts++; return true; });
                Assert.ThrowsException<InvalidOperationException>(attempt.Request);
                Assert.AreEqual(failedRead, reads); Assert.AreEqual(0, posts);
                Assert.ThrowsException<InvalidOperationException>(attempt.Request);
                Assert.AreEqual(failedRead, reads); Assert.AreEqual(0, posts);
            }
        }

        [DataTestMethod]
        [DataRow(1)][DataRow(2)][DataRow(3)]
        public void LostDesktopContextAtAnyBoundaryPreventsClose(int boundary)
        {
            int checks = 0, posts = 0;
            var attempt = new Q014Native.VbeCloseAttempt(CloseRoot(), () => new[] { CloseRoot() },
                () => { if (++checks == boundary) throw new InvalidOperationException("Desktop changed"); }, _ => { posts++; return true; });
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(0, posts);
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(boundary, checks);
        }

        [DataTestMethod]
        [DataRow(false)][DataRow(true)]
        public void FailedOrThrowingDeliveryIsNotRetried(bool throws)
        {
            int posts = 0;
            var attempt = new Q014Native.VbeCloseAttempt(CloseRoot(), () => new[] { CloseRoot() }, () => { }, _ =>
            { posts++; if (throws) throw new InvalidOperationException("Native failure"); return false; });
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(1, posts);
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(1, posts);
        }

        [TestMethod]
        public void CloseUsesItsCapturedIdentityAndOneDeliveryWithoutAssertingCompletion()
        {
            var root = CloseRoot(); int posts = 0, reads = 0, checks = 0;
            var attempt = new Q014Native.VbeCloseAttempt(root, () => { reads++; return new[] { CloseRoot() }; }, () => checks++, handle =>
            { Assert.AreEqual(100L, handle); posts++; return true; });
            root.Handle = 900; root.Caption = "Caller changed the snapshot";
            attempt.Request(); Assert.AreEqual(2, reads); Assert.AreEqual(3, checks); Assert.AreEqual(1, posts);
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(1, posts);
        }

        [TestMethod]
        public void CloseRejectsAbsentBoundariesAndPropagatesReadFailureWithoutDelivery()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new Q014Native.VbeCloseAttempt(CloseRoot(), null, () => { }, _ => true));
            Assert.ThrowsException<ArgumentNullException>(() => new Q014Native.VbeCloseAttempt(CloseRoot(), () => new[] { CloseRoot() }, null, _ => true));
            Assert.ThrowsException<ArgumentNullException>(() => new Q014Native.VbeCloseAttempt(CloseRoot(), () => new[] { CloseRoot() }, () => { }, null));
            int posts = 0;
            var attempt = new Q014Native.VbeCloseAttempt(CloseRoot(), () => { throw new InvalidOperationException("Inventory unavailable"); }, () => { }, _ => { posts++; return true; });
            Assert.ThrowsException<InvalidOperationException>(attempt.Request); Assert.AreEqual(0, posts);
        }

        private static Q014Native.FileDialogSnapshot Snapshot(string caption = "&Enregistrer")
        {
            return new Q014Native.FileDialogSnapshot
            {
                Dialog = new Q014Native.Window { Handle = 100, Pid = 42, Thread = 73, Class = "#32770", Caption = "Enregistrer sous", Visible = true },
                Filename = 101, Action = 102, ActionCaption = caption
            };
        }

        [DataTestMethod]
        [DataRow("Save", "Save")]
        [DataRow("Save", "&Save")]
        [DataRow("Save", "Enregistrer")]
        [DataRow("Save", "&Enregistrer")]
        [DataRow("Open", "Open")]
        [DataRow("Open", "&Open")]
        [DataRow("Open", "Ouvrir")]
        [DataRow("Open", "&Ouvrir")]
        public void TwoIndependentReadyReadsAdmitTheDeclaredAction(string action, string caption)
        {
            Assert.IsTrue(Q014Native.IsStableFileDialog(Snapshot(caption), Snapshot(caption), 42, action));
        }

        [DataTestMethod]
        [DataRow("Save", "Cancel")]
        [DataRow("Save", "Ouvrir")]
        [DataRow("Open", "Enregistrer")]
        [DataRow("Open", "open")]
        [DataRow("Save", "Enregistrer ")]
        [DataRow("Save", "")]
        [DataRow("Save", null)]
        public void UnexpectedActionCannotAuthorizeDelivery(string action, string caption)
        {
            Assert.IsFalse(Q014Native.IsStableFileDialog(Snapshot(caption), Snapshot(caption), 42, action));
        }

        [DataTestMethod]
        [DataRow("snapshot")]
        [DataRow("dialog")]
        [DataRow("handle")]
        [DataRow("negativeHandle")]
        [DataRow("pid")]
        [DataRow("thread")]
        [DataRow("class")]
        [DataRow("hidden")]
        [DataRow("filename")]
        [DataRow("negativeFilename")]
        [DataRow("action")]
        [DataRow("negativeAction")]
        [DataRow("sameControls")]
        [DataRow("filenameIsDialog")]
        [DataRow("actionIsDialog")]
        public void IncompleteOrForeignIdentityRejectsEitherRead(string fault)
        {
            var invalid = Snapshot();
            switch (fault)
            {
                case "snapshot": invalid = null; break;
                case "dialog": invalid.Dialog = null; break;
                case "handle": invalid.Dialog.Handle = 0; break;
                case "negativeHandle": invalid.Dialog.Handle = -1; break;
                case "pid": invalid.Dialog.Pid = 43; break;
                case "thread": invalid.Dialog.Thread = 0; break;
                case "class": invalid.Dialog.Class = "Button"; break;
                case "hidden": invalid.Dialog.Visible = false; break;
                case "filename": invalid.Filename = 0; break;
                case "negativeFilename": invalid.Filename = -1; break;
                case "action": invalid.Action = 0; break;
                case "negativeAction": invalid.Action = -1; break;
                case "sameControls": invalid.Filename = invalid.Action; break;
                case "filenameIsDialog": invalid.Filename = invalid.Dialog.Handle; break;
                case "actionIsDialog": invalid.Action = invalid.Dialog.Handle; break;
                default: Assert.Fail("Unknown fixture fault."); break;
            }
            Assert.IsFalse(Q014Native.IsStableFileDialog(invalid, Snapshot(), 42, "Save"));
            Assert.IsFalse(Q014Native.IsStableFileDialog(Snapshot(), invalid, 42, "Save"));
        }

        [DataTestMethod]
        [DataRow("dialog")]
        [DataRow("thread")]
        [DataRow("caption")]
        [DataRow("filename")]
        [DataRow("action")]
        [DataRow("actionCaption")]
        public void AReplacementRequiresItsOwnConsecutiveAdmissionAndNeverMatchesTheOriginal(string change)
        {
            var original = Snapshot(); var replacement = Snapshot(); var next = Snapshot();
            switch (change)
            {
                case "dialog": replacement.Dialog.Handle = next.Dialog.Handle = 200; break;
                case "thread": replacement.Dialog.Thread = next.Dialog.Thread = 74; break;
                case "caption": replacement.Dialog.Caption = next.Dialog.Caption = "Autre boîte"; break;
                case "filename": replacement.Filename = next.Filename = 201; break;
                case "action": replacement.Action = next.Action = 202; break;
                case "actionCaption": replacement.ActionCaption = next.ActionCaption = "Enregistrer"; break;
                default: Assert.Fail("Unknown fixture change."); break;
            }
            Assert.IsFalse(Q014Native.IsStableFileDialog(original, replacement, 42, "Save"));
            Assert.IsTrue(Q014Native.IsStableFileDialog(replacement, next, 42, "Save"));
            Assert.IsFalse(Q014Native.IsStableFileDialog(original, next, 42, "Save"));
        }

        [TestMethod]
        public void MissingReadBreaksConsecutiveAdmission()
        {
            Assert.IsFalse(Q014Native.IsStableFileDialog(Snapshot(), null, 42, "Save"));
            Assert.IsFalse(Q014Native.IsStableFileDialog(null, Snapshot(), 42, "Save"));
            Assert.IsFalse(Q014Native.IsStableFileDialog(null, null, 42, "Save"));
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(-1)]
        public void InvalidExpectedOwnerFailsBeforeReadingCandidates(int pid)
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => Q014Native.IsStableFileDialog(Snapshot(), Snapshot(), pid, "Save"));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => Q014Native.ObserveFileDialog(pid));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("Cancel")]
        [DataRow("save")]
        public void UnknownRequestedActionIsRejected(string action)
        {
            Assert.ThrowsException<ArgumentException>(() => Q014Native.IsStableFileDialog(Snapshot(), Snapshot(), 42, action));
        }

        [TestMethod]
        public void ZeroOrMultipleVisibleDialogsNeverAuthorizeAdmission()
        {
            var first = Snapshot().Dialog; var second = Snapshot().Dialog; second.Handle = 200;
            Assert.IsNull(Q014Native.UniqueFileDialogCandidate(new Q014Native.Window[0], 42));
            Assert.IsNull(Q014Native.UniqueFileDialogCandidate(new[] { first, second }, 42));
            Assert.IsFalse(Q014Native.IsStableFileDialog(Snapshot(), null, 42, "Save"));
            Assert.AreSame(second, Q014Native.UniqueFileDialogCandidate(new[] { second }, 42));
        }

        [DataTestMethod]
        [DataRow("null")]
        [DataRow("foreign")]
        [DataRow("hidden")]
        [DataRow("otherClass")]
        public void UnrelatedWindowsDoNotMakeTheOwnedVisibleDialogAmbiguous(string kind)
        {
            var candidate = Snapshot().Dialog; var ignored = Snapshot().Dialog;
            switch (kind)
            {
                case "null": ignored = null; break;
                case "foreign": ignored.Pid = 43; break;
                case "hidden": ignored.Visible = false; break;
                case "otherClass": ignored.Class = "Afx:"; break;
            }
            Assert.AreSame(candidate, Q014Native.UniqueFileDialogCandidate(new[] { ignored, candidate }, 42));
            Assert.IsNull(Q014Native.UniqueFileDialogCandidate(new[] { ignored }, 42));
        }

        [DataTestMethod]
        [DataRow("valid")][DataRow("modal")][DataRow("unknown")][DataRow("provider error")]
        [DataRow("null dialog")][DataRow("null frame")][DataRow("null inventory")][DataRow("empty")][DataRow("large")][DataRow("null provider")]
        [DataRow("dialog handle")][DataRow("same handle")][DataRow("dialog pid")][DataRow("dialog thread")][DataRow("foreign dialog")]
        [DataRow("other dialog thread")][DataRow("dialog class")][DataRow("dialog caption")][DataRow("dialog hidden")][DataRow("dialog disabled")]
        [DataRow("frame class")][DataRow("frame hidden")][DataRow("frame disabled")][DataRow("null row")][DataRow("duplicate")]
        [DataRow("foreign row")][DataRow("invalid row")][DataRow("missing dialog")][DataRow("missing frame")][DataRow("changed dialog")][DataRow("changed frame")]
        [DataRow("missing welcome")][DataRow("welcome thread")][DataRow("welcome disabled")][DataRow("welcome hidden")]
        [DataRow("second welcome")][DataRow("second dialog")][DataRow("second frame")]
        public void OnlyExactOwnedModelessWelcomeCanPassFileCommandAdmission(string fault)
        {
            var dialog = new Q014Native.Window { Handle = 200, Pid = 42, Thread = 73, Class = "#32770", Caption = "", Visible = true, Enabled = true };
            var frame = new Q014Native.Window { Handle = 100, Pid = 42, Thread = 73, Class = "Afx:owned", Caption = "SOLIDWORKS test", Visible = true, Enabled = true };
            var welcome = new Q014Native.Window { Handle = 300, Pid = 42, Thread = 73, Class = "swWelcomeScreen", Caption = "", Visible = true, Enabled = true };
            var inventory = new[] { dialog, frame, welcome }; int reads = 0;
            Func<long, bool?> provider = handle => { reads++; Assert.AreEqual(200L, handle); return false; };
            switch (fault)
            {
                case "modal": provider = _ => { reads++; return true; }; break;
                case "unknown": provider = _ => { reads++; return null; }; break;
                case "provider error": provider = _ => { reads++; throw new InvalidOperationException("Unavailable modality"); }; break;
                case "null dialog": dialog = null; break;
                case "null frame": frame = null; break;
                case "null inventory": inventory = null; break;
                case "empty": inventory = new Q014Native.Window[0]; break;
                case "large": inventory = new Q014Native.Window[8193]; break;
                case "null provider": provider = null; break;
                case "dialog handle": dialog.Handle = 0; break;
                case "same handle": dialog.Handle = frame.Handle; break;
                case "dialog pid": dialog.Pid = 0; break;
                case "dialog thread": dialog.Thread = 0; break;
                case "foreign dialog": dialog.Pid++; break;
                case "other dialog thread": dialog.Thread++; break;
                case "dialog class": dialog.Class = "Other"; break;
                case "dialog caption": dialog.Caption = "Unexpected prompt"; break;
                case "dialog hidden": dialog.Visible = false; break;
                case "dialog disabled": dialog.Enabled = false; break;
                case "frame class": frame.Class = null; break;
                case "frame hidden": frame.Visible = false; break;
                case "frame disabled": frame.Enabled = false; break;
                case "null row": inventory = new[] { dialog, frame, welcome, null }; break;
                case "duplicate": inventory = new[] { dialog, frame, welcome, welcome }; break;
                case "foreign row": welcome.Pid++; break;
                case "invalid row": welcome.Class = " "; break;
                case "missing dialog": inventory = new[] { frame, welcome }; break;
                case "missing frame": inventory = new[] { dialog, welcome }; break;
                case "changed dialog": inventory[0] = new Q014Native.Window { Handle = 200, Pid = 42, Thread = 73, Class = "#32770", Caption = "Changed", Visible = true, Enabled = true }; break;
                case "changed frame": inventory[1] = new Q014Native.Window { Handle = 100, Pid = 42, Thread = 73, Class = "Afx:owned", Caption = "Changed", Visible = true, Enabled = true }; break;
                case "missing welcome": inventory = new[] { dialog, frame }; break;
                case "welcome thread": welcome.Thread++; break;
                case "welcome disabled": welcome.Enabled = false; break;
                case "welcome hidden": welcome.Visible = false; break;
                case "second welcome": inventory = new[] { dialog, frame, welcome, new Q014Native.Window { Handle = 301, Pid = 42, Thread = 73, Class = "swWelcomeScreen", Visible = true, Enabled = true } }; break;
                case "second dialog": inventory = new[] { dialog, frame, welcome, new Q014Native.Window { Handle = 201, Pid = 42, Thread = 73, Class = "#32770", Visible = true, Enabled = true } }; break;
                case "second frame": inventory = new[] { dialog, frame, welcome, new Q014Native.Window { Handle = 101, Pid = 42, Thread = 73, Class = frame.Class, Caption = frame.Caption, Visible = true, Enabled = true } }; break;
            }
            if (fault == "provider error")
                Assert.ThrowsException<InvalidOperationException>(() => Q014Native.IsModelessWelcome(dialog, frame, inventory, provider));
            else Assert.AreEqual(fault == "valid", Q014Native.IsModelessWelcome(dialog, frame, inventory, provider));
            Assert.AreEqual(new[] { "valid", "modal", "unknown", "provider error" }.Contains(fault) ? 1 : 0, reads);
        }

        [TestMethod]
        public void CandidateInventoryRequiresAnExplicitOwnerAndObservedArray()
        {
            Assert.ThrowsException<ArgumentNullException>(() => Q014Native.UniqueFileDialogCandidate(null, 42));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => Q014Native.UniqueFileDialogCandidate(new Q014Native.Window[0], 0));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => Q014Native.UniqueFileDialogCandidate(new Q014Native.Window[0], -1));
        }
    }
}
