using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsNativeTests
    {
        [TestMethod]
        public void CaptureRejectsMissingHostWindowBeforeReadingPanes()
        {
            var native = new FakeNative();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.Capture(false, native));
            Assert.AreEqual(0, native.ListReads.Count);
            Assert.AreEqual(0, native.CallStackReads);
        }

        [TestMethod]
        public void CaptureRoutesVisiblePanesAndOnlyReadsRequestedCallStack()
        {
            var native = new FakeNative { Root = new IntPtr(1) };
            native.Panes["Variables locales"] = new IntPtr(11);
            native.Panes["Espions"] = new IntPtr(12);
            native.Panes["Exécution"] = new IntPtr(13);
            dynamic withoutStack = VbeDebugWindows.Capture(false, native);
            Assert.AreEqual(0, native.CallStackReads);
            Assert.IsNull((object)withoutStack.CallStack);
            CollectionAssert.AreEqual(new[] { new IntPtr(11), new IntPtr(12) }, native.ListReads.ToArray());
            Assert.AreEqual(new IntPtr(13), native.ImmediateHandle);
            Assert.IsTrue((bool)withoutStack.Locals.Available);
            Assert.IsTrue((bool)withoutStack.Watches.Available);

            dynamic withStack = VbeDebugWindows.Capture(true, native);
            Assert.AreEqual(1, native.CallStackReads);
            Assert.AreEqual(new IntPtr(11), native.CallStackHandle);
            Assert.IsTrue((bool)withStack.CallStack.Available);
        }

        [TestMethod]
        public void CaptureDistinguishesMissingPanesFromEmptyDebuggerCollections()
        {
            var native = new FakeNative { Root = new IntPtr(1) };
            dynamic snapshot = VbeDebugWindows.Capture(true, native);
            Assert.IsFalse((bool)snapshot.Locals.Available);
            Assert.IsFalse((bool)snapshot.Watches.Available);
            Assert.IsFalse((bool)snapshot.Immediate.Available);
            Assert.AreEqual(IntPtr.Zero, native.CallStackHandle);
            StringAssert.Contains((string)snapshot.Limits, "missing pane");
        }

        [TestMethod]
        public void ExistingCompileOrOptionsDialogPreventsOpeningAnother()
        {
            var native = new FakeNative { DialogHandle = new IntPtr(2) };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoCompileDialog(native));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoDebugOptionsDialog(native));
            native.DialogHandle = IntPtr.Zero;
            VbeDebugWindows.EnsureNoCompileDialog(native);
            VbeDebugWindows.EnsureNoDebugOptionsDialog(native);
        }

        [TestMethod]
        public void ReadDebugDialogReportsAbsenceAndFiltersInvisibleControls()
        {
            var native = new FakeNative();
            dynamic absent = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsFalse((bool)absent.Available);
            Assert.IsNull((string)absent.Diagnostic);
            native.DialogHandle = new IntPtr(2);
            native.Controls.Add(Control(3, "Static", "Compile error: Syntax error", false));
            native.Controls.Add(Control(4, "Static", "Run-time error '9'", true));
            native.Controls.Add(Control(5, "Button", "End", true));
            native.Controls.Add(Control(6, "Button", "hidden", false));
            dynamic found = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsTrue((bool)found.Available);
            Assert.AreEqual("Run-time error '9'", (string)found.Diagnostic);
            CollectionAssert.AreEqual(new[] { "End" }, (string[])found.Buttons);
            Assert.IsNull((string)found.Error);
        }

        [TestMethod]
        public void ReadDebugDialogRefusesAmbiguousDiagnosticText()
        {
            var native = new FakeNative { DialogHandle = new IntPtr(2) };
            dynamic missing = VbeDebugWindows.ReadDebugDialog(native);
            StringAssert.Contains((string)missing.Error, "found 0");
            native.Controls.Add(Control(3, "Static", "Compile error: one", true));
            native.Controls.Add(Control(4, "Static", "Compile error: two", true));
            dynamic result = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsTrue((bool)result.Available);
            Assert.IsNull((string)result.Diagnostic);
            StringAssert.Contains((string)result.Error, "found 2");
        }

        [TestMethod]
        public void RespondDebugDialogChecksExactMessageAndUniqueButtonBeforeClick()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            var request = new Request { Diagnostic = "Compile error: other", Button = "OK" };
            native.DialogHandle = IntPtr.Zero;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            native.DialogHandle = new IntPtr(2);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            request.Diagnostic = "Compile error: Syntax error";
            request.Button = "Cancel";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            request.Button = "OK";
            native.Controls.Add(Control(5, "Button", "OK", true));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            Assert.AreEqual(0, native.Clicks);
        }

        [TestMethod]
        public void RespondDebugDialogRejectsUnrecognizedDiagnosticsAndFailedClick()
        {
            var native = DialogWith("Generic information", "OK");
            var request = new Request { Diagnostic = "Generic information", Button = "OK" };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            native.Controls[0].Text = "Compile error: Syntax error";
            request.Diagnostic = native.Controls[0].Text;
            native.ClickSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            Assert.AreEqual(1, native.Clicks);
        }

        [TestMethod]
        public void RespondDebugDialogDistinguishesClosedAndPendingAfterClick()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            var request = new Request { Diagnostic = "Compile error: Syntax error", Button = "OK" };
            native.CloseAfterClick = true;
            dynamic closed = VbeDebugWindows.RespondDebugDialog(request, native);
            Assert.AreEqual("DialogClosed", (string)closed.Verification);
            Assert.IsFalse((bool)closed.VerificationPending);
            native.CloseAfterClick = false;
            native.PausesSinceReset = 0;
            dynamic pending = VbeDebugWindows.RespondDebugDialog(request, native);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual(40, native.PausesSinceReset);
        }

        [TestMethod]
        public void AwaitCompileDialogReturnsDiagnosticOnlyAfterOwnOkAndCompletion()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            using (var completed = new ManualResetEventSlim(true))
            {
                Assert.AreEqual("Compile error: Syntax error", VbeDebugWindows.AwaitCompileDialog(completed, native));
                Assert.AreEqual(1, native.Clicks);
            }
        }

        [TestMethod]
        public void AwaitCompileDialogRejectsMissingOkAndTimesOutWithoutDiagnostic()
        {
            var native = DialogWith("Compile error: Syntax error", "Cancel");
            using (var completed = new ManualResetEventSlim(true))
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
            native.Controls[1].Text = "OK";
            native.ClickSucceeds = false;
            using (var completed = new ManualResetEventSlim(true))
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
            native.DialogHandle = IntPtr.Zero;
            using (var completed = new ManualResetEventSlim(true))
                Assert.IsNull(VbeDebugWindows.AwaitCompileDialog(completed, native));
            using (var completed = new ManualResetEventSlim(false))
                Assert.ThrowsException<TimeoutException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
        }

        private static FakeNative DialogWith(string message, string button)
        {
            var native = new FakeNative { DialogHandle = new IntPtr(2), AccessibleMessage = message };
            native.Controls.Add(Control(3, "Static", message, true));
            native.Controls.Add(Control(4, "Button", button, true));
            return native;
        }

        private static VbeDebugWindows.NativeControl Control(int handle, string kind, string text, bool visible)
        {
            return new VbeDebugWindows.NativeControl { Handle = new IntPtr(handle), Kind = kind,
                Text = text, Visible = visible };
        }

        private sealed class FakeNative : VbeDebugWindows.INativeProbe
        {
            public IntPtr Root;
            public IntPtr DialogHandle;
            public readonly Dictionary<string, IntPtr> Panes = new Dictionary<string, IntPtr>();
            public readonly List<VbeDebugWindows.NativeControl> Controls = new List<VbeDebugWindows.NativeControl>();
            public readonly List<IntPtr> ListReads = new List<IntPtr>();
            public IntPtr ImmediateHandle;
            public IntPtr CallStackHandle;
            public int CallStackReads;
            public int Clicks;
            public bool ClickSucceeds = true;
            public bool CloseAfterClick;
            public int PausesSinceReset;
            public string AccessibleMessage;
            public IntPtr VbeRoot() { return Root; }
            public List<IntPtr> Children(IntPtr root) { return Panes.Values.ToList(); }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names)
            {
                foreach (string name in names)
                    if (Panes.TryGetValue(name, out IntPtr handle)) return handle;
                return IntPtr.Zero;
            }
            public object List(IntPtr handle)
            {
                ListReads.Add(handle);
                return new { Available = handle != IntPtr.Zero };
            }
            public object Immediate(IntPtr handle)
            {
                ImmediateHandle = handle;
                return new { Available = handle != IntPtr.Zero };
            }
            public object CallStack(IntPtr locals)
            {
                CallStackHandle = locals;
                CallStackReads++;
                return new { Available = locals != IntPtr.Zero };
            }
            public IntPtr Dialog(params string[] titles) { return DialogHandle; }
            public List<VbeDebugWindows.NativeControl> DialogControls(IntPtr dialog) { return Controls; }
            public string DialogMessage(IntPtr dialog) { return AccessibleMessage; }
            public bool Click(IntPtr handle) { Clicks++; return ClickSucceeds; }
            public bool Visible(IntPtr handle) { return !(CloseAfterClick && Clicks > 0); }
            public void Pause(int milliseconds) { PausesSinceReset++; }
        }
    }
}
