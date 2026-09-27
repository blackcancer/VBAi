using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsWatchProbeTests
    {
        private static Request Add() { return new Request { Project = "Book", Module = "Module1",
            Procedure = "Run", Expression = "x", WatchType = "break_when_true" }; }
        private static Request Edit() { return new Request { Project = "Book", Module = "Module1",
            Expression = "x", NewExpression = "y", Context = "Module1.Run",
            WatchType = "break_when_changed" }; }

        [TestMethod]
        public void AddWatchRequiresDialogAndExactContext()
        {
            var fake = new WatchFake { Open = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            fake.Texts[4858] = "Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(1, fake.Closes);
            fake.Open = true;
            fake.Texts[4858] = "Book";
            fake.Texts[4856] = "Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
        }

        [TestMethod]
        public void AddWatchChecksControlsTypeEditAndOk()
        {
            var fake = new WatchFake();
            fake.Missing = 4853;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Missing = 4851;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Missing = 0;
            fake.Check = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Check = true;
            fake.ReplaceEcho = false;
            fake.Texts[4853] = "";
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.ReplaceEcho = true;
            fake.RefuseClick = 1;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
        }

        [TestMethod]
        public void AddWatchReportsNativeRejectionTimeoutAndPendingOrVisibleReadback()
        {
            var fake = new WatchFake { ErrorOpen = true, CloseAfterOk = false };
            var rejection = Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            StringAssert.Contains(rejection.Message, "invalid expression");
            Assert.AreEqual(2, fake.Closes);

            fake = new WatchFake { CloseAfterOk = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(1, fake.Closes);

            fake = new WatchFake();
            dynamic pending = VbeDebugWindows.CompleteAddWatch(Add(), fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.AreEqual("break_when_true", (string)pending.WatchType);
            fake = new WatchFake { Root = new IntPtr(4), PaneHandle = new IntPtr(5) };
            dynamic visible = VbeDebugWindows.CompleteAddWatch(Add(), fake);
            Assert.AreEqual("ReadbackAvailable", (string)visible.Verification);
            Assert.AreEqual(new IntPtr(5), fake.ListHandle);
        }

        [TestMethod]
        public void EditWatchRefusesChangedSelectionAndNativeFailures()
        {
            var fake = new WatchFake { Open = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Open = true;
            fake.Texts[4853] = "other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Texts[4853] = "x";
            fake.Texts[4857] = "Other";
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Texts[4857] = "Module1";
            fake.Missing = 4852;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Missing = 0;
            fake.Check = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Check = true;
            fake.ReplaceEcho = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.ReplaceEcho = true;
            fake.RefuseClick = 1;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
        }

        [TestMethod]
        public void EditWatchDistinguishesRejectionPendingAndVerifiedReadback()
        {
            var fake = new WatchFake { ErrorOpen = true, CloseAfterOk = false };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteEditWatch(Edit(), fake)).Message, "invalid expression");
            fake = new WatchFake { CloseAfterOk = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake = new WatchFake();
            dynamic hidden = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.IsTrue((bool)hidden.VerificationPending);
            fake = new WatchFake { Root = new IntPtr(4), PaneHandle = new IntPtr(5) };
            dynamic pending = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.AreEqual(40, fake.WatchReadAttempts);
            fake = new WatchFake { Root = new IntPtr(4), PaneHandle = new IntPtr(5), NewMatches = 1 };
            dynamic verified = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("ReadbackVerified", (string)verified.Verification);
            Assert.IsFalse((bool)verified.VerificationPending);
            fake = new WatchFake { Root = new IntPtr(4), PaneHandle = new IntPtr(5),
                NewMatches = 1, OldMatches = 1 };
            dynamic oldStillPresent = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("Pending", (string)oldStillPresent.Verification);
        }

        private sealed class WatchFake : VbeDebugWindows.IWatchProbe
        {
            public readonly Dictionary<int, string> Texts = new Dictionary<int, string> {
                { 4853, "x" }, { 4858, "Book" }, { 4857, "Module1" }, { 4856, "Run" } };
            public bool Open = true, CloseAfterOk = true, Check = true, ReplaceEcho = true, ErrorOpen;
            public int Missing, RefuseClick, Closes, Pauses, NewMatches, OldMatches, WatchReadAttempts;
            public IntPtr Root, PaneHandle, ListHandle;
            private bool clickedOk;
            public IntPtr Dialog(params string[] titles)
            {
                if (titles[0].StartsWith("Microsoft")) return ErrorOpen && clickedOk ? new IntPtr(9) : IntPtr.Zero;
                return Open && !(clickedOk && CloseAfterOk) ? new IntPtr(2) : IntPtr.Zero;
            }
            public IntPtr Item(IntPtr dialog, int id) { return id == Missing ? IntPtr.Zero : new IntPtr(id); }
            public string Text(IntPtr handle) { return Texts.TryGetValue(handle.ToInt32(), out string value) ? value : ""; }
            public bool Click(IntPtr handle)
            {
                if (handle.ToInt32() == RefuseClick) return false;
                if (handle.ToInt32() == 1) clickedOk = true;
                return true;
            }
            public bool Checked(IntPtr handle) { return Check; }
            public void Replace(IntPtr handle, string value) { if (ReplaceEcho) Texts[handle.ToInt32()] = value; }
            public void Pause(int milliseconds) { Pauses++; }
            public string Message(IntPtr dialog) { return "invalid expression"; }
            public void Close(IntPtr dialog) { Closes++; if (dialog.ToInt32() == 2) Open = false; }
            public IntPtr VbeRoot() { return Root; }
            public List<IntPtr> Children(IntPtr root) { return new List<IntPtr> { PaneHandle }; }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return PaneHandle; }
            public object List(IntPtr handle) { ListHandle = handle; return new { Available = handle != IntPtr.Zero }; }
            public int WatchMatches(IntPtr pane, string expression, string context)
            { WatchReadAttempts++; return expression == "y" ? NewMatches : OldMatches; }
        }
    }
}
