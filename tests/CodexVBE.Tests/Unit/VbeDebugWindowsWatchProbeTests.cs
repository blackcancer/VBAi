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

        [TestMethod]
        public void QuickWatchRequiresDialogControlsExpressionAndExactContext()
        {
            var request = new Request { Project = "Book", Module = "Module1",
                Procedure = "Run", Expression = "x" };
            var fake = new WatchFake { Open = false };
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteQuickWatch(request, fake));
            Assert.AreEqual(60, fake.Pauses);
            foreach (int missing in new[] { 4751, 4752, 4753, 2 })
            {
                fake = QuickFake();
                fake.Missing = missing;
                Assert.ThrowsException<InvalidOperationException>(
                    () => VbeDebugWindows.CompleteQuickWatch(request, fake));
                Assert.AreEqual(1, fake.Closes);
            }
            fake = QuickFake();
            fake.Texts[4751] = "other";
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteQuickWatch(request, fake));
            fake = QuickFake();
            fake.Texts[4753] = "Other.Module1.Run";
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteQuickWatch(request, fake));
            fake = QuickFake();
            fake.Texts[4753] = "Book.Module1.Other";
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.CompleteQuickWatch(request, fake));
        }

        [TestMethod]
        public void QuickWatchReadbackIncludesDisplayedValueAndAlwaysClosesOwnDialog()
        {
            var fake = QuickFake();
            dynamic result = VbeDebugWindows.CompleteQuickWatch(new Request {
                Project = "Book", Module = "Module1", Expression = "x" }, fake);
            Assert.AreEqual("42", (string)result.Value);
            Assert.AreEqual("Book.Module1.Run", (string)result.Context);
            Assert.AreEqual("NativeDialogReadback", (string)result.Verification);
            Assert.AreEqual(1, fake.Closes);
        }

        [TestMethod]
        public void SelectWatchRequiresUniqueVisibleSelectableNativeRow()
        {
            var request = new Request { Expression = "x", Context = "Book.Module1.Run" };
            var fake = new WatchFake();
            Assert.ThrowsException<ArgumentException>(
                () => VbeDebugWindows.SelectWatch(null, fake));
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.SelectWatch(request, fake));
            fake.Root = new IntPtr(4);
            fake.PaneHandle = new IntPtr(5);
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.SelectWatch(request, fake));
            fake.OldMatches = 2;
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.SelectWatch(request, fake));
            fake.OldMatches = 1;
            fake.SelectSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.SelectWatch(request, fake));
            fake.SelectSucceeds = true;
            dynamic selected = VbeDebugWindows.SelectWatch(request, fake);
            Assert.IsTrue((bool)selected.Selected);
            Assert.AreEqual(2, fake.SelectAttempts);
        }

        [TestMethod]
        public void VerifyWatchRemovedDistinguishesHiddenPendingTimeoutAndObservedAbsence()
        {
            var request = new Request { Expression = "x", Context = "Book.Module1.Run" };
            var fake = new WatchFake();
            dynamic hidden = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)hidden.VerificationPending);
            Assert.IsFalse((bool)hidden.Removed);
            fake.Root = new IntPtr(4);
            fake.PaneHandle = new IntPtr(5);
            dynamic absent = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)absent.Removed);
            fake.OldMatches = 1;
            dynamic pending = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual(20, fake.Pauses);
            fake.DisappearAfter = fake.WatchReadAttempts + 3;
            dynamic removed = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)removed.Removed);
            Assert.IsFalse((bool)removed.VerificationPending);
        }

        private static WatchFake QuickFake()
        {
            var fake = new WatchFake();
            fake.Texts[4751] = "x";
            fake.Texts[4752] = "42";
            fake.Texts[4753] = "Book.Module1.Run";
            return fake;
        }

        private sealed class WatchFake : VbeDebugWindows.IWatchProbe
        {
            public readonly Dictionary<int, string> Texts = new Dictionary<int, string> {
                { 4853, "x" }, { 4858, "Book" }, { 4857, "Module1" }, { 4856, "Run" } };
            public bool Open = true, CloseAfterOk = true, Check = true, ReplaceEcho = true, ErrorOpen;
            public int Missing, RefuseClick, Closes, Pauses, NewMatches, OldMatches, WatchReadAttempts,
                SelectAttempts, DisappearAfter;
            public bool SelectSucceeds = true;
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
            {
                WatchReadAttempts++;
                if (DisappearAfter > 0 && WatchReadAttempts >= DisappearAfter) return 0;
                return expression == "y" ? NewMatches : OldMatches;
            }
            public bool SelectWatchRow(IntPtr pane, string expression, string context)
            { SelectAttempts++; return SelectSucceeds; }
        }
    }
}
