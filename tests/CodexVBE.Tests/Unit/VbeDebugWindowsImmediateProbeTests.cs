using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsImmediateProbeTests
    {
        [TestMethod]
        public void ImmediateRejectsNonPrintableMultilineOrOversizedCommandsBeforeNativeAccess()
        {
            var fake = new ImmediateFake();
            foreach (string command in new[] { null, "", " ", "a\nb", "a\rb", "a\0b",
                "a\tb", new string('x', 2049) })
                Assert.ThrowsException<ArgumentException>(
                    () => VbeDebugWindows.ExecuteImmediate(command, fake));
            Assert.AreEqual(0, fake.RootReads);
        }

        [TestMethod]
        public void ImmediateRequiresVisibleVbeAndPaneBeforePreparingSelection()
        {
            var fake = new ImmediateFake();
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            fake.Root = new IntPtr(1);
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(0, fake.Prepares);
        }

        [TestMethod]
        public void ImmediateDoesNotSendEnterUntilEveryCharacterEchoesExactly()
        {
            var fake = new ImmediateFake { Root = new IntPtr(1), PaneHandle = new IntPtr(2),
                RejectChar = '?' };
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(0, fake.Enters);
            fake = new ImmediateFake { Root = new IntPtr(1), PaneHandle = new IntPtr(2) };
            fake.Readbacks.Add("unexpected");
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(40, fake.Pauses);
            Assert.AreEqual(0, fake.Enters);
        }

        [TestMethod]
        public void ImmediateDistinguishesRejectedEnterPendingAndChangedOutput()
        {
            var fake = Ready("? 1");
            fake.EnterSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(
                () => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(1, fake.Enters);
            fake = Ready("? 1");
            dynamic pending = VbeDebugWindows.ExecuteImmediate("? 1", fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual("? 1", (string)pending.OutputDelta);
            fake = Ready("? 1");
            fake.Readbacks.Add("> ? 1\r\n1\r\n");
            dynamic changed = VbeDebugWindows.ExecuteImmediate("? 1", fake);
            Assert.AreEqual("ImmediateTextChangedAfterEnter", (string)changed.Verification);
            Assert.IsFalse((bool)changed.VerificationPending);
            Assert.AreEqual("? 1\r\n1\r\n", (string)changed.OutputDelta);
        }

        private static ImmediateFake Ready(string command)
        {
            var fake = new ImmediateFake { Root = new IntPtr(1), PaneHandle = new IntPtr(2) };
            fake.Readbacks.Add("> " + command);
            return fake;
        }

        private sealed class ImmediateFake : VbeDebugWindows.IImmediateProbe
        {
            public IntPtr Root, PaneHandle;
            public int RootReads, Prepares, Enters, Pauses;
            public char RejectChar;
            public bool EnterSucceeds = true;
            public readonly List<string> Readbacks = new List<string>();
            private int readIndex;
            private string last = "> ";
            public IntPtr VbeRoot() { RootReads++; return Root; }
            public List<IntPtr> Children(IntPtr root) { return new List<IntPtr> { PaneHandle }; }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return PaneHandle; }
            public string Prepare(IntPtr pane) { Prepares++; return "> "; }
            public string Text(IntPtr pane)
            {
                if (readIndex < Readbacks.Count) last = Readbacks[readIndex++];
                return last;
            }
            public bool PostChar(IntPtr pane, char character) { return character != RejectChar; }
            public bool PostEnter(IntPtr pane) { Enters++; return EnterSucceeds; }
            public void Pause(int milliseconds) { Pauses++; }
        }
    }
}
