using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OptionsCheckboxTests
    {
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void NativeCheckboxChangesExactlyOnceAndReadsTheResult(bool desired)
        {
            int state = desired ? 0 : 1, clicks = 0, reads = 0;
            VbeDebugWindows.SetOptionsCheckbox(desired, () => { reads++; return state; },
                () => { clicks++; state = desired ? 1 : 0; });
            Assert.AreEqual(1, clicks); Assert.AreEqual(2, reads);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void MatchingNativeCheckboxNeverClicks(bool desired)
        {
            int clicks = 0;
            VbeDebugWindows.SetOptionsCheckbox(desired, () => desired ? 1 : 0, () => clicks++);
            Assert.AreEqual(0, clicks);
        }

        [DataTestMethod, DataRow(-1), DataRow(2), DataRow(3)]
        public void UnreadableInitialCheckboxRefusesBeforeClick(int state)
        {
            int clicks = 0;
            Assert.ThrowsException<InvalidOperationException>(() =>
                VbeDebugWindows.SetOptionsCheckbox(false, () => state, () => clicks++));
            Assert.AreEqual(0, clicks);
        }

        [DataTestMethod, DataRow(0), DataRow(2), DataRow(-1)]
        public void UnverifiedClickIsNeverRepeated(int after)
        {
            int clicks = 0, reads = 0;
            Assert.ThrowsException<InvalidOperationException>(() =>
                VbeDebugWindows.SetOptionsCheckbox(true, () => ++reads == 1 ? 0 : after, () => clicks++));
            Assert.AreEqual(1, clicks); Assert.AreEqual(2, reads);
        }

        [TestMethod]
        public void ThrowingClickStopsWithoutAnotherReadOrClick()
        {
            int clicks = 0, reads = 0;
            var error = new InvalidOperationException("native provider failed");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() =>
                VbeDebugWindows.SetOptionsCheckbox(false, () => { reads++; return 1; },
                    () => { clicks++; throw error; })));
            Assert.AreEqual(1, clicks); Assert.AreEqual(1, reads);
        }
    }
}
