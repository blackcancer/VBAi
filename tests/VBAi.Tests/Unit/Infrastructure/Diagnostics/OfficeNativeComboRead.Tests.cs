using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeNativeComboReadTests
    {
        private static OfficeNativeComboRead.TextRead Label(string text)
            => new OfficeNativeComboRead.TextRead { Text = text, Characters = text.Length };

        [DataTestMethod][DataRow("")][DataRow("Ollama")][DataRow("Project · Qualification française.docm")]
        public void SelectedDisplayTextIsReadOnceWithinItsExplicitBufferBound(string value)
        {
            int textReads = 0, indices = 0, lengths = 0;
            string result = OfficeNativeComboRead.ReadBounded(() => { }, () => { indices++; return 2; },
                () => { lengths++; return value.Length; }, capacity =>
                { Assert.AreEqual(2049, capacity); textReads++; return Label(value); });
            Assert.AreEqual(value, result); Assert.AreEqual(1, textReads);
            Assert.AreEqual(2, indices); Assert.AreEqual(2, lengths);
        }

        [TestMethod]
        public void MaximumLengthIsSupportedAndWindowsLengthOverestimationIsNotTreatedAsTruncation()
        {
            string value = new string('x', 2048);
            Assert.AreEqual(value, OfficeNativeComboRead.ReadBounded(() => { }, () => 0, () => 2048, unused => Label(value)));
            Assert.AreEqual("é", OfficeNativeComboRead.ReadBounded(() => { }, () => 0, () => 2, unused => Label("é")));
        }

        [TestMethod]
        public void UnselectedListHasNoLabelAndDoesNotEnterAStringRead()
        {
            int guards = 0;
            Assert.AreEqual("", OfficeNativeComboRead.ReadBounded(() => guards++, () => -1,
                () => { Assert.Fail("No selected label."); return 0; }, unused => { Assert.Fail(); return null; }));
            Assert.AreEqual(2, guards);
        }

        [DataTestMethod][DataRow(-2L)][DataRow(4096L)][DataRow(long.MaxValue)]
        public void InvalidIndexCannotStartALabelRead(long index)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboRead.ReadBounded(() => { }, () => index,
                () => { Assert.Fail("Invalid index must be rejected first."); return 0; }, unused => { Assert.Fail(); return null; }));
        }

        [DataTestMethod][DataRow(-1L)][DataRow(2049L)][DataRow(long.MaxValue)]
        public void OversizedOrUnavailableLabelsNeverEnterTheTextMessage(long length)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboRead.ReadBounded(() => { }, () => 0,
                () => length, unused => { Assert.Fail("No text message allowed."); return null; }));
        }

        [DataTestMethod][DataRow(0)][DataRow(1)][DataRow(2)][DataRow(3)][DataRow(4)]
        public void MissingContradictoryAndTruncatedReadbackCannotBecomeAnExactSelection(int scenario)
        {
            var reads = new[] { null, new OfficeNativeComboRead.TextRead { Characters = 0 },
                new OfficeNativeComboRead.TextRead { Text = "Ollama", Characters = 5 },
                new OfficeNativeComboRead.TextRead { Text = "Ollama", Characters = -1 }, Label("longer than expected") };
            int indexReads = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboRead.ReadBounded(() => { },
                () => { indexReads++; return 0; }, () => 6, unused => reads[scenario]));
            Assert.AreEqual(1, indexReads, "Invalid text must fail before further observation.");
        }

        [DataTestMethod][DataRow(true)][DataRow(false)]
        public void ConcurrentSelectionOrLabelChangeFailsWithoutAnotherTextRead(bool changedIndex)
        {
            int indices = 0, lengths = 0, texts = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboRead.ReadBounded(() => { },
                () => ++indices == 2 && changedIndex ? 1 : 0,
                () => ++lengths == 2 && !changedIndex ? 7 : 6,
                unused => { texts++; return Label("Ollama"); }));
            Assert.AreEqual(1, texts);
        }

        [DataTestMethod][DataRow(1, 0)][DataRow(2, 1)][DataRow(3, 2)][DataRow(4, 3)][DataRow(5, 4)][DataRow(6, 5)]
        public void OwnershipLossAtEveryBoundaryStopsAllRemainingNativeReads(int refusedGuard, int expectedReads)
        {
            int guards = 0, reads = 0;
            var failure = new InvalidOperationException("Original owner ended.");
            Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => OfficeNativeComboRead.ReadBounded(
                () => { if (++guards == refusedGuard) throw failure; }, () => { reads++; return 0; },
                () => { reads++; return 6; }, unused => { reads++; return Label("Ollama"); })));
            Assert.AreEqual(expectedReads, reads);
        }

        [DataTestMethod][DataRow(0)][DataRow(1)][DataRow(2)]
        public void NativeReadFailureRemainsTheOriginalFailureWithoutARepeat(int stage)
        {
            var failure = new TimeoutException("Native message timed out.");
            int indexReads = 0, lengthReads = 0, textReads = 0;
            Assert.AreSame(failure, Assert.ThrowsException<TimeoutException>(() => OfficeNativeComboRead.ReadBounded(() => { },
                () => { indexReads++; if (stage == 0) throw failure; return 0; },
                () => { lengthReads++; if (stage == 1) throw failure; return 6; },
                unused => { textReads++; throw failure; })));
            Assert.AreEqual(1, indexReads); Assert.AreEqual(stage >= 1 ? 1 : 0, lengthReads);
            Assert.AreEqual(stage == 2 ? 1 : 0, textReads);
        }

        [DataTestMethod][DataRow(0)][DataRow(1)][DataRow(2)][DataRow(3)]
        public void MissingDependenciesAreRejectedWithoutObservation(int missing)
        {
            Assert.ThrowsException<ArgumentNullException>(() => OfficeNativeComboRead.ReadBounded(
                missing == 0 ? null : (Action)(() => Assert.Fail()), missing == 1 ? null : (Func<long>)(() => 0),
                missing == 2 ? null : (Func<long>)(() => 0), missing == 3 ? null : (Func<int, OfficeNativeComboRead.TextRead>)(unused => null)));
        }

        [TestMethod]
        public void ExactNativeDropDownIdentityRequiresEveryOwnerAndRoleBoundary()
        {
            Func<int, uint, string, long, bool, bool, bool, bool> valid = (pid, thread, cls, style, exists, child, visible) =>
                OfficeNativeComboRead.IdentityMatches(42, 7, pid, thread, cls, style, exists, child, visible);
            Assert.IsTrue(valid(42, 7, "WindowsForms10.COMBOBOX.app", 3, true, true, true));
            Assert.IsFalse(valid(43, 7, "WindowsForms10.COMBOBOX.app", 3, true, true, true));
            Assert.IsFalse(valid(42, 8, "WindowsForms10.COMBOBOX.app", 3, true, true, true));
            foreach (var cls in new[] { null, "", "COMBOBOX", "WindowsForms10.EDIT.app" })
                Assert.IsFalse(valid(42, 7, cls, 3, true, true, true));
            foreach (int style in new[] { 0, 1, 2 }) Assert.IsFalse(valid(42, 7, "WindowsForms10.COMBOBOX.app", style, true, true, true));
            Assert.IsFalse(valid(42, 7, "WindowsForms10.COMBOBOX.app", 3, false, true, true));
            Assert.IsFalse(valid(42, 7, "WindowsForms10.COMBOBOX.app", 3, true, false, true));
            Assert.IsFalse(valid(42, 7, "WindowsForms10.COMBOBOX.app", 3, true, true, false));
            Assert.IsFalse(OfficeNativeComboRead.IdentityMatches(0, 7, 0, 7, "WindowsForms10.COMBOBOX.app", 3, true, true, true));
            Assert.IsFalse(OfficeNativeComboRead.IdentityMatches(42, 0, 42, 0, "WindowsForms10.COMBOBOX.app", 3, true, true, true));
        }
    }
}
