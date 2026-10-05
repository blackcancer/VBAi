using System;
using System.Collections.Generic;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Prepares all post-selection state transitions without Office, native input or actual waiting.</summary>
    [TestClass]
    public sealed class WordChatScopeIdleTests
    {
        [DataTestMethod]
        [DataRow("exact")]
        [DataRow("zero")]
        [DataRow("replaced")]
        [DataRow("native-process")]
        [DataRow("ui-process")]
        [DataRow("thread")]
        [DataRow("parent")]
        [DataRow("type")]
        [DataRow("id")]
        [DataRow("missing")]
        [DataRow("expected-zero")]
        public void PickerGuardRequiresTheFrozenNativeAndUiIdentity(string change)
        {
            var actual = new WordChatScopeIdle.PickerIdentity {
                Handle = 41, NativeProcessId = 10, UiProcessId = 10, NativeThreadId = 20,
                WithinChat = true, IsComboBox = true, AutomationId = "scopePicker"
            };
            if (change == "zero") actual.Handle = 0;
            if (change == "replaced") actual.Handle = 42;
            if (change == "native-process") actual.NativeProcessId = 11;
            if (change == "ui-process") actual.UiProcessId = 11;
            if (change == "thread") actual.NativeThreadId = 21;
            if (change == "parent") actual.WithinChat = false;
            if (change == "type") actual.IsComboBox = false;
            if (change == "id") actual.AutomationId = "replacement";
            if (change == "missing") actual = null;
            if (change == "exact") WordChatScopeIdle.RequirePicker(41, 10, 20, actual);
            else Assert.ThrowsException<InvalidOperationException>(() => WordChatScopeIdle.RequirePicker(change == "expected-zero" ? 0 : 41, 10, 20, actual));
        }

        private static WordChatScopeIdle.Observation Sample(bool enabled = true, bool exact = true,
            ExpandCollapseState state = ExpandCollapseState.Collapsed)
        { return new WordChatScopeIdle.Observation { Enabled = enabled, ExactSelection = exact, State = state }; }

        private sealed class Probe
        {
            internal long Elapsed;
            internal int Reads, Collapses, Guards, Pauses;
            internal Func<WordChatScopeIdle.Observation> Read;
            internal Action Collapse, Guard, Pause;
            internal void Run()
            {
                WordChatScopeIdle.Wait(() => { Guards++; Guard?.Invoke(); },
                    () => { Reads++; return Read(); }, () => { Collapses++; Collapse?.Invoke(); },
                    () => Elapsed, () => { Pauses++; Elapsed += 50; Pause?.Invoke(); });
            }
        }

        [TestMethod]
        public void AlreadyCollapsedExactSelectionNeedsTwoSamplesAndNoAction()
        {
            var p = new Probe { Read = () => Sample() }; p.Run();
            Assert.AreEqual(2, p.Reads); Assert.AreEqual(2, p.Guards);
            Assert.AreEqual(1, p.Pauses); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void DisabledTransitionWaitsForExactCollapsedIdleWithoutCollapse()
        {
            var samples = new Queue<WordChatScopeIdle.Observation>(new[] { Sample(false, true, ExpandCollapseState.Expanded), Sample(false), Sample(), Sample() });
            var p = new Probe { Read = () => samples.Dequeue() }; p.Run();
            Assert.AreEqual(4, p.Reads); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void EnabledExactExpandedPopupClosesOnceBeforeTwoStableSamples()
        {
            bool closed = false;
            var p = new Probe { Read = () => Sample(state: closed ? ExpandCollapseState.Collapsed : ExpandCollapseState.Expanded), Collapse = () => closed = true };
            p.Run(); Assert.AreEqual(1, p.Collapses); Assert.AreEqual(3, p.Reads); Assert.AreEqual(4, p.Guards);
        }

        [TestMethod]
        public void DisabledExpandedPopupNeverReceivesCollapse()
        {
            var p = new Probe { Read = () => Sample(false, true, ExpandCollapseState.Expanded) };
            Assert.ThrowsException<TimeoutException>(() => p.Run());
            Assert.AreEqual(0, p.Collapses); Assert.AreEqual(15000L, p.Elapsed);
        }

        [TestMethod]
        public void EnabledWrongSelectionNeverReceivesCollapseOrAcceptance()
        {
            var p = new Probe { Read = () => Sample(true, false, ExpandCollapseState.Expanded) };
            Assert.ThrowsException<TimeoutException>(() => p.Run()); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void UnchangedExpandedPopupAfterOneCollapseIsNotRetried()
        {
            var p = new Probe { Read = () => Sample(state: ExpandCollapseState.Expanded) };
            Assert.ThrowsException<TimeoutException>(() => p.Run()); Assert.AreEqual(1, p.Collapses);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void DisabledOrWrongSecondSampleResetsStability(bool wrongSelection)
        {
            var q = new Queue<WordChatScopeIdle.Observation>(new[] { Sample(), Sample(wrongSelection, !wrongSelection), Sample(), Sample() });
            var p = new Probe { Read = () => q.Dequeue() }; p.Run();
            Assert.AreEqual(4, p.Reads); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void ExpiredOriginalBoundHasNoReadOrMutation()
        {
            var p = new Probe { Elapsed = 15000, Read = () => Sample() };
            Assert.ThrowsException<TimeoutException>(() => p.Run());
            Assert.AreEqual(0, p.Reads); Assert.AreEqual(0, p.Guards); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void OneGoodSampleAtDeadlineCannotBecomeAcceptance()
        {
            var p = new Probe { Elapsed = 14950, Read = () => Sample() };
            Assert.ThrowsException<TimeoutException>(() => p.Run()); Assert.AreEqual(1, p.Reads);
        }

        [DataTestMethod]
        [DataRow("initial-guard", 0)]
        [DataRow("second-good-read", 2)]
        [DataRow("expanded-read", 1)]
        [DataRow("collapse-guard", 1)]
        public void DelayedGuardOrReadCannotAdmitLateAcceptanceOrCollapse(string phase, int expectedReads)
        {
            var p = new Probe();
            p.Read = () => {
                if (phase == "expanded-read" || phase == "second-good-read" && p.Reads == 2) p.Elapsed = 15000;
                return Sample(state: phase == "expanded-read" || phase == "collapse-guard"
                    ? ExpandCollapseState.Expanded : ExpandCollapseState.Collapsed);
            };
            p.Guard = () => {
                if (phase == "initial-guard" || phase == "collapse-guard" && p.Guards == 2) p.Elapsed = 15000;
            };
            Assert.ThrowsException<TimeoutException>(() => p.Run());
            Assert.AreEqual(expectedReads, p.Reads); Assert.AreEqual(0, p.Collapses);
        }

        [DataTestMethod]
        [DataRow(ExpandCollapseState.PartiallyExpanded)]
        [DataRow(ExpandCollapseState.LeafNode)]
        [DataRow((ExpandCollapseState)42)]
        public void UnexpectedStateFailsBeforeAnyCollapse(ExpandCollapseState state)
        {
            var p = new Probe { Read = () => Sample(state: state) };
            Assert.ThrowsException<InvalidOperationException>(() => p.Run()); Assert.AreEqual(0, p.Collapses);
        }

        [TestMethod]
        public void MissingObservationFailsWithoutMutation()
        {
            var p = new Probe { Read = () => null };
            Assert.ThrowsException<InvalidOperationException>(() => p.Run()); Assert.AreEqual(0, p.Collapses);
        }

        [DataTestMethod]
        [DataRow("read")]
        [DataRow("guard")]
        [DataRow("collapse")]
        [DataRow("pause")]
        public void CallbackFailureIdentityIsPreservedAndNeverReplayed(string phase)
        {
            var original = new InvalidOperationException(phase);
            var p = new Probe { Read = () => Sample(state: ExpandCollapseState.Expanded) };
            if (phase == "read") p.Read = () => throw original;
            if (phase == "guard") p.Guard = () => throw original;
            if (phase == "collapse") p.Collapse = () => throw original;
            if (phase == "pause") p.Pause = () => throw original;
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => p.Run()));
            Assert.AreEqual(phase == "collapse" || phase == "pause" ? 1 : 0, p.Collapses);
        }

        [TestMethod]
        public void OwnerChangeBeforeCollapsePreventsDispatch()
        {
            var original = new InvalidOperationException("Owner changed"); int guards = 0;
            var p = new Probe { Read = () => Sample(state: ExpandCollapseState.Expanded), Guard = () => { if (++guards == 2) throw original; } };
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => p.Run()));
            Assert.AreEqual(0, p.Collapses);
        }

        [DataTestMethod]
        [DataRow("guard")]
        [DataRow("read")]
        [DataRow("collapse")]
        [DataRow("clock")]
        [DataRow("pause")]
        public void MissingDependencyFailsBeforeAnyCallback(string missing)
        {
            int calls = 0;
            Assert.ThrowsException<ArgumentNullException>(() => WordChatScopeIdle.Wait(
                missing == "guard" ? null : (Action)(() => calls++),
                missing == "read" ? null : (Func<WordChatScopeIdle.Observation>)(() => { calls++; return Sample(); }),
                missing == "collapse" ? null : (Action)(() => calls++),
                missing == "clock" ? null : (Func<long>)(() => { calls++; return 0; }),
                missing == "pause" ? null : (Action)(() => calls++)));
            Assert.AreEqual(0, calls);
        }
    }
}
