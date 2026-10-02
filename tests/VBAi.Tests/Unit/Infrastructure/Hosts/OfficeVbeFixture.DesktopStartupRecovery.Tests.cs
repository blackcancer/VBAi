using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureDesktopStartupRecoveryTests
    {
        private static OfficeVbeFixture.PublisherRecoverySnapshot Exact()
        {
            return new OfficeVbeFixture.PublisherRecoverySnapshot {
                Dialog = new IntPtr(11), Button = new IntPtr(12), ButtonRoot = new IntPtr(11),
                DialogPid = 100, ButtonPid = 100, DialogThread = 200, ButtonThread = 200, ButtonId = 7,
                DialogClass = "#32770", ButtonClass = "Button", Caption = "Microsoft Publisher",
                Message = OfficeVbeFixture.FrenchPublisherRecoveryMessage, ButtonText = "&Non" };
        }

        [TestMethod]
        public void RecoveryMatchesOnlyObservedFrenchQuestionAndNativeNoButton()
        {
            var value = Exact(); Assert.IsTrue(OfficeVbeFixture.IsExactPublisherRecovery(value, 100));
            value.ButtonText = "Non"; value.Message = value.Message.Replace("\n", "\r\n");
            Assert.IsTrue(OfficeVbeFixture.IsExactPublisherRecovery(value, 100));
            Assert.IsFalse(OfficeVbeFixture.IsExactPublisherRecovery(null, 100));
            Assert.IsFalse(OfficeVbeFixture.IsExactPublisherRecovery(value, 0));
        }

        [TestMethod]
        public void RecoveryRefusesWrongOwnerThreadClassIdAncestorMessageAndLocale()
        {
            foreach (var change in new Action<OfficeVbeFixture.PublisherRecoverySnapshot>[] {
                value => value.DialogPid = 101, value => value.ButtonPid = 101,
                value => value.DialogThread = 0, value => value.ButtonThread = 201,
                value => value.Dialog = IntPtr.Zero, value => value.Button = IntPtr.Zero,
                value => value.ButtonId = 6, value => value.DialogClass = "OtherDialog", value => value.ButtonClass = "OtherButton",
                value => value.ButtonRoot = new IntPtr(99), value => value.Caption = "Microsoft Access",
                value => value.Message = "Publisher failed to start. Start in safe mode?",
                value => value.Message = value.Message.Replace("\u00A0", " "),
                value => value.ButtonText = "Oui", value => value.Message += " Extra text." })
            {
                var value = Exact(); change(value);
                Assert.IsFalse(OfficeVbeFixture.IsExactPublisherRecovery(value, 100));
            }
        }

        [TestMethod]
        public void UncertainRecoveryClickRetainsClaimAndRejectsEveryReplay()
        {
            var gate = new OfficeVbeFixture.PublisherStartupRecoveryGate(100, 300);
            int clicks = 0, reads = 0; var rows = new List<IDictionary<string, object>>();
            Assert.ThrowsException<InvalidOperationException>(() => gate.RequestNormalStart(Exact(), () => { },
                () => { reads++; return Exact(); }, () => { clicks++; return false; }, rows.Add));
            Assert.IsTrue(gate.Claimed); Assert.AreEqual(1, clicks); Assert.AreEqual(1, reads);
            Assert.AreEqual("CLICK_PENDING", rows[0]["PublisherStartupRecovery"]);
            Assert.AreEqual("CLICK_UNCERTAIN", rows[1]["PublisherStartupRecovery"]);
            Assert.AreEqual(true, rows[1]["OutcomeUncertain"]); Assert.AreEqual(false, rows[1]["NormalModeProven"]);
            Assert.ThrowsException<InvalidOperationException>(() => gate.RequestNormalStart(Exact(), () => { },
                () => { Assert.Fail("No read replay after an uncertain click."); return Exact(); },
                () => { Assert.Fail("No native click replay."); return true; }, rows.Add));
            Assert.AreEqual(1, clicks); Assert.AreEqual(2, rows.Count);
        }

        [TestMethod]
        public void ReturnedRecoveryClickRecordsNormalRequestWithoutClaimingNativeModeProof()
        {
            var gate = new OfficeVbeFixture.PublisherStartupRecoveryGate(100, 300);
            int guards = 0, clicks = 0; var rows = new List<IDictionary<string, object>>();
            gate.RequestNormalStart(Exact(), () => guards++, Exact, () => { clicks++; return true; }, rows.Add);
            Assert.AreEqual(2, guards); Assert.AreEqual(1, clicks); Assert.IsTrue(gate.Claimed);
            Assert.AreEqual("CLICK_RETURNED", rows[1]["PublisherStartupRecovery"]);
            Assert.AreEqual(true, rows[1]["NormalStartRequested"]); Assert.AreEqual(false, rows[1]["NormalModeProven"]);
            Assert.AreEqual(1, rows[1]["NativeClickAttempts"]); Assert.AreEqual(100u, rows[1]["ProcessId"]);
            Assert.AreEqual(300L, rows[1]["OriginalHandle"]); Assert.AreEqual("fr-FR", rows[1]["Locale"]);
            Assert.ThrowsException<InvalidOperationException>(() => gate.RequestNormalStart(Exact(), () => { }, Exact,
                () => { clicks++; return true; }, rows.Add));
            Assert.AreEqual(1, clicks);
        }

        [TestMethod]
        public void ChangedRecoveryBeforeClickRefusesMutationAndUnknownPromptDoesNotClaimDecision()
        {
            var gate = new OfficeVbeFixture.PublisherStartupRecoveryGate(100, 300);
            var changed = Exact(); changed.Button = new IntPtr(99);
            Assert.ThrowsException<InvalidOperationException>(() => gate.RequestNormalStart(Exact(), () => { }, () => changed,
                () => { Assert.Fail("Changed HWND must refuse the native click."); return true; }, row => Assert.Fail("No click intent may be recorded.")));
            Assert.IsFalse(gate.Claimed);
            var unknown = Exact(); unknown.Message = "Unknown prompt";
            Assert.ThrowsException<InvalidOperationException>(() => gate.RequestNormalStart(unknown, () => Assert.Fail("Unknown prompt must refuse before guard/read."),
                Exact, () => { Assert.Fail("Unknown prompt must never be clicked."); return true; }, row => Assert.Fail("No click intent may be recorded.")));
            Assert.IsFalse(gate.Claimed);
        }
    }
}
