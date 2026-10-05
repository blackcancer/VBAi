using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureOwnedBootstrapTraceTests
    {
        [DataTestMethod]
        [DataRow(null, false)]
        [DataRow("", false)]
        [DataRow("Q027Private", true)]
        [DataRow("q027private", true)]
        public void TraceSelectionPreservesMainCustomPathAndUsesExactPrivateInheritance(
            string desktop, bool privateDesktop)
        {
            string custom = Path.Combine(Path.GetTempPath(), "custom-trace.jsonl");
            string inherited = Path.Combine(Path.GetTempPath(), "inherited-trace.jsonl");
            string selected = ExcelVbeFixture.SelectOwnedTracePath(custom, desktop, inherited);
            Assert.IsTrue(string.Equals(privateDesktop ? inherited : custom, selected, StringComparison.Ordinal));
        }

        [DataTestMethod]
        [DataRow(null, "missing")]
        [DataRow("", "empty")]
        [DataRow("relative-trace.jsonl", "relative")]
        [DataRow("C:\\trace.jsonl:alternate", "alternate-stream")]
        public void PrivateTraceSelectionRefusesAnAbsentOrNonlocalInheritedPath(string inherited, string reason)
        {
            Assert.ThrowsException<ArgumentException>(() =>
                ExcelVbeFixture.SelectOwnedTracePath("unused-custom.jsonl", "Q027Private", inherited), reason);
        }

        [DataTestMethod]
        [DataRow(null, "unused.jsonl", null, null, true)]
        [DataRow(null, "unused.jsonl", "other.jsonl", "diagnostic.json", true)]
        [DataRow("Q027Private", "exact.jsonl", "exact.jsonl", null, true)]
        [DataRow("Q027Private", "exact.jsonl", "other.jsonl", null, false)]
        [DataRow("Q027Private", "exact.jsonl", null, null, false)]
        [DataRow("Q027Private", null, null, null, false)]
        [DataRow("Q027Private", " ", " ", null, false)]
        [DataRow("Q027Private", "exact.jsonl", " ", null, false)]
        [DataRow("Q027Private", "exact.jsonl", "exact.jsonl", "diagnostic.json", false)]
        [DataRow("Q027Private", "exact.jsonl", "other.jsonl", "diagnostic.json", false)]
        public void LaunchIntentIsRecordedOnceOnlyAfterEveryPrivateGuard(
            string desktop, string trace, string inherited, string manifest, bool accepted)
        {
            var events = new List<string>();
            int startAttempts = 0;
            Action intent = () => { events.Add("LaunchIntent"); startAttempts++; };
            if (accepted)
                ExcelVbeFixture.RecordLaunchIntentAfterPrivateGuards(desktop, trace, inherited, manifest, intent);
            else
                Assert.ThrowsException<InvalidOperationException>(() =>
                    ExcelVbeFixture.RecordLaunchIntentAfterPrivateGuards(desktop, trace, inherited, manifest, intent));
            Assert.AreEqual(accepted ? 1 : 0, startAttempts,
                "A rejected prelaunch contract cannot claim that a native start was attempted.");
            CollectionAssert.AreEqual(accepted ? new[] { "LaunchIntent" } : new string[0], events.ToArray());
        }

        [TestMethod]
        public void MissingIntentCallbackIsRejectedBeforeRecordingAnything()
        {
            Assert.ThrowsException<ArgumentNullException>(() =>
                ExcelVbeFixture.RecordLaunchIntentAfterPrivateGuards("Q027Private", "exact.jsonl",
                    "exact.jsonl", null, null));
        }

        [TestMethod]
        public void LaunchIntentRecordingFailureRemainsTheSameExceptionWithoutRetry()
        {
            var failure = new IOException("Synthetic durable intent failure");
            int attempts = 0;
            var actual = Assert.ThrowsException<IOException>(() =>
                ExcelVbeFixture.RecordLaunchIntentAfterPrivateGuards("Q027Private", "exact.jsonl",
                    "exact.jsonl", null, () => { attempts++; throw failure; }));
            Assert.AreSame(failure, actual);
            Assert.AreEqual(1, attempts);
        }

        [DataTestMethod]
        [DataRow(null, "Default", "Default", true)]
        [DataRow("", "Default", "default", true)]
        [DataRow("Q027Private", "Q027Private", "Q027Private", true)]
        [DataRow(null, "Default", "Other", false)]
        [DataRow("Q027Private", "Q027Private", "Default", false)]
        [DataRow("Q027Private", "Default", "Default", false)]
        [DataRow(null, "", "Default", false)]
        [DataRow(null, "Default", "", false)]
        public void OwnedWindowDesktopMustMatchTheObservedLaunchThreadAndPrivateName(
            string configured, string launchThread, string window, bool accepted)
        {
            if (accepted)
                ExcelVbeFixture.RequireOwnedWindowDesktop(configured, launchThread, window);
            else
                Assert.ThrowsException<InvalidOperationException>(() =>
                    ExcelVbeFixture.RequireOwnedWindowDesktop(configured, launchThread, window));
        }
    }
}
