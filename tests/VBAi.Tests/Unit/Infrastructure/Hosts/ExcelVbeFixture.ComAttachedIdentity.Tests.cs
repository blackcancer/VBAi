using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureComAttachedIdentityTests
    {
        [TestMethod]
        public void RecordsRetainedHandleAndDesktopWithoutClaimingAnOriginalCreationHandle()
        {
            var receipt = new Dictionary<string, object>();
            ExcelVbeFixture.RecordComAttachedIdentity(receipt, 42, 42, 1234,
                "2026-10-05T10:00:00.0000000Z", 77, "Default", "Default", "Default");
            Assert.AreEqual("ComActivator_GetProcessById_Handle", receipt["ProcessIdentityOrigin"]);
            Assert.AreEqual(1234L, receipt["RetainedProcessHandle"]);
            Assert.AreEqual(42, receipt["ApplicationHwndProcessId"]);
            Assert.AreEqual("2026-10-05T10:00:00.0000000Z", receipt["HostStartedUtc"]);
            Assert.AreEqual("Default", receipt["OwnerThreadDesktop"]);
            Assert.AreEqual("Default", receipt["ApplicationWindowDesktop"]);
            Assert.IsFalse(receipt.ContainsKey("OriginalLaunchHandle"));
            var shutdown = new Dictionary<string, object>();
            ExcelVbeFixture.CopyComAttachedIdentityToShutdown(receipt, shutdown);
            Assert.AreEqual(receipt["HostStartedUtc"], shutdown["ProcessStartedUtc"]);
            foreach (string key in new[] { "ProcessIdentityOrigin", "RetainedProcessHandle", "ApplicationHwndProcessId",
                "ApplicationWindowThreadId", "OwnerThreadDesktop", "ApplicationWindowDesktop", "InputDesktopAtObservation" })
                Assert.AreEqual(receipt[key], shutdown[key], key);
            Assert.IsFalse(shutdown.ContainsKey("OriginalLaunchHandle"));
        }

        [DataTestMethod]
        [DataRow("pid")]
        [DataRow("window-pid")]
        [DataRow("handle")]
        [DataRow("negative-handle")]
        [DataRow("birth")]
        [DataRow("thread")]
        [DataRow("owner-desktop")]
        [DataRow("window-desktop")]
        [DataRow("input-desktop")]
        public void RejectsIncompleteOrChangedComIdentityBeforeRecordingAnyClaim(string fault)
        {
            var receipt = new Dictionary<string, object>();
            int pid = fault == "pid" ? 0 : 42;
            int windowPid = fault == "window-pid" ? 43 : 42;
            long handle = fault == "handle" ? 0 : fault == "negative-handle" ? -1 : 1234;
            string birth = fault == "birth" ? "invalid" : "2026-10-05T10:00:00.0000000Z";
            uint thread = fault == "thread" ? 0u : 77u;
            string owner = fault == "owner-desktop" ? "Another" : "Default";
            string window = fault == "window-desktop" ? "Another" : "Default";
            string input = fault == "input-desktop" ? " " : "Default";
            Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.RecordComAttachedIdentity(
                receipt, pid, windowPid, handle, birth, thread, owner, window, input));
            Assert.AreEqual(0, receipt.Count);
        }

        [TestMethod]
        public void RejectsMissingReceiptBeforeObservingIdentity()
        {
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.RecordComAttachedIdentity(
                null, 42, 42, 1234, "2026-10-05T10:00:00.0000000Z", 77, "Default", "Default", "Default"));
        }

        [TestMethod]
        public void ShutdownCopyRejectsIncompleteIdentityAndLeavesNoFalseCompletion()
        {
            var startup = new Dictionary<string, object>();
            ExcelVbeFixture.RecordComAttachedIdentity(startup, 42, 42, 1234,
                "2026-10-05T10:00:00.0000000Z", 77, "Default", "Default", "Default");
            startup.Remove("InputDesktopAtObservation");
            var shutdown = new Dictionary<string, object>();
            Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.CopyComAttachedIdentityToShutdown(
                startup, shutdown));
            Assert.AreEqual(0, shutdown.Count);
        }

        [TestMethod]
        public void ShutdownCopyLeavesExplicitProcessStartIdentitySeparate()
        {
            var startup = new Dictionary<string, object> { ["OriginalLaunchHandle"] = 1234L };
            var shutdown = new Dictionary<string, object>();
            ExcelVbeFixture.CopyComAttachedIdentityToShutdown(startup, shutdown);
            Assert.AreEqual(0, shutdown.Count);
        }
    }
}
