using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficePowerPointShutdownTests
    {
        [TestMethod]
        public void SoleOwnedPresentationQuitsOnceAfterSavedReadbackAndFinalValidation()
        {
            var calls = new List<string>();
            OfficeVbeFixture.PreparePowerPointQuit(() => calls.Add("validate"), () => calls.Add("saved"),
                () => { calls.Add("readback"); return true; }, () => calls.Add("quit"));
            CollectionAssert.AreEqual(new[] { "validate", "saved", "readback", "validate", "quit" }, calls);
        }

        [TestMethod]
        [DataRow("initial-validation", 0, 0), DataRow("saved-write", 1, 0)]
        [DataRow("saved-read", 1, 0), DataRow("final-validation", 1, 0), DataRow("quit", 1, 1)]
        public void BoundaryFailureStopsWithoutNativeReplay(string phase, int writes, int quits)
        {
            int validations = 0, actualWrites = 0, actualQuits = 0;
            var failure = new InvalidOperationException("Original uncertain " + phase);
            Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.PreparePowerPointQuit(
                    () => { validations++; if (phase == (validations == 1 ? "initial-validation" : "final-validation")) throw failure; },
                    () => { actualWrites++; if (phase == "saved-write") throw failure; },
                    () => { if (phase == "saved-read") throw failure; return true; },
                    () => { actualQuits++; if (phase == "quit") throw failure; })));
            Assert.AreEqual(writes, actualWrites); Assert.AreEqual(quits, actualQuits);
        }

        [TestMethod]
        public void UnretainedSavedFlagDoesNotAdmitQuit()
        {
            int writes = 0, quits = 0, guards = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.PreparePowerPointQuit(
                () => guards++, () => writes++, () => false, () => quits++));
            Assert.AreEqual(1, guards); Assert.AreEqual(1, writes); Assert.AreEqual(0, quits);
        }

        [TestMethod]
        [DataRow("validate"), DataRow("saved-write"), DataRow("saved-read"), DataRow("quit")]
        public void IncompleteShutdownBoundaryDoesNotInvokeAnyAction(string missing)
        {
            Action refused = () => Assert.Fail("No action is permitted.");
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.PreparePowerPointQuit(
                missing == "validate" ? null : refused, missing == "saved-write" ? null : refused,
                missing == "saved-read" ? null : new Func<bool>(() => { refused(); return true; }),
                missing == "quit" ? null : refused));
        }

        [TestMethod]
        public void ExactSoleDisposableInventoryIsAcceptedWithoutSavingOrClosingAnything()
            => OfficeVbeFixture.ValidatePowerPointQuitInventory(1, 1, true, @"C:\owned", @"C:\owned\Disposable.pptm",
                @"c:\OWNED\Disposable.pptm", @"C:\owned\Disposable.pptm", 2);

        [TestMethod]
        [DataRow("zero-document"), DataRow("extra-document"), DataRow("zero-project"), DataRow("extra-project")]
        [DataRow("foreign-identity"), DataRow("running"), DataRow("break"), DataRow("unknown-mode")]
        [DataRow("foreign-document"), DataRow("foreign-project"), DataRow("outside-root"), DataRow("prefix-sibling")]
        [DataRow("null-root"), DataRow("null-expected"), DataRow("null-document"), DataRow("null-project")]
        public void ChangedOrUnreadableInventoryIsRejectedBeforeSavedFlagOrQuit(string scenario)
        {
            int docs = scenario == "zero-document" ? 0 : scenario == "extra-document" ? 2 : 1;
            int projects = scenario == "zero-project" ? 0 : scenario == "extra-project" ? 2 : 1;
            int mode = scenario == "running" ? 0 : scenario == "break" ? 1 : scenario == "unknown-mode" ? 3 : 2;
            string root = scenario == "null-root" ? null : @"C:\owned";
            string expected = scenario == "null-expected" ? null : @"C:\owned\Disposable.pptm";
            string actual = scenario == "null-document" ? null : scenario == "foreign-document" ? @"C:\other.pptm" : expected;
            string project = scenario == "null-project" ? null : scenario == "foreign-project" ? @"C:\other.pptm" : expected;
            if (scenario == "outside-root") root = @"C:\other";
            if (scenario == "prefix-sibling") { expected = actual = project = @"C:\owned-sibling\Disposable.pptm"; }
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.ValidatePowerPointQuitInventory(
                docs, projects, scenario != "foreign-identity", root, expected, actual, project, mode));
        }
    }
}
