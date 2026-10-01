using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class PairedExportQualificationStateTests
    {
        private static readonly string[] Paths = { "local-guid", "local-guid/synthetic", "temp-guid", "temp-guid/synthetic" };
        private static IDictionary<string, object> Response(bool ok) => new Dictionary<string, object> { ["Ok"] = ok, ["Error"] = ok ? null : "Original native PATH_NOT_FOUND" };
        private static IDictionary<string, object> Observation() => new Dictionary<string, object> {
            ["ProcessId"] = 10, ["FinalProcessId"] = 10, ["NativeThreadId"] = 20u, ["FinalNativeThreadId"] = 20u,
            ["Apartment"] = "STA", ["AssemblyMvid"] = "candidate", ["EffectiveTokenBefore"] = new Dictionary<string, object> { ["State"] = "READ" },
            ["EffectiveTokenAfter"] = new Dictionary<string, object> { ["State"] = "READ" },
            ["Paths"] = Paths.Select(path => (object)new Dictionary<string, object> { ["Path"] = path,
                ["NativeAttributes"] = uint.MaxValue, ["NativeLastError"] = 3, ["NativeSucceeded"] = false,
                ["ManagedAttributesError"] = new Dictionary<string, object> { ["Type"] = "FileNotFoundException", ["HResult"] = unchecked((int)0x80070002) } }).ToArray() };
        private static PairedExportQualificationState Ready()
        {
            var state = new PairedExportQualificationState(); state.Begin(PathVisibilityDiagnostic.CommandName);
            state.Receive(PathVisibilityDiagnostic.CommandName, Response(true));
            state.VerifyOwner(Observation(), 10, 20, "candidate", Paths); return state;
        }

        [TestMethod]
        public void MissingOwnerObservationBlocksTheFirstExportBeforeAnyRequest()
        {
            var state = new PairedExportQualificationState();
            Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component"));
            Assert.AreEqual(0, state.ExportRequests); Assert.IsFalse(state.Pending);
        }

        [TestMethod]
        public void DifferingDeniedVisibilityRemainsEvidenceAndOneTerminalNegativeExportCannotReplay()
        {
            var state = Ready(); state.Begin("export_component"); var response = Response(false);
            state.Receive("export_component", response);
            Assert.IsFalse(state.Pending); Assert.AreEqual(1, state.ExportRequests);
            StringAssert.Contains(state.Stage, "TerminalFailure"); Assert.AreEqual("Original native PATH_NOT_FOUND", response["Error"]);
            // Read-only post-failure observation is permitted; a new export is always forbidden.
            state.Begin("read_module"); state.Receive("read_module", Response(true));
            Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component"));
            Assert.AreEqual(1, state.ExportRequests);
        }

        [TestMethod]
        public void ExportSuccessAlsoExhaustsTheSingleRequestAllowance()
        {
            var state = Ready(); state.Begin("export_component"); state.Receive("export_component", Response(true));
            Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component"));
            Assert.AreEqual(1, state.ExportRequests);
        }

        [TestMethod]
        public void MissingOrMalformedResponseRetainsPendingAndForbidsCleanupFollowupOrExportRetry()
        {
            foreach (var response in new[] { null, new Dictionary<string, object>(), new Dictionary<string, object> { ["Ok"] = "false" } })
            {
                var state = Ready(); state.Begin("export_component");
                Assert.ThrowsException<InvalidOperationException>(() => state.Receive("export_component", response));
                Assert.IsTrue(state.Pending); Assert.AreEqual(1, state.ExportRequests);
                Assert.ThrowsException<InvalidOperationException>(() => state.Begin("read_module"));
                Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component"));
            }
        }

        [TestMethod]
        public void InterveningHostRequestInvalidatesImmediatelyPrecedingDiagnostic()
        {
            var state = Ready(); state.Begin("component_properties"); state.Receive("component_properties", Response(true));
            Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component")); Assert.AreEqual(0, state.ExportRequests);
        }

        [TestMethod]
        public void WrongCommandCannotClearThePendingDelivery()
        {
            var state = Ready(); state.Begin("export_component");
            Assert.ThrowsException<InvalidOperationException>(() => state.Receive("read_module", Response(true))); Assert.IsTrue(state.Pending);
        }

        [DataTestMethod]
        [DataRow("ProcessId", 11)]
        [DataRow("FinalProcessId", 11)]
        [DataRow("NativeThreadId", 21)]
        [DataRow("FinalNativeThreadId", 21)]
        [DataRow("Apartment", "MTA")]
        [DataRow("AssemblyMvid", "other-candidate")]
        public void WrongOwnerOrCandidateCannotAuthorizeAnyExport(string key, object value)
        {
            var state = new PairedExportQualificationState(); state.Begin(PathVisibilityDiagnostic.CommandName);
            state.Receive(PathVisibilityDiagnostic.CommandName, Response(true));
            var observed = Observation(); observed[key] = value;
            Assert.ThrowsException<InvalidOperationException>(() => state.VerifyOwner(observed, 10, 20, "candidate", Paths));
            Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component")); Assert.AreEqual(0, state.ExportRequests);
        }

        [TestMethod]
        public void PartialTokenOrReorderedOrMissingAttributeEvidenceCannotAuthorizeExport()
        {
            foreach (string variant in new[] { "token", "order", "attributes" })
            {
                var state = new PairedExportQualificationState(); state.Begin(PathVisibilityDiagnostic.CommandName);
                state.Receive(PathVisibilityDiagnostic.CommandName, Response(true)); var observed = Observation();
                if (variant == "token") observed["EffectiveTokenAfter"] = new Dictionary<string, object> { ["State"] = "DENIED" };
                var rows = (object[])observed["Paths"];
                if (variant == "order") Array.Reverse(rows);
                if (variant == "attributes") ((IDictionary<string, object>)rows[0]).Remove("NativeAttributes");
                Assert.ThrowsException<InvalidOperationException>(() => state.VerifyOwner(observed, 10, 20, "candidate", Paths));
                Assert.ThrowsException<InvalidOperationException>(() => state.Begin("export_component"));
            }
        }
    }
}
