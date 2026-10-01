using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NativeTeardownTraceGateTests
    {
        private sealed class Scope : IDisposable
        {
            internal readonly NativeTeardownTraceGate.Identity Identity;
            internal readonly NativeTeardownTraceGate Gate;
            internal bool Owned = true, Exited, Attached = true;
            internal Scope()
            {
                Identity = new NativeTeardownTraceGate.Identity { ProcessId = 424242, ProcessStartedUtc = DateTime.UtcNow.ToString("o"),
                    Root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")),
                    Executable = @"C:\Program Files\Microsoft Office\EXCEL.EXE", AssemblyPath = @"C:\Owned\VBAi.dll",
                    AssemblyMvid = Guid.NewGuid().ToString("D"), Scenario = "NativeVariantArraysRoundTripWithBoundsAndOneInvocation" };
                Directory.CreateDirectory(Identity.Root);
                Gate = new NativeTeardownTraceGate(Identity, () => Owned, () => Exited, () => Attached);
            }
            internal string Marker(string phase) => new JavaScriptSerializer().Serialize(new { Phase = phase, Identity.ProcessId,
                Identity.ProcessStartedUtc, Identity.Nonce, Identity.AssemblyMvid });
            internal void Arm() => File.WriteAllText(Path.Combine(Identity.Root, "teardown.armed.json"), Marker("Armed"), new UTF8Encoding(false));
            public void Dispose() => Directory.Delete(Identity.Root, true);
        }

        private static IDictionary<string, object> Operation(bool pending, bool uncertain = false) => new Dictionary<string, object> {
            ["Ok"] = true, ["Data"] = new Dictionary<string, object> { ["Query"] = "owned-query", ["Pending"] = pending, ["Uncertain"] = uncertain }
        };

        [TestMethod]
        public void PendingInvocationBlocksCleanupUntilActualTerminalStatusWithoutReplay()
        {
            using (var scope = new Scope())
            {
                int emissions = 0;
                scope.Gate.Command(new { Command = "run_procedure_values" }, () => { emissions++; return Operation(true); });
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
                Assert.IsFalse(File.Exists(Path.Combine(scope.Identity.Root, "teardown.pending.json")));
                scope.Gate.Command(new { Command = "procedure_values_status" }, () => { emissions++; return Operation(false); });
                scope.Arm(); scope.Gate.BeforeCleanup();
                Assert.AreEqual(2, emissions, "Only the requested invocation and status may be emitted.");
                var row = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(Path.Combine(scope.Identity.Root, "teardown-cleanup-intent.json")));
                Assert.AreEqual(2, row["CommandEmissions"]); Assert.AreEqual(1, row["InvocationRequests"]);
                Assert.AreEqual(0, row["PendingQueries"]); Assert.AreEqual(1, row["CleanupAttempts"]);
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
            }
        }

        [TestMethod]
        public void TransportExceptionIsPreservedAndCleanupRefusedEvenAfterLaterSuccessfulRead()
        {
            using (var scope = new Scope())
            {
                var primary = new IOException("Owned delivery uncertain"); int emissions = 0;
                var observed = Assert.ThrowsException<IOException>(() => scope.Gate.Command(new { Command = "run_procedure_values" },
                    () => { emissions++; throw primary; }));
                Assert.AreSame(primary, observed);
                scope.Gate.Command(new { Command = "read_module" }, () => { emissions++; return new Dictionary<string, object> { ["Ok"] = true }; });
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
                Assert.AreEqual(2, emissions); Assert.IsFalse(File.Exists(Path.Combine(scope.Identity.Root, "teardown.pending.json")));
            }
        }

        [TestMethod]
        public void MissingResponseAndNativeUncertainFlagEachRefuseCleanup()
        {
            foreach (bool missing in new[] { false, true }) using (var scope = new Scope())
            {
                scope.Gate.Command(new { Command = "run_procedure_values" }, () => missing ? null : Operation(false, true));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
            }
        }

        [TestMethod]
        public void ArmTimeoutPreservesHostAndPreventsSecondArmingOrCleanupIntent()
        {
            using (var scope = new Scope())
            {
                DateTime now = DateTime.UtcNow; scope.Gate.Clock = () => now;
                scope.Gate.WaitTick = () => now = now.AddSeconds(46);
                Assert.ThrowsException<TimeoutException>(() => scope.Gate.BeforeCleanup());
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
                Assert.IsTrue(File.Exists(Path.Combine(scope.Identity.Root, "teardown.pending.json")));
                Assert.IsFalse(File.Exists(Path.Combine(scope.Identity.Root, "teardown-cleanup-intent.json")));
                Assert.IsFalse(scope.Exited);
            }
        }

        [TestMethod]
        public void MarkerIdentityAndIndependentDebuggerStateAreBothRequired()
        {
            using (var scope = new Scope())
            {
                foreach (string phase in new[] { "Detached", "Stale" })
                    Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.ValidateMarker(scope.Marker(phase), "Armed"));
                string original = scope.Identity.Nonce; scope.Identity.Nonce = Guid.NewGuid().ToString("N");
                string wrong = scope.Marker("Armed"); scope.Identity.Nonce = original;
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.ValidateMarker(wrong, "Armed"));
                scope.Attached = false; scope.Arm();
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
                Assert.IsFalse(File.Exists(Path.Combine(scope.Identity.Root, "teardown-cleanup-intent.json")));
            }
        }

        [TestMethod]
        public void DetachMarkerCannotOverrideAttachedStateButOwnedTerminalExitNeedsNoFurtherNativeAction()
        {
            using (var scope = new Scope())
            {
                scope.Arm(); scope.Gate.BeforeCleanup();
                File.WriteAllText(Path.Combine(scope.Identity.Root, "teardown.detached.json"), scope.Marker("Detached"));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.AfterCleanup());
                scope.Exited = true; scope.Gate.AfterCleanup();
                Assert.IsTrue(File.Exists(Path.Combine(scope.Identity.Root, "teardown-host-terminal.json")));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
            }
        }

        [TestMethod]
        public void ChangedOwnerAndNonProcedureScenariosAreRejectedBeforeArming()
        {
            using (var scope = new Scope())
            {
                scope.Owned = false;
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.BeforeCleanup());
                Assert.IsFalse(File.Exists(Path.Combine(scope.Identity.Root, "teardown.pending.json")));
                scope.Identity.Scenario = "PersonalMacro";
                Assert.ThrowsException<ArgumentException>(() => NativeTeardownTraceGate.ValidateIdentity(scope.Identity));
            }
        }

        [DataTestMethod]
        [DataRow("ProcessId")]
        [DataRow("ProcessStartedUtc")]
        [DataRow("AssemblyMvid")]
        public void StaleMarkerIdentityNeverAuthorizesCleanup(string field)
        {
            using (var scope = new Scope())
            {
                var serializer = new JavaScriptSerializer();
                var marker = serializer.Deserialize<Dictionary<string, object>>(scope.Marker("Armed"));
                marker[field] = field == "ProcessId" ? (object)1 : "stale";
                Assert.ThrowsException<InvalidOperationException>(() => scope.Gate.ValidateMarker(serializer.Serialize(marker), "Armed"));
            }
        }
    }
}
