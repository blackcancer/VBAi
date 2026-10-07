using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors one-time manifest arming and post-wait trace reads without launching Excel.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ExcelVbeFixtureAddInShutdownTests
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        private static ExcelVbeFixture Fixture() { return (ExcelVbeFixture)Activator.CreateInstance(typeof(ExcelVbeFixture), true); }
        private static void Set(ExcelVbeFixture fixture, string name, object value) { typeof(ExcelVbeFixture).GetField(name, Private).SetValue(fixture, value); }

        [TestMethod]
        public void ManifestPublicationIsCanonicalAndRefusesASecondWriteWithoutReplacingTheFirst()
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                var identity = ShutdownDiagnosticTestData.Identity();
                string path = ExcelVbeFixture.PublishAddInShutdownRequest(directory.Root, identity), original = File.ReadAllText(path);
                AddInShutdownDiagnostic.RequireSameIdentity(identity, AddInShutdownDiagnostic.DecodeRequest(original, Path.GetFileName(directory.Root)));
                Assert.ThrowsException<IOException>(() => ExcelVbeFixture.PublishAddInShutdownRequest(directory.Root, identity));
                Assert.AreEqual(original, File.ReadAllText(path)); Assert.IsTrue(File.Exists(path + ".tmp"));
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ArmingIsDisabledByDefaultAndRefusesUnownedFixturesBeforeAccessingLoadedState(bool enabled)
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, enabled ? "invalid" : null);
                var fixture = Fixture(); var arm = typeof(ExcelVbeFixture).GetMethod("ArmAddInShutdownObservation", Private);
                if (enabled)
                { var error = Assert.ThrowsException<TargetInvocationException>(() => arm.Invoke(fixture, new object[] { 73u, null })); Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException)); }
                else arm.Invoke(fixture, new object[] { 73u, null });
                Assert.IsNull(typeof(ExcelVbeFixture).GetField("addInShutdownIdentity", Private).GetValue(fixture));
            }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        [DataTestMethod, DataRow("valid trace"), DataRow("missing trace"), DataRow("foreign trace")]
        public void PostWaitObservationNeverChangesTheOriginalExitVerdictOrAttemptsNativeCleanup(string kind)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                var fixture = Fixture(); var identity = ShutdownDiagnosticTestData.Identity();
                string fixtureRoot = Path.Combine(directory.Root, "fixture"); Directory.CreateDirectory(fixtureRoot);
                typeof(ExcelVbeFixture).GetProperty("Root", Private | BindingFlags.Public).SetValue(fixture, fixtureRoot);
                typeof(ExcelVbeFixture).GetProperty("ProcessId", Private | BindingFlags.Public).SetValue(fixture, identity.ProcessId);
                Set(fixture, "addInShutdownRoot", directory.Root); Set(fixture, "addInShutdownIdentity", identity);
                string output = Path.Combine(directory.Root, "process-71");
                if (kind != "missing trace") { Directory.CreateDirectory(output); AddInShutdownDiagnosticReceiptTests.PublishChain(output, Path.GetFileName(directory.Root)); }
                if (kind == "foreign trace") identity.ProductSha256 = new string('B', 64);
                var diagnostics = new Dictionary<string, object> { ["Exited"] = false, ["ExitWaitAttempted"] = true, ["ExitWaitReturned"] = true };
                typeof(ExcelVbeFixture).GetMethod("ObserveAddInShutdownTrace", Private).Invoke(fixture, new object[] { diagnostics });
                Assert.AreEqual(false, diagnostics["Exited"]); Assert.AreEqual(true, diagnostics["ExitWaitReturned"]);
                Assert.IsNull(typeof(ExcelVbeFixture).GetField("ownedProcess", Private).GetValue(fixture));
                Assert.IsNull(typeof(ExcelVbeFixture).GetField("application", Private).GetValue(fixture));
                if (kind == "valid trace") { Assert.IsTrue(diagnostics.ContainsKey("AddInShutdownTrace")); Assert.IsFalse(diagnostics.ContainsKey("AddInShutdownTraceError")); }
                else { Assert.IsTrue(diagnostics.ContainsKey("AddInShutdownTraceError")); Assert.IsNotNull(typeof(ExcelVbeFixture).GetField("addInShutdownFailure", Private).GetValue(fixture)); }
            }
        }
    }

    public sealed partial class ExcelVbeFixtureShutdownTests
    {
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void CleanupTraceFailureRetainsTheOriginalDescriptorAndPreservesTheIndependentExitVerdict(bool exited)
        {
            WithFixture((fixture, process, calls) =>
            {
                using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
                {
                    var identity = ShutdownDiagnosticTestData.Identity(); identity.ProcessId = process.Id;
                    Set(fixture, "addInShutdownIdentity", identity); Set(fixture, "addInShutdownRoot", directory.Root);
                    fixture.WaitForOwnedExcelExit = (actual, bound) => { Assert.AreSame(process, actual); Assert.AreEqual(10000, bound); calls.Add("wait"); return exited; };
                    fixture.ReadOwnedExcelExitCode = actual => { calls.Add("code"); return 0; };
                    var failure = Capture(fixture.Dispose); Assert.IsNotNull(failure);
                    Assert.AreSame(process, Get(fixture, "ownedProcess")); Assert.IsFalse(calls.Contains("release"));
                    Assert.AreEqual(exited, fixture.ShutdownDiagnostics["Exited"]);
                    Assert.AreEqual(true, fixture.ShutdownDiagnostics["ProcessHandleRetained"]);
                    Assert.IsTrue(fixture.ShutdownDiagnostics.ContainsKey("AddInShutdownTraceError"));
                    if (!exited) { Assert.IsInstanceOfType(failure, typeof(AggregateException)); StringAssert.Contains(failure.ToString(), "Excel did not exit"); }
                    int count = calls.Count; Assert.ThrowsException<InvalidOperationException>(fixture.Dispose); Assert.AreEqual(count, calls.Count);
                }
            });
        }
    }
}
