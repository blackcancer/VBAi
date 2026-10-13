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

        /// <summary>Checks all classic admission boundaries without activating Excel or reading another process.</summary>
        /// <param name="variation">Identity, opt-in or retained-image boundary under test.</param>
        [DataTestMethod]
        [DataRow("Disabled")]
        [DataRow("Whitespace")]
        [DataRow("Unowned")]
        [DataRow("NonDurable")]
        [DataRow("MissingFixturePid")]
        [DataRow("WrongRetainedPid")]
        [DataRow("MissingOwnerThread")]
        [DataRow("MissingLoadedStatus")]
        [DataRow("MissingLoadedPid")]
        [DataRow("WrongLoadedPid")]
        [DataRow("StringLoadedPid")]
        [DataRow("MissingImageReader")]
        [DataRow("InvalidRoot")]
        [DataRow("RelativeImage")]
        [DataRow("WrongImage")]
        [DataRow("MissingImage")]
        [DataRow("ImageReadFailure")]
        [DataRow("Valid")]
        public void ClassicObservationRequiresOwnedDurableMatchingHostBeforeImageRead(string variation)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                string configured = variation == "Disabled" ? null : variation == "Whitespace" ? " " :
                    variation == "InvalidRoot" ? @"C:\OutsideScope" : directory.Root;
                IDictionary<string, object> loaded = new Dictionary<string, object> { ["HostProcessId"] = 424242 };
                if (variation == "MissingLoadedStatus") loaded = null;
                if (variation == "MissingLoadedPid") loaded.Clear();
                if (variation == "WrongLoadedPid") loaded["HostProcessId"] = 424243;
                if (variation == "StringLoadedPid") loaded["HostProcessId"] = "424242";
                int reads = 0;
                var originalReadFailure = new IOException("Original retained image read failed.");
                Func<string> readImage = () =>
                {
                    reads++;
                    if (variation == "ImageReadFailure") throw originalReadFailure;
                    return variation == "RelativeImage" ? "EXCEL.EXE" : variation == "WrongImage" ? @"C:\Owned\WINWORD.EXE" :
                        variation == "MissingImage" ? null : @"C:\Owned\EXCEL.EXE";
                };
                if (variation == "MissingImageReader") readImage = null;
                Func<string> capture = () => ExcelVbeFixture.CaptureClassicShutdownImage(configured,
                    variation != "Unowned", variation != "NonDurable", variation == "MissingFixturePid" ? 0 : 424242,
                    variation == "WrongRetainedPid" ? 424243 : 424242, variation == "MissingOwnerThread" ? 0u : 19u,
                    loaded, readImage);
                if (variation == "Disabled" || variation == "Whitespace")
                {
                    Assert.IsNull(capture()); Assert.AreEqual(0, reads);
                }
                else if (variation == "Valid")
                {
                    Assert.AreEqual(@"C:\Owned\EXCEL.EXE", capture()); Assert.AreEqual(1, reads);
                }
                else
                {
                    Exception failure = null;
                    try { capture(); } catch (Exception error) { failure = error; }
                    Assert.IsNotNull(failure, "An invalid classic observation was admitted.");
                    if (variation == "ImageReadFailure") Assert.AreSame(originalReadFailure, failure);
                    else Assert.IsTrue(failure is InvalidOperationException || failure is ArgumentException,
                        "The original admission failure must be preserved: " + failure.GetType().FullName);
                    Assert.AreEqual(variation == "RelativeImage" || variation == "WrongImage" || variation == "MissingImage" ||
                        variation == "ImageReadFailure" ? 1 : 0, reads, "Rejected identity/root must not read the process image.");
                }
            }
        }

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
