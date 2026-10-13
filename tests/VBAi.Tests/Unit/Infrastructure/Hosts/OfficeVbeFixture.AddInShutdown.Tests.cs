using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors Word observer admission and post-cleanup reads without launching Office or changing a lifecycle verdict.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureAddInShutdownTests
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        private static OfficeVbeFixture Fixture() => (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
        private static void Set(OfficeVbeFixture fixture, string field, object value) => typeof(OfficeVbeFixture).GetField(field, Private).SetValue(fixture, value);
        private static object Get(OfficeVbeFixture fixture, string field) => typeof(OfficeVbeFixture).GetField(field, Private).GetValue(fixture);
        private static void Property(OfficeVbeFixture fixture, string property, object value) =>
            typeof(OfficeVbeFixture).GetProperty(property, Private | BindingFlags.Public).SetValue(fixture, value);

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6)]
        [DataRow(7), DataRow(8), DataRow(9), DataRow(10), DataRow(11), DataRow(12), DataRow(13), DataRow(14)]
        public void ScopeRequiresOwnedOriginalHandlesAndTheVerifiedVisibleVbeThread(int fault)
        {
            string kind = "Word", root = @"C:\Qualification";
            bool main = true, owned = true; int pid = 71;
            IntPtr launch = new IntPtr(11), process = new IntPtr(12);
            var vbe = new IsolatedTestDesktop.MainWindow
            { Handle = 13, ProcessId = 71, ThreadId = 7, ClassName = "wndclass_desked_gsk", Visible = true };
            switch (fault)
            {
                case 1: kind = "Excel"; break;
                case 2: main = false; break;
                case 3: owned = false; break;
                case 4: pid = 0; break;
                case 5: pid = -1; break;
                case 6: launch = IntPtr.Zero; break;
                case 7: process = IntPtr.Zero; break;
                case 8: root = null; break;
                case 9: root = " "; break;
                case 10: vbe = null; break;
                case 11: vbe.Handle = 0; break;
                case 12: vbe.ProcessId = 72; break;
                case 13: vbe.ThreadId = 0; break;
                case 14: vbe.Visible = false; break;
            }
            Action run = () => OfficeVbeFixture.RequireWordShutdownObservationScope(kind, main, owned, pid, launch, process, root, vbe);
            if (fault == 0)
            {
                run(); vbe.ClassName = "OpusApp";
                Assert.ThrowsException<InvalidOperationException>(run);
            }
            else Assert.ThrowsException<InvalidOperationException>(run);
        }

        [DataTestMethod, DataRow(null), DataRow(""), DataRow(" "), DataRow("invalid")]
        public void ArmingIsOptInAndRefusesUnownedFixturesBeforeReadingProductPins(string configured)
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, configured);
                var fixture = Fixture();
                Action arm = () => typeof(OfficeVbeFixture).GetMethod("ArmWordShutdownObservation", Private).Invoke(fixture, new object[] { null, null });
                if (string.IsNullOrWhiteSpace(configured)) arm();
                else Assert.IsInstanceOfType(Assert.ThrowsException<TargetInvocationException>(arm).InnerException, typeof(InvalidOperationException));
                Assert.IsNull(Get(fixture, "wordShutdownIdentity"));
            }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void UnarmedPostCleanupObservationDoesNotReadOrWriteEvidence(bool hasRoot)
        {
            var fixture = Fixture();
            if (hasRoot) Property(fixture, "Root", @"C:\Nonexistent\NeverWrite");
            typeof(OfficeVbeFixture).GetMethod("ObserveWordShutdownTrace", Private).Invoke(fixture, null);
            Assert.AreEqual(0, ((List<object>)Get(fixture, "steps")).Count);
        }

        [DataTestMethod]
        [DataRow("valid", false), DataRow("valid", true), DataRow("missing", false), DataRow("foreign", false)]
        public void PostCleanupTraceIsIndependentOfTheOriginalProcessExitVerdict(string kind, bool exited)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                var fixture = Fixture(); var identity = ShutdownDiagnosticTestData.Identity();
                string root = Path.Combine(directory.Root, "fixture"); Directory.CreateDirectory(root);
                Property(fixture, "Root", root);
                Set(fixture, "wordShutdownRoot", directory.Root); Set(fixture, "wordShutdownIdentity", identity);
                var lifecycle = new OfficeOwnedShutdownEvidence(71, identity.ProcessStartedUtc, identity.HostImagePath, "original-handle", identity.ProductMvid);
                lifecycle.Record["ProcessExitObserved"] = exited;
                lifecycle.Record["QuitEntries"] = 1;
                lifecycle.Record["QuitOutcome"] = "RETURNED";
                string before = ShutdownDiagnosticTestData.Json.Serialize(lifecycle.Record);
                Set(fixture, "shutdownEvidence", lifecycle);
                if (kind != "missing")
                {
                    string output = Path.Combine(directory.Root, "process-71"); Directory.CreateDirectory(output);
                    AddInShutdownDiagnosticReceiptTests.PublishChain(output, Path.GetFileName(directory.Root));
                }
                if (kind == "foreign") identity.ProductSha256 = new string('B', 64);
                Action read = () => typeof(OfficeVbeFixture).GetMethod("ObserveWordShutdownTrace", Private).Invoke(fixture, null);
                if (kind == "valid") read();
                else Assert.IsNotNull(Assert.ThrowsException<TargetInvocationException>(read).InnerException);
                Assert.AreEqual(before, ShutdownDiagnosticTestData.Json.Serialize(lifecycle.Record));
                Assert.IsTrue(File.Exists(Path.Combine(root, "addin-shutdown-observed.json")));
                var observation = (IDictionary<string, object>)((List<object>)Get(fixture, "steps"))[0];
                Assert.AreSame(lifecycle.Record, observation["OriginalLifecycle"]);
                Assert.AreEqual(false, observation["HostExitProvenByTrace"]);
                Assert.AreEqual(false, observation["NativeRcwReleaseProven"]);
                Assert.AreEqual(kind == "valid", observation.ContainsKey("Trace"));
                Assert.AreEqual(kind != "valid", observation.ContainsKey("TraceError"));
                Assert.IsNull(Get(fixture, "application")); Assert.IsNull(Get(fixture, "ownedProcess"));
            }
        }
    }
}
