using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks Outlook observer admission without opening Outlook or publishing a runtime request.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OutlookVbaTestFixtureAddInShutdownTests
    {
        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6), DataRow(7)]
        [DataRow(8), DataRow(9), DataRow(10), DataRow(11), DataRow(12), DataRow(13), DataRow(14), DataRow(15), DataRow(16)]
        public void OnlyVerifiedDisposableMainOutlookAndOneOwnedVbeWindowAreAdmitted(int fault)
        {
            bool main = true, absent = true, verified = true; int pid = 71; IntPtr handle = new IntPtr(12); string root = @"C:\Qualification";
            var frame = new IsolatedTestDesktop.MainWindow { Handle = 13, ProcessId = 71, ThreadId = 7, ClassName = "wndclass_desked_gsk", Visible = true };
            var windows = new[] { frame };
            switch (fault)
            {
                case 1: main = false; break;
                case 2: absent = false; break;
                case 3: verified = false; break;
                case 4: pid = 0; break;
                case 5: handle = IntPtr.Zero; break;
                case 6: root = null; break;
                case 7: root = "relative"; break;
                case 8: windows = null; break;
                case 9: windows = new IsolatedTestDesktop.MainWindow[0]; break;
                case 10: windows = new IsolatedTestDesktop.MainWindow[] { null }; break;
                case 11: frame.ProcessId = 72; break;
                case 12: frame.Handle = 0; break;
                case 13: frame.ThreadId = 0; break;
                case 14: frame.Visible = false; break;
                case 15: frame.ClassName = "another-window"; break;
                case 16: windows = new[] { frame, frame }; break;
            }
            if (fault == 0) Assert.AreSame(frame, OutlookVbaTestFixture.RequireOutlookShutdownScope(main, absent, verified, pid, handle, root, windows));
            else Assert.ThrowsException<InvalidOperationException>(() => OutlookVbaTestFixture.RequireOutlookShutdownScope(main, absent, verified, pid, handle, root, windows));
        }

        [TestMethod]
        public void DisabledObserverDoesNotInspectAHostAndDoesNotChangeTheShutdownVerdict()
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, null);
                var fixture = new OutlookVbaTestFixture { VerifyOwnedProcess = () => Assert.Fail("Disabled observation touched a host.") };
                fixture.ArmOutlookShutdownObservation(null);
                var failures = new List<string> { "Original Quit failed" };
                fixture.ObserveOutlookShutdownTrace(failures); fixture.ObserveOutlookShutdownTrace(failures);
                CollectionAssert.AreEqual(new[] { "Original Quit failed" }, failures);
            }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        [TestMethod]
        public void EnabledObserverRefusesAnUnverifiedFixtureBeforeAnyHostInspection()
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, @"C:\Qualification\Observer");
                var fixture = new OutlookVbaTestFixture { VerifyOwnedProcess = () => Assert.Fail("Unverified observation touched a host.") };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.ArmOutlookShutdownObservation(null));
            }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        [DataTestMethod]
        [DataRow("valid"), DataRow("missing"), DataRow("foreign"), DataRow("receipt failure")]
        public void TraceIsReadOnceAndNeverRehabilitatesTheOriginalQuitFailure(string kind)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                var fixture = new OutlookVbaTestFixture();
                var identity = ShutdownDiagnosticTestData.Identity();
                string root = Path.Combine(directory.Root, "fixture"); Directory.CreateDirectory(root);
                typeof(OutlookVbaTestFixture).GetProperty("Root", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, root);
                Set(fixture, "outlookShutdownIdentity", identity); Set(fixture, "outlookShutdownRoot", directory.Root);
                string output = Path.Combine(directory.Root, "process-71");
                if (kind != "missing")
                {
                    Directory.CreateDirectory(output);
                    AddInShutdownDiagnosticReceiptTests.PublishChain(output, Path.GetFileName(directory.Root));
                }
                if (kind == "foreign") identity.ProductSha256 = new string('B', 64);
                string receipt = Path.Combine(root, "addin-shutdown-observed.json");
                if (kind == "receipt failure") Directory.CreateDirectory(receipt);
                var failures = new List<string> { "Original Application.Quit failed with 0x800706BE" };
                fixture.ObserveOutlookShutdownTrace(failures);
                Assert.AreEqual("Original Application.Quit failed with 0x800706BE", failures[0]);
                Assert.AreEqual(kind == "valid" ? 1 : 2, failures.Count);
                if (kind == "receipt failure") StringAssert.Contains(failures[1], "observation receipt");
                else
                {
                    var observation = (IDictionary<string, object>)ShutdownDiagnosticTestData.Json.DeserializeObject(File.ReadAllText(receipt));
                    Assert.AreEqual(true, observation["OriginalShutdownVerdictUnchanged"]);
                    Assert.AreEqual(false, observation["HostExitProvenByTrace"]);
                    Assert.AreEqual(false, observation["NativeRcwReleaseProven"]);
                    Assert.AreEqual(kind == "valid", observation.ContainsKey("Trace"));
                    Assert.AreEqual(kind != "valid", observation.ContainsKey("TraceError"));
                }
                // Remove the evidence after the first observation. A second read would now fail.
                if (Directory.Exists(output)) Directory.Delete(output, true);
                string before = ShutdownDiagnosticTestData.Json.Serialize(failures);
                fixture.ObserveOutlookShutdownTrace(failures);
                Assert.AreEqual(before, ShutdownDiagnosticTestData.Json.Serialize(failures));
            }
        }

        [TestMethod]
        public void AnAlreadyArmedObserverRefusesDuplicatePublicationBeforeHostInspection()
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, @"C:\Qualification\Observer");
                var fixture = new OutlookVbaTestFixture { VerifyOwnedProcess = () => Assert.Fail("Duplicate publication touched a host.") };
                Set(fixture, "outlookShutdownIdentity", ShutdownDiagnosticTestData.Identity());
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.ArmOutlookShutdownObservation(null)).Message, "already armed");
            }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        private static void Set(OutlookVbaTestFixture fixture, string field, object value) =>
            typeof(OutlookVbaTestFixture).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, value);
    }
}
