using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.ExceptionServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Records real Access identities without imposing document-wrapper stability.</summary>
    [TestClass, TestCategory("Office"), TestCategory("AccessReadOnlyIdentity"), DoNotParallelize]
    public sealed class AccessReadOnlyIdentityQualificationTests
    {
        public TestContext TestContext { get; set; }

        /// <summary>Reads successive CurrentProject and mapped/selected VBProject identities before all edits.</summary>
        [STATestMethod]
        public void Access16CurrentProjectIdentityBeforeAnyMutation()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_ACCESS_IDENTITY_PROBE") != "1")
                Assert.Inconclusive("Set VBAi_RUN_ACCESS_IDENTITY_PROBE=1 for this explicitly authorized read-only probe.");
            OfficeVbeFixture fixture = null;
            Exception failure = null;
            try
            {
                string executable;
                using (var registration = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE"))
                    executable = registration?.GetValue(null) as string;
                Assert.IsFalse(string.IsNullOrWhiteSpace(executable), "Installed x64 Access App Paths registration is required.");
                Assert.IsTrue(Path.IsPathRooted(executable) && File.Exists(executable), "The installed Access path must exist before activation.");
                fixture = OfficeVbeFixture.StartAccessIdentityProbe();
                fixture.RecordAccessIdentityProbe(executable);
            }
            catch (Exception error) { failure = error; }
            finally
            {
                if (fixture != null)
                {
                    try { fixture.Dispose(); }
                    catch (Exception cleanup) { failure = failure == null ? cleanup : new AggregateException("Identity probe and owned cleanup both failed.", failure, cleanup); }
                    foreach (string name in new[] { "access-identity-probe.json", "adapter-only-progress.json", "qualification.json" })
                    {
                        string path = Path.Combine(fixture.Root, name);
                        try { if (File.Exists(path)) TestContext?.AddResultFile(path); }
                        catch (Exception attach) { failure = failure == null ? attach : new AggregateException("Identity probe and evidence attachment both failed.", failure, attach); }
                    }
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
