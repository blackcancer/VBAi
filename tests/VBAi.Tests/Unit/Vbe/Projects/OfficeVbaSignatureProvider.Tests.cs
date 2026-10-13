using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks exact VBA registration, trusted module discovery and one-shot owner-thread leases.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbaSignatureProviderTests
    {
        [DataTestMethod]
        [DataRow("CryptSIPDllVerifyIndirectData", "Dll"), DataRow("CryptSIPDllGetSignedDataMsg", "Dll"), DataRow("CryptSIPDllCreateIndirectData", "Dll")]
        [DataRow("CryptSIPDllVerifyIndirectData", "FuncName"), DataRow("CryptSIPDllGetSignedDataMsg", "FuncName"), DataRow("CryptSIPDllCreateIndirectData", "FuncName")]
        public void MissingVbaRegistrationRefusesBeforeTrustOrLoad(string section, string field)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.Registration.Remove(section + "/" + field);
                Assert.IsNull(fixture.Provider.TryAcquire());
                Assert.AreEqual(0, fixture.TrustCalls); Assert.AreEqual(0, fixture.LoadCalls); Assert.AreEqual(0, fixture.FreeCalls);
            }
        }

        [DataTestMethod]
        [DataRow(" "), DataRow("other.dll"), DataRow("sub\\mso.dll"), DataRow("..\\mso.dll"), DataRow("C:mso.dll"), DataRow("\\mso.dll")]
        public void UnexpectedRegisteredLibraryCannotReachNativeLoading(string dll)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.Registration["CryptSIPDllVerifyIndirectData/Dll"] = dll;
                Assert.IsNull(fixture.Provider.TryAcquire()); Assert.AreEqual(0, fixture.TrustCalls); Assert.AreEqual(0, fixture.LoadCalls);
            }
        }

        [DataTestMethod]
        [DataRow("CryptSIPDllVerifyIndirectData"), DataRow("CryptSIPDllGetSignedDataMsg"), DataRow("CryptSIPDllCreateIndirectData")]
        public void ADocumentOrUnexpectedSipFunctionCannotReplaceTheVbaContract(string section)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.Registration[section + "/FuncName"] = "DocumentSignatureOnly";
                Assert.IsNull(fixture.Provider.TryAcquire()); Assert.AreEqual(0, fixture.LoadCalls);
            }
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6), DataRow(7), DataRow(8)]
        public void UnverifiedModulePathsRefuseWithoutLoading(int scenario)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                if (scenario == 0) fixture.LoadedPath = "relative\\mso.dll";
                if (scenario == 1) fixture.LoadedPath = Path.Combine(fixture.Root, "other.dll");
                if (scenario == 2) fixture.LoadedPath = Path.Combine(fixture.Root, "missing", "mso.dll");
                if (scenario == 3) fixture.LoadedPath = null;
                if (scenario == 4) { fixture.LoadedPath = null; fixture.InstallationRoot = "relative"; }
                if (scenario == 5) fixture.Registration["CryptSIPDllGetSignedDataMsg/Dll"] = Path.Combine(fixture.Root, "different", "mso.dll");
                if (scenario == 6) fixture.LoadedPath = "C:mso.dll";
                if (scenario == 7) fixture.LoadedPath = "\\mso.dll";
                if (scenario == 8) { fixture.LoadedPath = null; fixture.InstallationRoot = "C:relative"; }
                Assert.IsNull(fixture.Provider.TryAcquire()); Assert.AreEqual(0, fixture.TrustCalls); Assert.AreEqual(0, fixture.LoadCalls);
            }
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2)]
        public void VerifiedModuleDiscoveryBalancesTheAcquiredReference(int source)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                if (source == 1) { fixture.LoadedPath = null; fixture.Registration["CryptSIPDllVerifyIndirectData/Dll"] = fixture.ModulePath; }
                if (source == 2)
                {
                    fixture.LoadedPath = null; fixture.InstallationRoot = fixture.Root;
                    string path = Path.Combine(fixture.Root, @"root\vfs\ProgramFilesCommonX64\Microsoft Shared\Office16\MSO.DLL");
                    Directory.CreateDirectory(Path.GetDirectoryName(path)); File.Copy(fixture.ModulePath, path); fixture.ModulePath = path;
                }
                var lease = fixture.Provider.TryAcquire();
                Assert.IsNotNull(lease);
                try
                {
                    Assert.AreEqual(fixture.ModulePath, lease.Path); Assert.AreEqual(fixture.ModulePath, fixture.LoadedArgument);
                    Assert.AreEqual(0x1100u, fixture.LoadedFlags); Assert.AreEqual(1, fixture.TrustCalls); Assert.AreEqual(1, fixture.LoadCalls);
                    Assert.AreEqual(3, fixture.ExportNames.Count); Assert.AreEqual(0, fixture.FreeCalls);
                    Assert.ThrowsException<IOException>(() => File.AppendAllText(fixture.ModulePath, "cannot change the verified module"));
                }
                finally { lease.Dispose(); }
                lease.Dispose(); Assert.AreEqual(1, fixture.FreeCalls); Assert.AreEqual(new IntPtr(123), fixture.FreedHandle);
                File.AppendAllText(fixture.ModulePath, "lock released");
            }
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3)]
        public void TrustAndLoadRefusalsNeverAcquireAnExportOrRetry(int scenario)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                if (scenario == 0) fixture.TrustResult = false;
                if (scenario == 1) fixture.TrustFailure = new InvalidOperationException("trust observation failed");
                if (scenario == 2) fixture.LoadResult = IntPtr.Zero;
                if (scenario == 3) fixture.LoadFailure = new InvalidOperationException("load failed");
                if (scenario == 1 || scenario == 3) Assert.ThrowsException<InvalidOperationException>(() => fixture.Provider.TryAcquire());
                else Assert.IsNull(fixture.Provider.TryAcquire());
                Assert.AreEqual(scenario < 2 ? 0 : 1, fixture.LoadCalls); Assert.AreEqual(0, fixture.FreeCalls); Assert.AreEqual(0, fixture.ExportNames.Count);
                File.AppendAllText(fixture.ModulePath, "lock released on refusal");
            }
        }

        [DataTestMethod]
        [DataRow("MsoVBADigSigVerifyIndirectData"), DataRow("MsoVBADigSigGetSignedDataMsg"), DataRow("MsoVBADigSigCreateIndirectData")]
        public void MissingRegisteredExportReleasesExactlyOneReference(string function)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.MissingExport = function;
                Assert.IsNull(fixture.Provider.TryAcquire()); Assert.AreEqual(1, fixture.LoadCalls); Assert.AreEqual(1, fixture.FreeCalls);
                File.AppendAllText(fixture.ModulePath, "lock released");
            }
        }

        [DataTestMethod]
        [DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void UncertainModuleReleaseIsReportedOnceAndCannotLeaveTheFileLocked(bool preparing, bool throwing)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.FreeResult = false;
                if (throwing) fixture.FreeFailure = new InvalidOperationException("release failed");
                if (preparing)
                {
                    fixture.MissingExport = "MsoVBADigSigVerifyIndirectData";
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.Provider.TryAcquire());
                }
                else
                {
                    var lease = fixture.Provider.TryAcquire();
                    Assert.ThrowsException<InvalidOperationException>(() => lease.Dispose()); lease.Dispose();
                }
                Assert.AreEqual(1, fixture.FreeCalls); Assert.AreEqual(1, fixture.LoadCalls);
                File.AppendAllText(fixture.ModulePath, "lock released after uncertain module release");
            }
        }

        [TestMethod]
        public void CrossThreadDisposalLeavesTheLeaseForItsOwner()
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                var lease = fixture.Provider.TryAcquire();
                try
                {
                    Task.Run(() => Assert.ThrowsException<InvalidOperationException>(() => lease.Dispose())).GetAwaiter().GetResult();
                    Assert.AreEqual(0, fixture.FreeCalls);
                }
                finally { lease.Dispose(); }
                Assert.AreEqual(1, fixture.FreeCalls);
            }
        }

        [TestMethod]
        public void ExportObservationFailureStillReleasesTheOwnedModule()
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.ExportFailure = new InvalidOperationException("export observation failed");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Provider.TryAcquire());
                Assert.AreEqual(1, fixture.FreeCalls); Assert.AreEqual(1, fixture.ExportNames.Count);
                File.AppendAllText(fixture.ModulePath, "lock released");
            }
        }

        /// <summary>An unreadable acquired module path releases the sole acquisition and its file lock.</summary>
        [TestMethod]
        public void AcquiredModulePathFailureReleasesWithoutObservingExports()
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.Provider.ReadAcquiredPath = handle => throw new InvalidOperationException("module path unavailable");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Provider.TryAcquire());
                Assert.AreEqual(1, fixture.FreeCalls); Assert.AreEqual(0, fixture.ExportNames.Count);
                File.AppendAllText(fixture.ModulePath, "file lock released");
            }
        }

        [DataTestMethod]
        [DataRow(0), DataRow(1), DataRow(2)]
        public void AcquiredModuleRebindingCannotReachTheVbaFunctions(int scenario)
        {
            using (var fixture = new OfficeVbaSignatureProviderFixture())
            {
                fixture.Provider.ReadAcquiredPath = handle => scenario == 0 ? null : scenario == 1 ? "mso.dll" : Path.Combine(fixture.Root, "different", "mso.dll");
                Assert.IsNull(fixture.Provider.TryAcquire()); Assert.AreEqual(1, fixture.FreeCalls); Assert.AreEqual(0, fixture.ExportNames.Count);
                File.AppendAllText(fixture.ModulePath, "verified file unlocked");
            }
        }
    }
}
