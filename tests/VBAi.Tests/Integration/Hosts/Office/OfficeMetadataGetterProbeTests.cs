using System;
using System.IO;
using System.Runtime.ExceptionServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Optional getter-only baseline on a new disposable Access database; no diagnostic setters or additional saves.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeMetadataGetterProbe"), DoNotParallelize]
    public sealed class OfficeMetadataGetterProbeTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void FreshAccessMetadataGetterContractsReadOnly()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_GETTER_PROBE") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_METADATA_GETTER_PROBE=1 for the explicit read-only getter probe.");
            OfficeVbeFixture fixture = null; Exception failure = null;
            try
            {
                fixture = OfficeVbeFixture.Start("Access");
                fixture.RequireAdapterOnlyCleanup();
                fixture.RecordMetadataGetterProbe("FreshBaseline", true);
            }
            catch (Exception error)
            {
                failure = error;
                if (fixture != null)
                    try { fixture.RecordAdapterFailure(error); }
                    catch (Exception evidence) { failure = new AggregateException("Getter-only probe and failure evidence persistence failed.", error, evidence); }
            }
            finally
            {
                if (fixture != null)
                {
                    try { fixture.Dispose(); }
                    catch (Exception cleanup) { failure = failure == null ? cleanup : new AggregateException("Getter-only probe and normal owned cleanup failed.", failure, cleanup); }
                    foreach (string name in new[] { "metadata-getters-FreshBaseline.json", "adapter-only-progress.json", "qualification.json" })
                    {
                        string file = Path.Combine(fixture.Root, name);
                        try { if (File.Exists(file)) TestContext?.AddResultFile(file); }
                        catch (Exception attach) { failure = failure == null ? attach : new AggregateException("Getter-only probe and evidence attachment failed.", failure, attach); }
                    }
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
