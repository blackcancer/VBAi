using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies explicit refusal before the legacy COM setter, without replaying historical partial writes.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), DoNotParallelize]
    public sealed class LegacyHelpMetadataRefusalTests
    {
        public TestContext TestContext { get; set; }
        [STATestMethod] public void AccessLegacyHelpFileRefusedBeforeWrite() { Qualify("Access", "HelpFile"); }
        [STATestMethod] public void AccessLegacyHelpContextRefusedBeforeWrite() { Qualify("Access", "HelpContextID"); }
        [STATestMethod] public void PublisherLegacyHelpFileRefusedBeforeWrite() { Qualify("Publisher", "HelpFile"); }
        [STATestMethod] public void PublisherLegacyHelpContextRefusedBeforeWrite() { Qualify("Publisher", "HelpContextID"); }

        private void Qualify(string host, string property)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_TESTS=1 for disposable host qualification.");
            NativeTestDesktop.Current();
            OfficeVbeFixture fixture = null; Exception failure = null;
            try
            {
                fixture = host == "Publisher" ? OfficeVbeFixture.StartPublisherSerializedQualificationSeed() : OfficeVbeFixture.Start(host);
                fixture.RequireAdapterOnlyCleanup();
                var status = fixture.Data("status");
                Assert.AreEqual(true, status["Connected"]);
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                var before = fixture.Data("project_properties");
                var references = fixture.Data("list_references");
                var sources = ReadSourceManifest(fixture);
                var saved = fixture.Data("project_persistence_status")["ProjectSaved"];
                Assert.AreEqual(true, saved, "The disposable baseline must already be saved.");
                var properties = ((object[])before["Properties"]).Select(VbeBridgeClient.Object).ToArray();
                string type = Convert.ToString(properties.Single(p => Equals(p["Name"], "Type"))["Value"]);
                Assert.IsTrue(type == "100" || type == "vbext_pt_HostProject", "The original project must expose the native host-project type.");
                foreach (string name in new[] { "HelpFile", "HelpContextID" })
                    Assert.AreEqual(VbeProjectComponents.LegacyHelpMetadataSetterStatus, properties.Single(p => Equals(p["Name"], name))["SetterStatus"]);
                string marker = Path.Combine(fixture.Root, "OwnedLegacyRefusal.chm");
                if (property == "HelpFile") File.WriteAllText(marker, "Inert owned path marker; no help content is executed.");
                object value = property == "HelpFile" ? (object)marker : 322;
                fixture.RecordAdapterStage("SingleLegacyHelpMetadataRequestStarting", new
                {
                    Host = host,
                    Property = property,
                    ExpectedProjectVersion = before["Version"],
                    ProcessId = fixture.ProcessId,
                    ExpectedMvid = status["AssemblyModuleVersionId"],
                    LegacyRequestLimit = 1,
                    ExpectedSetterEntries = 0,
                    SaveAllowed = false,
                    ReopenAllowed = false,
                    RetryAllowed = false
                });
                fixture.NativeExecutionUnsettled = true;
                var response = fixture.Response("set_project_property", "Property", property, "Value", value,
                    "ExpectedProjectVersion", before["Version"], "ExpectedMode", 2);
                Assert.AreEqual(false, response["Ok"], "The legacy route must explicitly fail before mutation.");
                string error = Convert.ToString(response["Error"]);
                StringAssert.Contains(error, "Legacy COM " + property + " write is unsupported");
                StringAssert.Contains(error, "no setter was invoked");
                StringAssert.Contains(error, "read_project_general"); StringAssert.Contains(error, "set_project_general");
                Assert.IsFalse(error.Contains("Scalar property phase:"), "A pre-write policy refusal cannot claim setter entry or retention failure.");
                Assert.AreEqual(before["Version"], fixture.Data("project_properties")["Version"]);
                Assert.AreEqual(saved, fixture.Data("project_persistence_status")["ProjectSaved"]);
                CollectionAssert.AreEqual(sources, ReadSourceManifest(fixture));
                OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(references, fixture.Data("list_references"));
                var finalStatus = fixture.Data("status");
                Assert.AreEqual(true, finalStatus["Connected"]);
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(finalStatus["HostProcessId"]));
                Assert.AreEqual(status["AssemblyModuleVersionId"], finalStatus["AssemblyModuleVersionId"]);
                var finalOwner = fixture.RecordAdapterObservation("AfterLegacyHelpMetadataRefusal");
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(finalOwner["ProcessId"]));
                Assert.AreEqual(fixture.DocumentPath, finalOwner["DocumentPath"]);
                fixture.RecordAdapterStage("LegacyHelpMetadataRefusalVerified", new
                {
                    Host = host,
                    Property = property,
                    OriginalResponse = response,
                    ProjectVersionUnchanged = true,
                    SourceUnchanged = true,
                    ReferencesUnchanged = true,
                    SavedStateUnchanged = true,
                    ExpectedSetterEntries = 0,
                    PreSetterRefusalVerified = true,
                    SetterEntryObservation = "NotInstrumented",
                    SaveEntries = 0,
                    FreshReopenEntries = 0,
                    AlternativeInvoked = false,
                    NormalOriginalExitRequired = true,
                    RetryAllowed = false
                });
                fixture.NativeExecutionUnsettled = false;
            }
            catch (Exception error)
            {
                failure = error;
                if (fixture != null)
                    try { fixture.RecordAdapterFailure(error); }
                    catch (Exception receipt) { failure = new AggregateException(failure, receipt); }
            }
            finally
            {
                if (fixture != null)
                {
                    try { fixture.Dispose(); }
                    catch (Exception cleanup) { failure = failure == null ? cleanup : new AggregateException(failure, cleanup); }
                    foreach (string name in new[] { "adapter-only-progress.json", "report.json", "host-shutdown.json" })
                    {
                        string path = Path.Combine(fixture.Root, name);
                        if (File.Exists(path)) TestContext?.AddResultFile(path);
                    }
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private static string[] ReadSourceManifest(OfficeVbeFixture fixture) => fixture.Items("list_modules")
            .OrderBy(item => Convert.ToString(item["Name"]), StringComparer.Ordinal)
            .Select(item => Convert.ToString(item["Name"]) + "|" + Convert.ToString(item["Type"]) + "|" +
                Convert.ToString(fixture.Data("read_module", "Module", item["Name"])["Sha256"]))
            .ToArray();
    }
}
