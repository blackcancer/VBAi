using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies one native adapter invocation and fresh disk readback without intervening compilation.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), DoNotParallelize]
    public sealed class OfficeAdapterOnlyQualificationTests
    {
        /// <summary>Compares active-module-only and module-plus-class saves in separate Access 16 processes/files.</summary>
        [STATestMethod]
        public void Access16ActiveModuleOnlyAdapterSaveReopen() { Qualify("Access", false, false); }

        /// <summary>Qualifies a separate Access 16 file with pending edits in both a module and a class.</summary>
        [STATestMethod]
        public void Access16ModuleAndClassAdapterSaveReopen() { Qualify("Access", true, false); }

        /// <summary>Retains application-owner diagnostics before testing the pathless Publisher association.</summary>
        [STATestMethod]
        public void PublisherAdapterOnlySaveReopen() { Qualify("Publisher", true, true); }

        /// <summary>Qualifies code and UserForm saves from an audited existing serialized Publisher publication.</summary>
        [STATestMethod] public void PublisherSerializedAdapterOnlySaveReopen() { Qualify("Publisher", true, true, true); }

        /// <summary>Requires native Word form creation and a verified adapter-only document round-trip.</summary>
        [STATestMethod]
        public void WordAdapterOnlySaveReopen() { Qualify("Word", true, true); }

        /// <summary>Requires a fresh native PowerPoint form and verified adapter-only save/reopen.</summary>
        [STATestMethod]
        public void PowerPointAdapterOnlySaveReopen() { Qualify("PowerPoint", true, true); }

        /// <summary>Creates a saved baseline, dirties only synthetic code, and preserves the original save outcome.</summary>
        private static void Qualify(string host, bool editClass, bool includeForm, bool serializedPublisherSeed = false)
        {
            var fixture = serializedPublisherSeed ? OfficeVbeFixture.StartPublisherSerializedQualificationSeed() : OfficeVbeFixture.Start(host);
            Exception trialError = null, originalOutcomeError = null;
            try
            {
                fixture.RequireAdapterOnlyCleanup();
                try
                {
                    var startup = fixture.RecordAdapterObservation("BeforeBaselineCreation");
                    if (host == "Access") Assert.IsTrue(Convert.ToString(startup["ApplicationVersion"]).StartsWith("16.", StringComparison.Ordinal), "Q-012 requires Access 16, not an older CurVer server.");
                    var names = new[] { "AdapterOnlyModule", "AdapterOnlyClass" };
                    foreach (string name in names)
                    {
                        fixture.Data(name.EndsWith("Class", StringComparison.Ordinal) ? "create_class" : "create_module", "Module", name, "ExpectedMode", 2);
                        ReplaceWithMarker(fixture, name, "baseline-" + Guid.NewGuid().ToString("N"));
                    }
                    if (includeForm)
                    {
                        fixture.RecordAdapterStage("CreateFormStarting", new { Form = "AdapterOnlyForm" });
                        fixture.Data("create_form", "Form", "AdapterOnlyForm");
                        fixture.RecordAdapterStage("CreateFormAnswered", new { Form = "AdapterOnlyForm" });
                        fixture.Data("add_form_control", "Form", "AdapterOnlyForm", "Control", "MarkerLabel", "ControlType", "Forms.Label.1",
                            "Left", 12, "Top", 12, "Width", 140, "Height", 24, "Caption", "Adapter-only baseline",
                            "ExpectedFormVersion", fixture.Data("form_state", "Form", "AdapterOnlyForm")["Version"]);
                    }
                    fixture.SaveAdapterBaseline(names);
                    Assert.AreEqual(true, fixture.Data("project_persistence_status")["ProjectSaved"], "The baseline must be saved before the differential pending edit.");
                    var baselineHashes = ReadHashes(fixture, names);
                    fixture.RecordAdapterStage("BaselineSourceHashes", baselineHashes);
                    ReplaceWithMarker(fixture, names[0], "pending-module-" + Guid.NewGuid().ToString("N"));
                    if (editClass) ReplaceWithMarker(fixture, names[1], "pending-class-" + Guid.NewGuid().ToString("N"));
                    var expectedHashes = ReadHashes(fixture, names);
                    fixture.RecordAdapterStage("PendingSourceHashes", new { EditClass = editClass, Hashes = expectedHashes });
                    Assert.AreNotEqual(baselineHashes[names[0]], expectedHashes[names[0]]);
                    if (editClass) Assert.AreNotEqual(baselineHashes[names[1]], expectedHashes[names[1]]);
                    else Assert.AreEqual(baselineHashes[names[1]], expectedHashes[names[1]], "The single-object trial must leave class source unchanged.");
                    var selected = fixture.Data("read_module", "Module", names[0]);
                    fixture.Data("select_code", "Module", names[0], "StartLine", 1, "ExpectedSha256", selected["Sha256"]);
                    var before = fixture.RecordAdapterObservation("BeforeAdapter");
                    Assert.AreEqual(names[0], before["ActiveCodeComponent"]);
                    Assert.AreEqual(false, before["ProjectSaved"], "Fresh pending edits are required for adapter qualification.");
                    string expectedVersion = (string)fixture.Data("project_properties")["Version"];
                    fixture.RecordAdapterStage("SingleAdapterSaveStarting", new { Project = fixture.Project,
                        ExpectedHostPath = fixture.DocumentPath, ExpectedProjectVersion = expectedVersion, RetryAllowed = false });
                    var original = fixture.Response("save_host_document", "ExpectedHostPath", fixture.DocumentPath,
                        "ExpectedProjectVersion", expectedVersion);
                    try { AssertVerifiedAdapterOutcome(original, host); }
                    catch (Exception error) { originalOutcomeError = error; fixture.RecordAdapterFailure(error); }
                    fixture.ObserveAdapterOutcome(original);
                    foreach (var source in expectedHashes)
                        Assert.AreEqual(source.Value, fixture.Data("read_module", "Module", source.Key)["Sha256"], "Live source changed during the one save invocation.");

                    // A failed/uncertain original response remains a failed test even if disk readback succeeds.
                    // Access closes loaded objects with acSaveNo; there is no compile or helper Save here.
                    fixture.ReopenFromDisk();
                    fixture.RecordAdapterObservation("FreshDiskReopen");
                    var actualHashes = ReadHashes(fixture, names);
                    fixture.RecordAdapterStage("FreshDiskSourceHashes", actualHashes);
                    foreach (var source in expectedHashes)
                        Assert.AreEqual(source.Value, actualHashes[source.Key], "Adapter-only persisted source differs: " + source.Key);
                    if (includeForm)
                    {
                        var form = fixture.Data("form_state", "Form", "AdapterOnlyForm");
                        var label = ((object[])form["Controls"]).Select(VbeBridgeClient.Object).Single(c => (string)c["Name"] == "MarkerLabel");
                        Assert.AreEqual("Adapter-only baseline", label["Caption"]);
                    }
                    if (originalOutcomeError != null) ExceptionDispatchInfo.Capture(originalOutcomeError).Throw();
                    fixture.FlushAdapterEvidence();
                }
                catch (Exception error)
                {
                    trialError = originalOutcomeError != null && !ReferenceEquals(error, originalOutcomeError)
                        ? new AggregateException("Original adapter result and subsequent adapter-only verification both failed.", originalOutcomeError, error)
                        : error;
                    fixture.RecordAdapterFailure(trialError);
                }
            }
            finally
            {
                try { fixture.Dispose(); }
                catch (Exception cleanupError)
                {
                    trialError = trialError == null ? cleanupError : new AggregateException(
                        "Adapter-only trial and owned host cleanup both failed.", trialError, cleanupError);
                }
            }
            if (trialError != null) ExceptionDispatchInfo.Capture(trialError).Throw();
        }

        /// <summary>Validates only the original response; later observations never promote an uncertain save.</summary>
        private static void AssertVerifiedAdapterOutcome(IDictionary<string, object> original, string host)
        {
            Assert.AreEqual(true, original["Ok"], Convert.ToString(original["Error"]));
            var result = VbeBridgeClient.Object(original["Data"]);
            Assert.AreEqual(true, result["Verified"], "The original adapter result is unverified; subsequent persistence cannot turn it into a verified response.");
            Assert.AreEqual(false, result["Uncertain"]);
            Assert.AreEqual(true, result["MutationInvoked"]);
            Assert.AreEqual(false, result["PersistenceReopenVerified"]);
            if (host == "Access") Assert.IsNull(result["HostSaved"]);
        }

        /// <summary>Replaces only synthetic code under the current source hash; never executes it.</summary>
        private static void ReplaceWithMarker(OfficeVbeFixture fixture, string name, string marker)
        {
            var before = fixture.Data("read_module", "Module", name);
            int lines = Convert.ToInt32(fixture.Data("component_properties", "Module", name)["CodeLines"]);
            fixture.Data("replace_lines", "Module", name, "StartLine", 1, "Count", lines,
                "ExpectedSha256", before["Sha256"], "Text", "Option Explicit\r\n' Synthetic adapter-only qualification: " + marker + "\r\n");
        }

        /// <summary>Captures exact module and class source hashes for independent live and disk comparisons.</summary>
        private static IDictionary<string, string> ReadHashes(OfficeVbeFixture fixture, string[] names)
        {
            return names.ToDictionary(name => name, name => (string)fixture.Data("read_module", "Module", name)["Sha256"]);
        }
    }
}
