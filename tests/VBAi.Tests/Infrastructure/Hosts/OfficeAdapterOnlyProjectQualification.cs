using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies disposable metadata/reference changes with one adapter save and independent fresh-process readback.</summary>
    internal static class OfficeAdapterOnlyProjectQualification
    {
        private static readonly string[] Modules = { "ProjectPersistenceModule", "ProjectPersistenceClass" };

        /// <summary>Prepares a baseline before final mutations, preserving original adapter and independent cleanup failures.</summary>
        internal static void Run(string host, string scenario, Action<OfficeVbeFixture> baseline,
            Action<OfficeVbeFixture> mutate, Action<OfficeVbeFixture> verify, TestContext context)
        {
            OfficeVbeFixture fixture = null;
            Exception failure = null, originalOutcome = null;
            try
            {
                fixture = OfficeVbeFixture.Start(host);
                fixture.RequireAdapterOnlyCleanup();
                fixture.RecordAdapterStage("ProjectQualificationScope", new { Scenario = scenario,
                    NativeMutationRetryAllowed = false, CompileAllowed = false, PostAdapterHelperSaveAllowed = false });
                var startup = fixture.RecordAdapterObservation("ProjectQualificationStartup");
                if (host == "Access") Assert.IsTrue(Convert.ToString(startup["ApplicationVersion"]).StartsWith("16.", StringComparison.Ordinal), "Access 16 is required.");
                foreach (string module in Modules)
                {
                    fixture.Data(module.EndsWith("Class", StringComparison.Ordinal) ? "create_class" : "create_module", "Module", module, "ExpectedMode", 2);
                    ReplaceMarker(fixture, module, "baseline-" + Guid.NewGuid().ToString("N"));
                }
                baseline?.Invoke(fixture);
                fixture.SaveAdapterBaseline(Modules);
                Assert.AreEqual(true, fixture.Data("project_persistence_status")["ProjectSaved"]);
                var baselineHashes = ReadHashes(fixture);
                fixture.RecordAdapterStage("ProjectBaseline", new { SourceHashes = baselineHashes,
                    Properties = fixture.Data("project_properties"), References = fixture.Data("list_references") });
                if (scenario.StartsWith("Metadata.", StringComparison.Ordinal)) fixture.RecordMetadataGetterProbe("BeforeExistingMutation");
                mutate(fixture);
                if (scenario.StartsWith("Metadata.", StringComparison.Ordinal)) fixture.RecordMetadataGetterProbe("AfterExistingMutation");
                CollectionAssert.AreEqual(baselineHashes, ReadHashes(fixture), "Metadata/reference mutations must preserve synthetic source.");
                verify(fixture);
                // A new source edit guarantees the native Save is exercised even when the host writes metadata immediately.
                foreach (string module in Modules) ReplaceMarker(fixture, module, "pending-" + Guid.NewGuid().ToString("N"));
                var expectedHashes = ReadHashes(fixture);
                for (int index = 0; index < Modules.Length; index++) Assert.AreNotEqual(baselineHashes[index], expectedHashes[index]);
                var expectedReferences = fixture.Data("list_references");
                var expectedMetadata = ReadMetadata(fixture);
                fixture.RecordAdapterStage("FinalPendingProjectState", new { SourceHashes = expectedHashes,
                    Properties = fixture.Data("project_properties"), Metadata = expectedMetadata, References = expectedReferences });
                fixture.Data("select_code", "Module", Modules[0], "StartLine", 1,
                    "ExpectedSha256", fixture.Data("read_module", "Module", Modules[0])["Sha256"]);
                var selected = fixture.RecordAdapterObservation("BeforeProjectAdapter");
                Assert.AreEqual(Modules[0], selected["ActiveCodeComponent"]);
                Assert.AreEqual(false, selected["ProjectSaved"]);
                int originalPid = fixture.ProcessId;
                string path = fixture.DocumentPath, project = fixture.Project;
                string version = (string)fixture.Data("project_properties")["Version"];
                fixture.RecordAdapterStage("SingleProjectAdapterSaveStarting", new { Scenario = scenario, Project = project,
                    HostPath = path, ProcessId = originalPid, ExpectedProjectVersion = version,
                    ExpectedMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D") });
                var response = fixture.Response("save_host_document", "ExpectedHostPath", path, "ExpectedProjectVersion", version);
                fixture.RecordAdapterStage("OriginalProjectAdapterResponse", response);
                try { AssertOriginalResponse(response, host); }
                catch (Exception error) { originalOutcome = error; }
                fixture.ObserveAdapterOutcome(response);
                if (scenario.StartsWith("Metadata.", StringComparison.Ordinal)) fixture.RecordMetadataGetterProbe("AfterExistingSave");
                CollectionAssert.AreEqual(expectedHashes, ReadHashes(fixture), "Save must preserve every synthetic source hash.");
                AssertReferencesEqual(expectedReferences, fixture.Data("list_references"));
                CollectionAssert.AreEqual(expectedMetadata, ReadMetadata(fixture), "Save must preserve all project name/description/help metadata.");
                verify(fixture);
                fixture.ReopenFromDisk(() => {
                    using (var stream = File.OpenRead(path))
                    using (var hash = SHA256.Create())
                        fixture.RecordAdapterStage("ClosedNativeFileEvidence", new { Path = path, Bytes = stream.Length,
                            Sha256 = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant(),
                            PreviousProcessId = originalPid, HostExitedBeforeHash = true });
                });
                Assert.AreEqual(path, fixture.DocumentPath);
                Assert.AreEqual(project, fixture.Project);
                fixture.RecordAdapterObservation("ProjectFreshDiskReopen");
                if (scenario.StartsWith("Metadata.", StringComparison.Ordinal)) fixture.RecordMetadataGetterProbe("FreshDiskReopen");
                fixture.RecordAdapterStage("ProjectFreshDiskState", new { SourceHashes = ReadHashes(fixture),
                    Properties = fixture.Data("project_properties"), References = fixture.Data("list_references"),
                    PreviousProcessId = originalPid, ReopenProcessId = fixture.ProcessId });
                CollectionAssert.AreEqual(expectedHashes, ReadHashes(fixture), "Fresh disk readback must retain both final module/class hashes.");
                AssertReferencesEqual(expectedReferences, fixture.Data("list_references"));
                CollectionAssert.AreEqual(expectedMetadata, ReadMetadata(fixture), "Fresh native disk readback changed project metadata.");
                verify(fixture);
                if (originalOutcome != null) ExceptionDispatchInfo.Capture(originalOutcome).Throw();
            }
            catch (Exception error)
            {
                failure = originalOutcome != null && !ReferenceEquals(originalOutcome, error)
                    ? new AggregateException("The original adapter outcome and subsequent readback both failed.", originalOutcome, error) : error;
                if (fixture != null)
                    try { fixture.RecordAdapterFailure(failure); }
                    catch (Exception evidenceError) { failure = new AggregateException("Project qualification and durable evidence both failed.", failure, evidenceError); }
            }
            finally
            {
                if (fixture != null)
                {
                    try { fixture.Dispose(); }
                    catch (Exception cleanup) { failure = failure == null ? cleanup : new AggregateException("Project qualification and owned cleanup both failed.", failure, cleanup); }
                    foreach (string file in new[] { "adapter-only-progress.json", "qualification.json" })
                    {
                        string report = Path.Combine(fixture.Root, file);
                        try { if (File.Exists(report)) context?.AddResultFile(report); }
                        catch (Exception attach) { failure = failure == null ? attach : new AggregateException("Project qualification and report attachment both failed.", failure, attach); }
                    }
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        /// <summary>Compares the full ordered native reference state, including its production revision.</summary>
        internal static void AssertReferencesEqual(IDictionary<string, object> expected, IDictionary<string, object> actual)
        {
            Assert.AreEqual(expected["Version"], actual["Version"], "Native reference revision must persist exactly.");
            var json = new JavaScriptSerializer();
            var before = ((object[])expected["References"]).Select(json.Serialize).ToArray();
            var after = ((object[])actual["References"]).Select(json.Serialize).ToArray();
            CollectionAssert.AreEqual(before, after, "Ordered native GUID/version/name/path/broken/built-in inventory differs.");
        }

        /// <summary>Reads one exposed scalar metadata property, preserving unavailable/error states as failures.</summary>
        internal static object ReadProperty(OfficeVbeFixture fixture, string name)
        {
            var descriptors = ((object[])fixture.Data("project_properties")["Properties"]).Select(VbeBridgeClient.Object);
            var property = descriptors.Single(item => string.Equals(Convert.ToString(item["Name"]), name, StringComparison.OrdinalIgnoreCase));
            Assert.AreEqual("scalar", property["Kind"], "Metadata must have a safe scalar getter.");
            Assert.IsNull(property["Error"], Convert.ToString(property["Error"]));
            return property["Value"];
        }

        /// <summary>Uses only the current project revision for one requested metadata mutation.</summary>
        internal static void SetProperty(OfficeVbeFixture fixture, string name, object value)
        {
            fixture.RecordAdapterStage("ProjectMetadataMutationStarting", new { Property = name, Value = value });
            fixture.Data("set_project_property", "Property", name, "Value", value,
                "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
            Assert.AreEqual(Convert.ToString(value, CultureInfo.InvariantCulture),
                Convert.ToString(ReadProperty(fixture, name), CultureInfo.InvariantCulture), "The immediate metadata mutation was not retained.");
        }

        private static string[] ReadMetadata(OfficeVbeFixture fixture)
        {
            var properties = ((object[])fixture.Data("project_properties")["Properties"]).Select(VbeBridgeClient.Object).ToArray();
            return new[] { "Name", "Description", "HelpFile", "HelpContextID" }.Select(name => {
                var property = properties.Single(item => string.Equals(Convert.ToString(item["Name"]), name, StringComparison.OrdinalIgnoreCase));
                Assert.AreEqual("scalar", property["Kind"]); Assert.IsNull(property["Error"], Convert.ToString(property["Error"]));
                return name + "=" + Convert.ToString(property["Value"], CultureInfo.InvariantCulture);
            }).ToArray();
        }

        private static void AssertOriginalResponse(IDictionary<string, object> response, string host)
        {
            Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"]));
            var value = VbeBridgeClient.Object(response["Data"]);
            Assert.AreEqual(true, value["Verified"], "Later persistence cannot promote an original uncertain response.");
            Assert.AreEqual(false, value["Uncertain"]); Assert.AreEqual(true, value["MutationInvoked"]);
            Assert.AreEqual(false, value["PersistenceReopenVerified"]);
            if (host == "Access") Assert.IsNull(value["HostSaved"]);
        }

        private static string[] ReadHashes(OfficeVbeFixture fixture) => Modules.Select(module =>
            (string)fixture.Data("read_module", "Module", module)["Sha256"]).ToArray();

        private static void ReplaceMarker(OfficeVbeFixture fixture, string module, string marker)
        {
            var before = fixture.Data("read_module", "Module", module);
            fixture.Data("replace_lines", "Module", module, "StartLine", 1,
                "Count", fixture.Data("component_properties", "Module", module)["CodeLines"],
                "ExpectedSha256", before["Sha256"], "Text", "Option Explicit\r\n' Synthetic project qualification " + marker + "\r\n");
        }
    }
}
