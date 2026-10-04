using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies explicit native General writes without invoking a COM metadata setter or executing VBA.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), TestCategory("PublisherGeneral"), DoNotParallelize]
    public sealed class PublisherGeneralQualificationTests
    {
        private static readonly string[] Modules = { "PublisherGeneralModule", "PublisherGeneralClass" };
        private static readonly string[] NativeFields = { "Name", "Description", "HelpFile", "HelpContextText", "ConditionalCompilation" };
        public TestContext TestContext { get; set; }

        /// <summary>Persists an exact Unicode path to an inert owned marker; this does not qualify compiled help content.</summary>
        [STATestMethod]
        public void PublisherNativeHelpFileSaveReopen() { Qualify("HelpFile"); }

        /// <summary>Performs one initial native Int32 write and verifies it independently after one adapter Save and fresh reopen.</summary>
        [STATestMethod]
        public void PublisherNativeHelpContextSaveReopen() { Qualify("HelpContextID"); }

        /// <summary>Qualifies exact accented text representable by the observed ANSI General control, including fresh disk persistence.</summary>
        [STATestMethod]
        public void PublisherNativeAnsiHelpFileSaveReopen() { Qualify("HelpFile", "Owned help été_ß.chm"); }

        /// <summary>Checks no-write refusal for unsupported text; retained modal cleanup remains a campaign failure.</summary>
        [STATestMethod]
        public void PublisherNativeUnicodeHelpFileRefusedBeforeWrite()
        { Qualify("HelpFile", "Owned help été_日本_ß.chm", true); }

        internal void Qualify(string property, string helpFileMarker = "Owned help été_日本_ß.chm",
            bool expectEncodingRefusal = false, string host = "Publisher")
        {
            RequirePrivateHostOptIn(host);
            OfficeVbeFixture fixture = null;
            Exception failure = null;
            byte[] refusalBaselineBytes = null;
            string refusalDocumentPath = null;
            bool refusalVerified = false;
            try
            {
                fixture = host == "Publisher" ? OfficeVbeFixture.StartPublisherSerializedQualificationSeed() : OfficeVbeFixture.Start(host);
                fixture.RequireAdapterOnlyCleanup();
                Assert.AreEqual(host, fixture.Kind, "The launched host must match the selected native qualification.");
                fixture.RecordAdapterStage(host + "GeneralQualificationScope", new { Host = fixture.Kind, Property = property,
                    NativeWriteInvocationLimit = 1, AdapterSaveInvocationLimit = 1, ComMetadataSetterInvoked = false,
                    MacroExecutionAllowed = false, HelpInvoked = false, CompileAllowed = false,
                    PostAdapterHelperSaveAllowed = false, RetryAllowed = false });
                fixture.RecordAdapterStage(host + "GeneralEncodingScenario", new { Host = fixture.Kind, ExpectedEncodingRefusal = expectEncodingRefusal,
                    HelpFileMarker = property == "HelpFile" ? helpFileMarker : null,
                    RetainedModalIsNormalCleanupSuccess = false, UnsupportedUnicodePersistenceQualified = false });
                RequireBridge(fixture);
                if (expectEncodingRefusal)
                {
                    Assert.AreEqual("Publisher", host);
                    Assert.IsTrue(fixture.PublisherSerializedSeed, "The refusal case requires the audited saved serialized seed, without preparation writes.");
                    refusalDocumentPath = fixture.DocumentPath;
                    refusalBaselineBytes = File.ReadAllBytes(OfficeVbeFixture.PublisherSeedPath);
                    string sourceSha;
                    using (var hash = SHA256.Create()) sourceSha = BitConverter.ToString(hash.ComputeHash(refusalBaselineBytes)).Replace("-", "");
                    OfficeVbeFixture.RequirePublisherSeedBytes(OfficeVbeFixture.PublisherSeedPath, refusalBaselineBytes.LongLength, sourceSha);
                    fixture.RecordAdapterStage("PublisherEncodingRefusalClosedSeedBaseline", new {
                        Source = OfficeVbeFixture.PublisherSeedPath, SourceSha256 = sourceSha, SourceBytes = refusalBaselineBytes.LongLength,
                        OwnedCopy = refusalDocumentPath, BaselineFromAuditedPreOpenSeedCopy = true,
                        LivePublicationBytesRead = false, SourceWrites = 0, PreparationSaveAttempts = 0 });
                }
                else
                {
                    foreach (string name in Modules)
                    {
                        if (!fixture.PublisherSerializedSeed)
                            fixture.Data(name.EndsWith("Class", StringComparison.Ordinal) ? "create_class" : "create_module", "Module", name, "ExpectedMode", 2);
                        ReplaceMarker(fixture, name, "baseline-" + Guid.NewGuid().ToString("N"));
                    }
                    SelectOwnedCode(fixture);
                    fixture.SaveAdapterBaseline(Modules);
                }
                Assert.AreEqual(true, fixture.Data("project_persistence_status")["ProjectSaved"]);
                var baselineHashes = ReadHashes(fixture);
                var baselineReferences = fixture.Data("list_references");
                var before = ReadGeneral(fixture, "BeforeInitialNativeWrite");
                object expected;
                if (property == "HelpFile")
                {
                    string marker = Path.GetFullPath(Path.Combine(fixture.Root, helpFileMarker));
                    Assert.AreEqual(Path.GetFullPath(fixture.Root), Path.GetDirectoryName(marker), true);
                    Assert.IsFalse(File.Exists(marker), "The owned inert HelpFile marker must be fresh.");
                    File.WriteAllText(marker, "VBAi disposable path-storage marker; not compiled CHM content.\r\n", new UTF8Encoding(false));
                    expected = marker;
                    fixture.RecordAdapterStage("OwnedExactHelpFileMarker", new { Path = marker,
                        CompiledHelpContent = false, ContentQualification = false, HelpInvoked = false });
                }
                else
                {
                    int original;
                    Assert.IsTrue(int.TryParse(Text(before, "HelpContextText"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out original), "The native baseline HelpContextID must be an exact integer.");
                    expected = original == 321 ? 322 : 321;
                }
                string nativeField = property == "HelpFile" ? "HelpFile" : "HelpContextText";
                Assert.AreNotEqual(Text(before, nativeField), Convert.ToString(expected, CultureInfo.InvariantCulture));
                if (expectEncodingRefusal)
                {
                    Assert.AreEqual("HelpFile", property);
                    string refusalProjectVersion = (string)fixture.Data("project_properties")["Version"];
                    AssertEncodingRefusal(fixture, expected, before);
                    CollectionAssert.AreEqual(baselineHashes, ReadHashes(fixture), "Certain no-write refusal changed source.");
                    OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(baselineReferences, fixture.Data("list_references"));
                    Assert.AreEqual(refusalProjectVersion, fixture.Data("project_properties")["Version"], "Certain no-write refusal changed the project revision.");
                    Assert.AreEqual(true, fixture.Data("project_persistence_status")["ProjectSaved"]);
                    RequireBridge(fixture);
                    refusalVerified = true;
                    // The known refusal is terminal; no Save, reopen or General retry follows.
                }
                else
                {
                var originalWrite = WriteGeneral(fixture, property, expected, before);
                AssertUnchangedOtherFields(before, originalWrite, nativeField);
                Assert.AreEqual(Convert.ToString(expected, CultureInfo.InvariantCulture), Text(originalWrite, nativeField),
                    "The original native General result did not retain the exact requested value.");
                var live = ReadGeneral(fixture, "AfterOriginalNativeWrite");
                AssertUnchangedOtherFields(before, live, nativeField);
                Assert.AreEqual(Convert.ToString(expected, CultureInfo.InvariantCulture), Text(live, nativeField));
                CollectionAssert.AreEqual(baselineHashes, ReadHashes(fixture), "General must not change synthetic source.");
                OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(baselineReferences, fixture.Data("list_references"));

                // Pending source guarantees the adapter itself is exercised even if metadata was written immediately.
                foreach (string name in Modules) ReplaceMarker(fixture, name, "pending-" + Guid.NewGuid().ToString("N"));
                var expectedHashes = ReadHashes(fixture);
                for (int i = 0; i < Modules.Length; i++) Assert.AreNotEqual(baselineHashes[i], expectedHashes[i]);
                SelectOwnedCode(fixture);
                var selected = fixture.RecordAdapterObservation("Before" + host + "GeneralAdapter");
                Assert.AreEqual(Modules[0], selected["ActiveCodeComponent"]);
                Assert.AreEqual(false, selected["ProjectSaved"]);
                string path = fixture.DocumentPath, previousSelector = fixture.Project;
                int previousPid = fixture.ProcessId;
                string version = (string)fixture.Data("project_properties")["Version"];
                fixture.RecordAdapterStage("Single" + host + "GeneralAdapterSaveStarting", new { Host = fixture.Kind, Property = property,
                    ExpectedHostPath = path, ExpectedProjectVersion = version, ProcessId = previousPid,
                    ExpectedNativeMetadata = NativeMetadata(live), ExpectedHashes = expectedHashes, RetryAllowed = false });
                var saveResponse = fixture.Response("save_host_document", "ExpectedHostPath", path, "ExpectedProjectVersion", version);
                fixture.RecordAdapterStage("Original" + host + "GeneralAdapterResponse", saveResponse);
                if (!Equals(Field(saveResponse, "Ok"), true)) fixture.NativeExecutionUnsettled = true;
                Assert.AreEqual(true, saveResponse["Ok"], Convert.ToString(saveResponse["Error"]));
                var save = VbeBridgeClient.Object(saveResponse["Data"]);
                Assert.AreEqual(host, Field(save, "Host"), "The original adapter response must identify the selected host.");
                if (host == "Access")
                {
                    Assert.AreEqual("VBE.CommandBars.ID3", Field(save, "SaveApi"));
                    Assert.IsNull(Field(save, "HostSaved"), "Access must not fabricate a document Saved property.");
                }
                if (!Equals(Field(save, "Uncertain"), false)) fixture.NativeExecutionUnsettled = true;
                Assert.AreEqual(true, Field(save, "Verified"), "Fresh readback cannot promote the original adapter response.");
                Assert.AreEqual(false, Field(save, "Uncertain"));
                Assert.AreEqual(true, Field(save, "MutationInvoked"));
                Assert.AreEqual(false, Field(save, "PersistenceReopenVerified"));
                fixture.ObserveAdapterOutcome(saveResponse);
                CollectionAssert.AreEqual(expectedHashes, ReadHashes(fixture), "Save changed synthetic source.");
                OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(baselineReferences, fixture.Data("list_references"));
                var afterSave = ReadGeneral(fixture, "AfterOriginalAdapterSave");
                CollectionAssert.AreEqual(NativeMetadata(live), NativeMetadata(afterSave), "Save changed native General fields.");

                fixture.ReopenFromDisk(() => {
                    using (var stream = File.OpenRead(path))
                    using (var hash = SHA256.Create())
                        fixture.RecordAdapterStage(host + "GeneralClosedFileEvidence", new { Host = host, Path = path, Bytes = stream.Length,
                            Sha256 = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", ""),
                            PreviousProcessId = previousPid, HostExitedBeforeHash = true });
                });
                var status = RequireBridge(fixture);
                var fresh = ReadGeneral(fixture, "IndependentFreshDiskRead");
                var reopened = fixture.RecordAdapterObservation(host + "GeneralFreshDiskReopen");
                Assert.AreEqual(true, reopened["ProjectSaved"], "The independent reopened native project is not saved.");
                OfficeProjectReopenIdentity.Require(host, path, previousSelector, previousPid, NativeMetadata(live),
                    typeof(VbeSession).Module.ModuleVersionId, new OfficeProjectReopenIdentity.Evidence {
                        Host = fixture.Kind, DocumentPath = fixture.DocumentPath, Selector = fixture.Project,
                        ProcessId = fixture.ProcessId, Metadata = NativeMetadata(fresh), Observation = reopened, Status = status });
                CollectionAssert.AreEqual(NativeMetadata(live), NativeMetadata(fresh), "Fresh disk native General differs.");
                Assert.AreEqual(Convert.ToString(expected, CultureInfo.InvariantCulture), Text(fresh, nativeField));
                CollectionAssert.AreEqual(expectedHashes, ReadHashes(fixture), "Fresh disk source differs.");
                OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(baselineReferences, fixture.Data("list_references"));
                fixture.RecordAdapterStage(host + "GeneralIndependentReadbackVerified", new { Host = fixture.Kind, Property = property,
                    OriginalProcessId = previousPid, FreshProcessId = fixture.ProcessId, Expected = expected,
                    OriginalGeneralWriteVerified = true, OriginalAdapterResponseVerified = true,
                    NativeFreshDiskReadbackVerified = true, ComMetadataGetterAcceptanceRequired = false,
                    HelpContentQualified = false, MacroExecuted = false });
                }
            }
            catch (Exception error)
            {
                failure = error;
                if (fixture != null)
                    try { fixture.RecordAdapterFailure(error); }
                    catch (Exception evidence) { failure = new AggregateException(host + " General failure and evidence failure.", failure, evidence); }
            }
            finally
            {
                if (fixture != null)
                {
                    try { fixture.Dispose(); }
                    catch (Exception cleanup) { failure = failure == null ? cleanup : new AggregateException(host + " General trial and cleanup failure.", failure, cleanup); }
                    if (failure == null && refusalVerified)
                    {
                        try
                        {
                            Assert.AreEqual(refusalDocumentPath, fixture.DocumentPath);
                            byte[] closedBytes = File.ReadAllBytes(refusalDocumentPath);
                            CollectionAssert.AreEqual(refusalBaselineBytes, closedBytes, "Certain no-write refusal changed the publication compared with its audited pre-open seed copy.");
                            string closedSha;
                            using (var hash = SHA256.Create()) closedSha = BitConverter.ToString(hash.ComputeHash(closedBytes)).Replace("-", "");
                            Assert.AreEqual(OfficeVbeFixture.PublisherSeedSha, closedSha);
                            fixture.RecordAdapterStage("PublisherEncodingRefusalClosedCopyVerified", new {
                                Path = refusalDocumentPath, Bytes = closedBytes.LongLength, Sha256 = closedSha,
                                OriginalSource = OfficeVbeFixture.PublisherSeedPath, OriginalSourceSha256 = OfficeVbeFixture.PublisherSeedSha,
                                HostExitedBeforeRead = true, NormalOriginalDisposeVerified = true,
                                RefusedBeforeWrite = true, SaveAttempts = 0, FreshReopenAttempts = 0, RetryAllowed = false });
                        }
                        catch (Exception disk) { failure = disk; }
                    }
                    try
                    {
                        foreach (string name in new[] { "adapter-only-progress.json", "report.json", "host-shutdown.json" })
                        {
                            string artifact = Path.Combine(fixture.Root, name);
                            if (File.Exists(artifact)) TestContext?.AddResultFile(artifact);
                        }
                    }
                    catch (Exception attach) { failure = failure == null ? attach : new AggregateException(host + " General report attachment failed.", failure, attach); }
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private static void RequirePrivateHostOptIn(string host)
        {
            Assert.IsTrue(host == "Publisher" || host == "Access", "Only explicitly selected native hosts are supported.");
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_TESTS=1 for disposable " + host + " qualification.");
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            if (string.IsNullOrWhiteSpace(desktop))
                Assert.Inconclusive(host + " General qualification requires the existing isolated private desktop helper.");
            IsolatedTestDesktop.RequireCurrent(desktop);
        }

        private static IDictionary<string, object> RequireBridge(OfficeVbeFixture fixture)
        {
            var status = fixture.Data("status");
            Assert.AreEqual(true, status["Connected"]);
            Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(status["HostProcessId"]));
            Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
            return status;
        }

        private static string GeneralCaption(OfficeVbeFixture fixture)
        {
            var captions = fixture.Items("list_commands", "Limit", 500)
                .Where(c => Convert.ToInt32(c["Id"]) == 2578 && Equals(c["Enabled"], true))
                .Select(c => Convert.ToString(c["Caption"])).Distinct(StringComparer.Ordinal).ToArray();
            Assert.AreEqual(1, captions.Length, "Exactly one enabled General command caption is required.");
            Assert.IsFalse(string.IsNullOrWhiteSpace(captions[0]));
            return captions[0];
        }

        private static IDictionary<string, object> ReadGeneral(OfficeVbeFixture fixture, string phase)
        {
            string caption = GeneralCaption(fixture), version = (string)fixture.Data("project_properties")["Version"];
            fixture.RecordAdapterStage(fixture.Kind + "NativeGeneralReadStarting", new { Host = fixture.Kind, Phase = phase,
                ExpectedProjectVersion = version, ControlCaption = caption, ProcessId = fixture.ProcessId, RetryAllowed = false });
            var response = fixture.Response("read_project_general", "ExpectedMode", 2,
                "ExpectedProjectVersion", version, "ControlCaption", caption);
            return AssertGeneralResponse(fixture, response, false, phase);
        }

        private static IDictionary<string, object> WriteGeneral(OfficeVbeFixture fixture, string property, object value,
            IDictionary<string, object> before)
        {
            string caption = GeneralCaption(fixture), version = (string)fixture.Data("project_properties")["Version"];
            fixture.RecordAdapterStage("Single" + fixture.Kind + "NativeGeneralWriteStarting", new { Host = fixture.Kind, Property = property, Value = value,
                ExpectedProjectVersion = version, ExpectedOptionsVersion = Text(before, "OptionsVersion"),
                ControlCaption = caption, ProcessId = fixture.ProcessId, RetryAllowed = false });
            var response = fixture.Response("set_project_general", "Property", property, "Value", value,
                "ExpectedMode", 2, "ExpectedProjectVersion", version, "ControlCaption", caption,
                "ExpectedOptionsVersion", Text(before, "OptionsVersion"));
            return AssertGeneralResponse(fixture, response, true, "OriginalNativeGeneralWrite");
        }

        private static void AssertEncodingRefusal(OfficeVbeFixture fixture, object value, IDictionary<string, object> before)
        {
            string caption = GeneralCaption(fixture), version = (string)fixture.Data("project_properties")["Version"];
            fixture.RecordAdapterStage("Single" + fixture.Kind + "NativeEncodingRefusalStarting", new { Host = fixture.Kind, Value = value,
                ExpectedProjectVersion = version, ExpectedOptionsVersion = Text(before, "OptionsVersion"),
                ControlCaption = caption, ProcessId = fixture.ProcessId, FieldInvocationLimit = 0,
                AdapterSaveInvocationLimit = 0, FreshReopenAllowed = false, RetryAllowed = false });
            var response = fixture.Response("set_project_general", "Property", "HelpFile", "Value", value,
                "ExpectedMode", 2, "ExpectedProjectVersion", version, "ControlCaption", caption,
                "ExpectedOptionsVersion", Text(before, "OptionsVersion"));
            fixture.RecordAdapterStage("Original" + fixture.Kind + "NativeGeneralResponse", new {
                Phase = "OriginalUnsupportedEncodingRefusal", Write = true, Response = response });
            // Mark before assertions: an unknown or mismatched original refusal remains quarantined without cleanup retry.
            fixture.NativeExecutionUnsettled = true;
            Assert.AreEqual(true, Field(response, "Ok"), "An explicit General result is required for the no-write claim.");
            var result = VbeBridgeClient.Object(response["Data"]);
            Assert.AreEqual(true, Field(result, "Terminal"));
            Assert.AreEqual(false, Field(result, "Available"));
            Assert.AreEqual(false, Field(result, "Uncertain"));
            Assert.AreEqual(true, Field(result, "CommandEntered"));
            Assert.AreEqual(true, Field(result, "DialogClosed"));
            Assert.AreEqual(true, Field(result, "OriginalExecuteReturned"));
            Assert.AreEqual(true, Field(result, "RefusedBeforeWrite"));
            Assert.AreEqual(false, Field(result, "MutationInvoked"));
            Assert.AreEqual(false, Field(result, "ControlValueVerified"));
            Assert.AreEqual(false, Field(result, "CommittedRequested"));
            Assert.AreEqual(false, Field(result, "PersistenceVerified"));
            Assert.AreEqual(false, Field(result, "RetryAllowed"));
            Assert.AreEqual(1, Field(result, "OpenAttempts"));
            Assert.AreEqual(1, Field(result, "CancelAttempts"));
            foreach (string attempts in new[] { "FieldAttempts", "OkAttempts" })
                Assert.AreEqual(0, Field(result, attempts), "Unsupported encoding must be refused before field dispatch: " + attempts);
            foreach (string field in NativeFields.Concat(new[] { "OptionsVersion" }))
                Assert.IsNull(Field(result, field), "Refused metadata must be redacted: " + field);
            string error = Convert.ToString(Field(result, "Error"));
            StringAssert.Contains(error, "cannot preserve the exact requested value");
            StringAssert.Contains(error, "no field write was entered");
            fixture.RecordAdapterStage(fixture.Kind + "NativeEncodingRefusalVerified", new {
                ExpectedNativeEncodingRefusal = true, FieldMutationInvoked = false, FieldAttempts = 0,
                OkAttempts = 0, CancelAttempts = 1, SaveAttempts = 0, FreshReopenAttempts = 0,
                MetadataRedacted = true, OriginalModalRetained = false, NormalHostCleanupRequired = true,
                UnsupportedUnicodePersistenceQualified = false, RetryAllowed = false });
            fixture.NativeExecutionUnsettled = false; // Only the fully verified original Cancel makes ordinary read-only checks and cleanup safe.
        }

        private static IDictionary<string, object> AssertGeneralResponse(OfficeVbeFixture fixture,
            IDictionary<string, object> response, bool write, string phase)
        {
            fixture.RecordAdapterStage("Original" + fixture.Kind + "NativeGeneralResponse", new { Host = fixture.Kind, Phase = phase, Write = write, Response = response });
            if (!Equals(Field(response, "Ok"), true))
            {
                // This helper never attempts a fallback setter, Cancel, Save or native cleanup after unknown General execution.
                fixture.NativeExecutionUnsettled = true;
                Assert.Fail("Original General response failed; no retry: " + Convert.ToString(Field(response, "Error")));
            }
            var result = VbeBridgeClient.Object(response["Data"]);
            if (!Equals(Field(result, "Terminal"), true) || !Equals(Field(result, "Uncertain"), false) ||
                (Equals(Field(result, "CommandEntered"), true) &&
                 (!Equals(Field(result, "DialogClosed"), true) || !Equals(Field(result, "OriginalExecuteReturned"), true))))
                fixture.NativeExecutionUnsettled = true;
            Assert.IsNull(Field(result, "Error"), Convert.ToString(Field(result, "Error")));
            Assert.AreEqual(true, Field(result, "Terminal"));
            Assert.AreEqual(true, Field(result, "Available"));
            Assert.AreEqual(false, Field(result, "Uncertain"));
            Assert.AreEqual(true, Field(result, "CommandEntered"));
            Assert.AreEqual(true, Field(result, "OriginalExecuteReturned"));
            Assert.AreEqual(true, Field(result, "DialogClosed"));
            Assert.AreEqual(false, Field(result, "PersistenceVerified"));
            Assert.AreEqual(false, Field(result, "RetryAllowed"));
            Assert.AreEqual(write, Field(result, "MutationInvoked"));
            Assert.AreEqual(write, Field(result, "ControlValueVerified"));
            Assert.AreEqual(write, Field(result, "CommittedRequested"));
            Assert.AreEqual(1, Field(result, "OpenAttempts"));
            Assert.AreEqual(write ? 1 : 0, Field(result, "FieldAttempts"));
            Assert.AreEqual(write ? 1 : 0, Field(result, "OkAttempts"));
            Assert.AreEqual(write ? 0 : 1, Field(result, "CancelAttempts"));
            Assert.IsFalse(string.IsNullOrWhiteSpace(Text(result, "OptionsVersion")));
            foreach (string field in NativeFields) Assert.IsInstanceOfType(Field(result, field), typeof(string), "Missing native General field: " + field);
            return result;
        }

        private static object Field(IDictionary<string, object> data, string key)
        { object value; return data.TryGetValue(key, out value) ? value : null; }
        private static string Text(IDictionary<string, object> data, string key) => Field(data, key) as string;
        private static string[] NativeMetadata(IDictionary<string, object> data) => new[] {
            "Name=" + Text(data, "Name"), "Description=" + Text(data, "Description"), "HelpFile=" + Text(data, "HelpFile"),
            "HelpContextID=" + Text(data, "HelpContextText"), "ConditionalCompilation=" + Text(data, "ConditionalCompilation") };
        private static void AssertUnchangedOtherFields(IDictionary<string, object> before, IDictionary<string, object> after, string changed)
        { foreach (string field in NativeFields.Where(field => field != changed)) Assert.AreEqual(Text(before, field), Text(after, field), "General changed unrelated field: " + field); }
        private static string[] ReadHashes(OfficeVbeFixture fixture) => Modules.Select(name => (string)fixture.Data("read_module", "Module", name)["Sha256"]).ToArray();
        private static void SelectOwnedCode(OfficeVbeFixture fixture) => fixture.Data("select_code", "Module", Modules[0],
            "StartLine", 1, "ExpectedSha256", fixture.Data("read_module", "Module", Modules[0])["Sha256"]);
        private static void ReplaceMarker(OfficeVbeFixture fixture, string name, string marker)
        {
            var original = fixture.Data("read_module", "Module", name);
            fixture.Data("replace_lines", "Module", name, "StartLine", 1, "Count", fixture.Data("component_properties", "Module", name)["CodeLines"],
                "ExpectedSha256", original["Sha256"], "Text", "Option Explicit\r\n' Synthetic native General qualification: " + marker + "\r\n");
        }
    }

    /// <summary>Runs the same owner-thread General and adapter persistence qualification against disposable Access projects.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), TestCategory("AccessGeneral"), DoNotParallelize]
    public sealed class AccessGeneralQualificationTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void AccessNativeAnsiHelpFileSaveReopen()
        { new PublisherGeneralQualificationTests { TestContext = TestContext }.Qualify("HelpFile", "Owned help été_ß.chm", host: "Access"); }

        [STATestMethod]
        public void AccessNativeHelpContextSaveReopen()
        { new PublisherGeneralQualificationTests { TestContext = TestContext }.Qualify("HelpContextID", host: "Access"); }
    }
}
