using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        internal const string PublisherSeedSha = "25FD2FC6C2261D61255CB416A5811A06E70C4780493903633C50BCF780FB9117";
        internal const string PublisherSeedPath = @"C:\Users\init-\Documents\Codex\q012-20261003-next\publisher-general-ansiguard\native\17-PublisherNativeHelpContextSaveReopen\host-evidence\Publisher\973d982211ee44f5b65aaa757e6945bb\Disposable.pub";
        internal bool PublisherSerializedSeed { get; private set; }

        /// <summary>Qualifies an existing serialized VBA publication; it does not qualify NewDocument or first SaveAs.</summary>
        internal static OfficeVbeFixture StartPublisherSerializedQualificationSeed()
        {
            string path = Environment.GetEnvironmentVariable("VBAi_TEST_PUBLISHER_SERIALIZED_SEED");
            if (path == null) Assert.Inconclusive("Set VBAi_TEST_PUBLISHER_SERIALIZED_SEED to the frozen owned Publisher qualification seed.");
            RequirePublisherSeedFile(path);
            // The selected desktop is checked before any host creation. Main requires its
            // explicit opt-in and empty private descriptors; private keeps RequireCurrent.
            NativeTestDesktop.Current();
            return Start("Publisher", false, serializedPublisherSeed: path);
        }

        internal static void RequirePublisherSeedBytes(string path, long bytes, string sha)
        {
            if (!string.Equals(path, PublisherSeedPath, StringComparison.OrdinalIgnoreCase) || bytes != 90624 || sha != PublisherSeedSha)
                throw new InvalidOperationException("Only the exact frozen owned serialized Publisher seed is permitted.");
        }

        private static string SeedFileHash(string path)
        {
            using (var file = File.OpenRead(path))
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
        }

        private static void RequirePublisherSeedFile(string path)
        {
            RequirePublisherSeedBytes(Path.GetFullPath(path), new FileInfo(path).Length, SeedFileHash(path));
            RequirePublisherSeedReceipts(
                File.ReadAllText(Path.Combine(Path.GetDirectoryName(path), "shutdown-before-reopen-183824.json")),
                File.ReadAllText(Path.Combine(Path.GetDirectoryName(path), "adapter-only-progress.json")));
        }

        /// <summary>Validates the actual typed-deserializer receipt shape without constructing an Office instance.</summary>
        internal static void RequirePublisherSeedReceipts(string shutdownJson, string progressJson)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var receipt = serializer.Deserialize<Dictionary<string, object>>(shutdownJson);
            var life = VbeBridgeClient.Object(receipt["Lifecycle"]);
            RequirePublisherSeedProducer(Convert.ToInt32(life["ProcessId"]), Convert.ToString(life["ProcessStartedUtc"]),
                Equals(life["ProcessExitObserved"], true) && Equals(life["ExitCodeObserved"], true),
                Convert.ToInt32(life["ExitCode"]), Equals(life["ForcedTermination"], true),
                Convert.ToString(life["State"]), Convert.ToString(receipt["DocumentPath"]));
            var progress = serializer.Deserialize<Dictionary<string, object>>(progressJson);
            var steps = progress["Steps"] as IList;
            if (steps == null) throw new InvalidOperationException("The seed progress Steps must be a JSON array.");
            // Deserialize<Dictionary<string, object>> returns ArrayList for nested arrays;
            // DeserializeObject (used by the bridge) returns object[]. Both implement IList.
            var matches = steps.Cast<object>().Select(VbeBridgeClient.Object)
                .Where(s => Equals(Field(s, "AdapterStage"), "PublisherGeneralClosedFileEvidence")).ToArray();
            Assert.AreEqual(1, matches.Length, "The seed must have one original closed-file receipt.");
            var disk = VbeBridgeClient.Object(matches[0]["Inputs"]);
            RequirePublisherSeedBytes(Convert.ToString(disk["Path"]), Convert.ToInt64(disk["Bytes"]), Convert.ToString(disk["Sha256"]));
            Assert.AreEqual(183824, Convert.ToInt32(disk["PreviousProcessId"]));
            Assert.AreEqual(true, disk["HostExitedBeforeHash"]);
        }

        internal static void RequirePublisherSeedProducer(int pid, string birth, bool exited, int exitCode, bool forced, string state, string path)
        {
            if (pid != 183824 || birth != "2026-10-03T19:42:01.6309813Z" || !exited || exitCode != 0 || forced ||
                state != "EXIT_OBSERVED_HANDLE_RELEASED" || path != PublisherSeedPath)
                throw new InvalidOperationException("The original Publisher seed producer did not close normally with its exact identity.");
        }

        private void CopyPublisherQualificationSeed(string source)
        {
            RequireApplicationOwner(); RecheckPrivatePublisherBootstrap(); RequirePublisherSeedFile(source);
            Assert.IsFalse(File.Exists(DocumentPath), "A seed copy cannot overwrite a publication.");
            using (var input = File.Open(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var sha = SHA256.Create())
                    RequirePublisherSeedBytes(source, input.Length, BitConverter.ToString(sha.ComputeHash(input)).Replace("-", ""));
                input.Position = 0;
                RecheckPrivatePublisherBootstrap();
                using (var output = File.Open(DocumentPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) input.CopyTo(output);
            }
            Assert.AreEqual(PublisherSeedSha, SeedFileHash(DocumentPath));
            RequirePublisherSeedFile(source); RecheckPrivatePublisherBootstrap();
            PublisherSerializedSeed = true;
            RecordAdapterStage("PublisherSerializedQualificationSeed", new
            {
                Source = source,
                SourceSha256 = PublisherSeedSha,
                SourceBytes = 90624,
                OriginalProducerPid = 183824,
                Scope = "ExistingSerializedPublication",
                NewPublicationQualified = false,
                FirstSaveAsQualified = false,
                CopyEntries = 1,
                OpenEntriesLimit = 1,
                BootstrapSaveAllowed = false,
                MacroExecutionAllowed = false
            });
        }

        internal static void RequirePublisherSeedComponent(string name, int type, string code, string sha)
        {
            string expectedCode, expectedSha; int expectedType;
            if (name == "ThisDocument") { expectedType = 100; expectedCode = ""; expectedSha = null; }
            else if (name == "PublisherGeneralModule")
            {
                expectedType = 1; expectedCode = "Option Explicit\r\n' Synthetic native General qualification: pending-c2468bacba634672be8b7ce6d656fd37\r\n";
                expectedSha = "48f46f56d9e9a767aabd8d9f79aadc4873b63cf4ff8672d735f5164eb91477e4";
            }
            else if (name == "PublisherGeneralClass")
            {
                expectedType = 2; expectedCode = "Option Explicit\r\n' Synthetic native General qualification: pending-fd1eb2856671459ca28c14179f6b599a\r\n";
                expectedSha = "9bee783febb00e45ef1246b1b2d5babc215d50d79c52ae72bb32040aa45adcd4";
            }
            else throw new InvalidOperationException("An extra or foreign seed component is refused.");
            if (type != expectedType || code != expectedCode || (expectedSha != null && sha != expectedSha))
                throw new InvalidOperationException("The seed component's exact type, inert source or revision changed.");
        }

        internal void AuditPublisherSerializedSeed()
        {
            Assert.IsTrue(PublisherSerializedSeed); RequirePublisherPublication("BeforeSerializedSeedAudit", true);
            string revision = Convert.ToString(Data("project_properties")["Version"]);
            var modules = Items("list_modules");
            Assert.AreEqual(3, modules.Length);
            CollectionAssert.AreEquivalent(new[] { "ThisDocument", "PublisherGeneralModule", "PublisherGeneralClass" }, modules.Select(m => Convert.ToString(m["Name"])).ToArray());
            foreach (var item in modules)
            {
                string name = Convert.ToString(item["Name"]); var source = Data("read_module", "Module", name);
                RequirePublisherSeedComponent(name, Convert.ToInt32(item["Type"]), Convert.ToString(source["Code"]), Convert.ToString(source["Sha256"]));
            }
            var references = Data("list_references");
            Assert.AreEqual("6741f0c3bfcd11bc639966bbdd3ee3ed6342f1711396ed13f6cbc043945cb2bf", references["Version"]);
            Assert.AreEqual(4, ((object[])references["References"]).Length);
            var expectedReferences = new[] {
                "VBA|{000204EF-0000-0000-C000-000000000046}|4|2|True|C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL",
                "Publisher|{0002123C-0000-0000-C000-000000000046}|2|3|True|C:\\Program Files\\Microsoft Office\\root\\Office16\\MSPUB.TLB",
                "stdole|{00020430-0000-0000-C000-000000000046}|2|0|False|C:\\Windows\\System32\\stdole2.tlb",
                "Office|{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}|2|8|False|C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL" };
            var observedReferences = ((object[])references["References"]).Select(entry =>
            {
                var reference = VbeBridgeClient.Object(entry); Assert.AreEqual(false, reference["IsBroken"]);
                return string.Join("|", new[] { "Name", "Guid", "Major", "Minor", "BuiltIn", "FullPath" }.Select(key => Convert.ToString(reference[key])));
            }).ToArray();
            CollectionAssert.AreEquivalent(expectedReferences, observedReferences);
            Assert.AreEqual(revision, Data("project_properties")["Version"], "The seed project changed during its read-only audit.");
            RequirePublisherPublication("AfterSerializedSeedAudit", true);
            var captions = Items("list_commands", "Limit", 500).Where(c => Convert.ToInt32(c["Id"]) == 2578 && Equals(c["Enabled"], true))
                .Select(c => Convert.ToString(c["Caption"])).Distinct(StringComparer.Ordinal).ToArray();
            Assert.AreEqual(1, captions.Length);
            var response = Response("read_project_general", "ExpectedMode", 2, "ExpectedProjectVersion", revision, "ControlCaption", captions[0]);
            if (!Equals(Field(response, "Ok"), true)) { NativeExecutionUnsettled = true; Assert.Fail("Seed native General baseline failed; no retry."); }
            var general = VbeBridgeClient.Object(response["Data"]);
            if (!Equals(Field(general, "Terminal"), true) || !Equals(Field(general, "Uncertain"), false) ||
                !Equals(Field(general, "DialogClosed"), true) || !Equals(Field(general, "OriginalExecuteReturned"), true))
                NativeExecutionUnsettled = true;
            Assert.IsFalse(NativeExecutionUnsettled);
            Assert.AreEqual(true, general["Available"]); Assert.AreEqual(false, general["MutationInvoked"]);
            Assert.AreEqual(0, general["FieldAttempts"]); Assert.AreEqual(0, general["OkAttempts"]); Assert.AreEqual(1, general["CancelAttempts"]);
            RequirePublisherSeedGeneral(general);
            Assert.AreEqual(revision, Data("project_properties")["Version"]);
            RequirePublisherPublication("AfterSerializedSeedGeneralBaseline", true);
            RecordAdapterStage("PublisherSerializedSeedAudit", new
            {
                Components = modules,
                References = references,
                ProjectVersion = revision,
                General = general,
                ExecutableProcedures = 0,
                SourceMutationInvoked = false,
                AuditVerified = true
            });
        }

        internal static void RequirePublisherSeedGeneral(IDictionary<string, object> general)
        {
            foreach (var field in new[] { "Description", "HelpFile", "ConditionalCompilation" }) Assert.AreEqual("", general[field]);
            Assert.AreEqual("Project", general["Name"]); Assert.AreEqual("321", general["HelpContextText"]);
        }
    }
}
