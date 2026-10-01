using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual semantic reopen gate without starting Office or emitting native commands.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeProjectReopenIdentityTests
    {
        private static readonly Guid Candidate = new Guid("d8f31d57-0000-4000-8000-000000000001");
        private static readonly string FilePath = Path.Combine(Path.GetTempPath(), "VBAi-reopen-identity", "Disposable.pub");
        private static readonly string[] Metadata = { "Name=Project", "Description=synthetic", "HelpFile=", "HelpContextID=0" };

        [DataTestMethod]
        [DataRow("name", "path")]
        [DataRow("path", "name")]
        public void CanonicalizedSelectorRetainsExactNativeDocumentAndMetadataIdentity(string before, string after)
        {
            var actual = Create();
            actual.Selector = Selector(after);
            Persistence(actual)["Project"] = actual.Selector;
            Require(actual, Selector(before));
        }

        [DataTestMethod]
        [DataRow("name")]
        [DataRow("path")]
        public void UnchangedSelectorAlsoRequiresNativeProof(string selector)
        {
            var actual = Create(); actual.Selector = Selector(selector);
            Persistence(actual)["Project"] = actual.Selector;
            Require(actual, actual.Selector);
        }

        [TestMethod]
        public void AccessReopenRequiresFreshProcessButWordCanRetainItsOwnedProcess()
        {
            foreach (string host in new[] { "Access", "Word", "PowerPoint" })
            {
                var actual = Create(); actual.Host = host; Persistence(actual)["Host"] = host;
                OfficeProjectReopenIdentity.Require(host, FilePath, "Project", host == "Access" ? 100 : 200,
                    Metadata, Candidate, actual);
            }
        }

        [DataTestMethod]
        [DataRow("host")]
        [DataRow("document-path")]
        [DataRow("relative-document-path")]
        [DataRow("blank-document-path")]
        [DataRow("selector-name")]
        [DataRow("selector-path")]
        [DataRow("old-selector-name")]
        [DataRow("old-selector-path")]
        [DataRow("metadata-name")]
        [DataRow("metadata-name-missing")]
        [DataRow("metadata-name-duplicate")]
        [DataRow("metadata-name-blank")]
        [DataRow("reused-pid")]
        [DataRow("zero-pid")]
        [DataRow("old-zero-pid")]
        [DataRow("candidate")]
        [DataRow("status-pid")]
        [DataRow("status-pid-missing")]
        [DataRow("status-missing")]
        [DataRow("observation-missing")]
        [DataRow("observation-error")]
        [DataRow("observation-pid")]
        [DataRow("window-pid")]
        [DataRow("observation-path")]
        [DataRow("persistence-missing")]
        [DataRow("persistence-host")]
        [DataRow("unavailable")]
        [DataRow("unverified")]
        [DataRow("identity-missing")]
        [DataRow("persistence-pid")]
        [DataRow("persistence-path")]
        [DataRow("persistence-selector")]
        public void CanonicalizationDoesNotAdmitForeignIncompleteOrUnverifiedIdentity(string defect)
        {
            var actual = Create(); string before = "Project"; int previousPid = 100;
            switch (defect)
            {
                case "host": actual.Host = "Access"; break;
                case "document-path": actual.DocumentPath = ForeignPath(); break;
                case "relative-document-path": actual.DocumentPath = "Disposable.pub"; break;
                case "blank-document-path": actual.DocumentPath = ""; break;
                case "selector-name": actual.Selector = "RecoveredProject"; break;
                case "selector-path": actual.Selector = ForeignPath(); break;
                case "old-selector-name": before = "RecoveredProject"; break;
                case "old-selector-path": before = ForeignPath(); break;
                case "metadata-name": actual.Metadata = new[] { "Name=ForeignProject" }; break;
                case "metadata-name-missing": actual.Metadata = new[] { "Description=synthetic" }; break;
                case "metadata-name-duplicate": actual.Metadata = new[] { "Name=Project", "Name=Project" }; break;
                case "metadata-name-blank": actual.Metadata = new[] { "Name=" }; break;
                case "reused-pid": previousPid = actual.ProcessId; break;
                case "zero-pid": actual.ProcessId = 0; break;
                case "old-zero-pid": previousPid = 0; break;
                case "candidate": actual.Status["AssemblyModuleVersionId"] = Guid.Empty.ToString("D"); break;
                case "status-pid": actual.Status["HostProcessId"] = 999; break;
                case "status-pid-missing": actual.Status.Remove("HostProcessId"); break;
                case "status-missing": actual.Status = null; break;
                case "observation-missing": actual.Observation = null; break;
                case "observation-error": actual.Observation["ObservationError"] = "Getter failed"; break;
                case "observation-pid": actual.Observation["ProcessId"] = 999; break;
                case "window-pid": actual.Observation["ApplicationOwnerPid"] = 999; break;
                case "observation-path": actual.Observation["DocumentPath"] = ForeignPath(); break;
                case "persistence-missing": actual.Observation.Remove("Persistence"); break;
                case "persistence-host": Persistence(actual)["Host"] = "Access"; break;
                case "unavailable": Persistence(actual)["HostAvailable"] = false; break;
                case "unverified": Persistence(actual)["IdentityVerified"] = false; break;
                case "identity-missing": Persistence(actual).Remove("IdentityVerified"); break;
                case "persistence-pid": Persistence(actual)["OwnerProcessId"] = 999; break;
                case "persistence-path": Persistence(actual)["HostPath"] = ForeignPath(); break;
                case "persistence-selector": Persistence(actual)["Project"] = "RecoveredProject"; break;
                default: Assert.Fail("Unknown prepared countercase."); break;
            }
            Assert.ThrowsException<AssertFailedException>(() => OfficeProjectReopenIdentity.Require("Publisher", FilePath,
                before, previousPid, Metadata, Candidate, actual));
        }

        [TestMethod]
        public void ReopenIdentityDoesNotReplaceFullMetadataSourceOrReferencePersistenceChecks()
        {
            var actual = Create(); actual.Metadata = new[] { "Name=Project", "Description=changed" };
            Require(actual, "Project");
            // The native workflow still performs its full ordered metadata/hash/reference comparisons after this gate.
            Assert.ThrowsException<AssertFailedException>(() => CollectionAssert.AreEqual(Metadata, actual.Metadata));
        }

        private static void Require(OfficeProjectReopenIdentity.Evidence value, string previous) =>
            OfficeProjectReopenIdentity.Require("Publisher", FilePath, previous, 100, Metadata, Candidate, value);

        private static OfficeProjectReopenIdentity.Evidence Create() => new OfficeProjectReopenIdentity.Evidence {
            Host = "Publisher", DocumentPath = FilePath, Selector = FilePath, ProcessId = 200,
            Metadata = (string[])Metadata.Clone(),
            Status = new Dictionary<string, object> { ["AssemblyModuleVersionId"] = Candidate.ToString("D"), ["HostProcessId"] = 200 },
            Observation = new Dictionary<string, object> { ["ProcessId"] = 200, ["DocumentPath"] = FilePath,
                ["ApplicationOwnerPid"] = (uint)200, ["Persistence"] = new Dictionary<string, object> {
                    ["Project"] = FilePath, ["Host"] = "Publisher", ["HostAvailable"] = true,
                    ["IdentityVerified"] = true, ["OwnerProcessId"] = 200, ["HostPath"] = FilePath } }
        };

        private static IDictionary<string, object> Persistence(OfficeProjectReopenIdentity.Evidence value) =>
            (IDictionary<string, object>)value.Observation["Persistence"];
        private static string Selector(string kind) => kind == "name" ? "Project" : FilePath;
        private static string ForeignPath() => Path.Combine(Path.GetTempPath(), "VBAi-foreign-reopen", "Disposable.pub");
    }
}
