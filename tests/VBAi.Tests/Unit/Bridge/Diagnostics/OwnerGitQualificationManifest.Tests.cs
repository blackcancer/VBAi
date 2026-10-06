using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OwnerGitQualificationManifestTests
    {
        private const string Id = "0123456789abcdef0123456789abcdef";
        private const string Sha = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";
        private static readonly string Fixture = Path.Combine(@"C:\Owned", Id);
        private static readonly string Evidence = Path.Combine(@"C:\Evidence", "LabelButton-" + Id);
        private static readonly string Workbook = Path.Combine(Fixture, "owned.xlsm");
        private static readonly string Before = Path.Combine(Evidence, "before");
        private static readonly string Changed = Path.Combine(Evidence, "changed");

        private static OwnerGitQualificationStep Step(string verb, int number, string error = null) => new OwnerGitQualificationStep {
            Id = number.ToString("x32"), Verb = verb, ExpectedSnapshotDirectory = Before,
            ExpectedSnapshotSha256 = Sha, TargetSnapshotDirectory = Changed, TargetSnapshotSha256 = Sha,
            CheckpointId = verb == "checkpoint_restore" ? "20261006000000000-abcdef12" : null, ExpectedErrorSubstring = error };

        private static OwnerGitQualificationManifest Plan(params OwnerGitQualificationStep[] steps) => new OwnerGitQualificationManifest {
            Version = 1, OwnerPid = 42, OwnerBirthUtcTicks = 638000000000000000,
            OwnerNativeTid = 100, VbeHandle = 200,
            AssemblyMvid = "00000000-0000-0000-0000-000000000001", AssemblySha256 = Sha,
            FixtureRoot = Fixture, WorkbookPath = Workbook, EvidenceRoot = Evidence,
            RepoRelativePath = "local.git", Project = Workbook, Branch = "qualification-layout",
            RemoteUrl = null, RemoteCommit = null, Steps = steps };

        private static string Json(OwnerGitQualificationManifest plan) => new JavaScriptSerializer().Serialize(plan);

        [TestMethod]
        public void FixedLocalAndCorruptionSequencesAcceptExactPinnedPaths()
        {
            var local = OwnerGitQualificationManifest.Parse(Json(Plan(Step("checkpoint_restore", 1),
                Step("controlled_interruption", 2), Step("rollback", 3))));
            Assert.AreEqual(3, local.Steps.Length);
            var corrupt = OwnerGitQualificationManifest.Parse(Json(Plan(Step("checkpoint_restore", 1, "Form resources are missing"))));
            Assert.AreEqual("checkpoint_restore", corrupt.Steps[0].Verb);
            Assert.AreEqual(local.Steps[0], OwnerGitQualificationManifest.RequireStep(local, 0, false, local.Steps[0].Id));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireStep(local, 0, false, local.Steps[1].Id));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireStep(local, 0, true, local.Steps[0].Id));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireStep(local, 3, false, local.Steps[2].Id));
        }

        [TestMethod]
        public void PullAcceptsOnlyTheNamedSyntheticRemoteAndFrozenCommit()
        {
            var plan = Plan(Step("pull", 1));
            plan.RemoteUrl = "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git";
            plan.RemoteCommit = new string('a', 40);
            plan.Branch = "qualification-userform-20261006";
            Assert.AreEqual("pull", OwnerGitQualificationManifest.Parse(Json(plan)).Steps.Single().Verb);
            plan.RemoteUrl = "https://github.com/private/production.git";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.RemoteUrl = "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git";
            plan.Branch = "main";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Branch = "qualification-userform-20261006";
            plan.RemoteCommit = "unfrozen";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
        }

        [TestMethod]
        public void UnexpectedFieldsUnboundProjectsAndArbitrarySequencesAreRefused()
        {
            var plan = Plan(Step("checkpoint_restore", 1), Step("controlled_interruption", 2), Step("rollback", 3));
            string valid = Json(plan);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(valid.TrimEnd('}') + ",\"Code\":\"run\"}"));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(valid.Replace("\"OwnerPid\":42", "\"OwnerPid\":\"42\"")));
            plan.Project = "unrelated";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Project = Workbook;
            plan.Steps = new[] { Step("rollback", 1), Step("controlled_interruption", 2), Step("checkpoint_restore", 3) };
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Steps = new[] { Step("checkpoint_restore", 1), Step("controlled_interruption", 1), Step("rollback", 3) };
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
        }

        [TestMethod]
        public void PathsCannotLeaveFrozenRootsOrContainTraversalOrAlternateStreams()
        {
            var plan = Plan(Step("checkpoint_restore", 1), Step("controlled_interruption", 2), Step("rollback", 3));
            plan.RepoRelativePath = @"..\production.git";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.RepoRelativePath = "local.git";
            plan.Steps[1].TargetSnapshotDirectory = @"C:\Other\snapshot";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Steps[1].TargetSnapshotDirectory = Changed;
            plan.WorkbookPath = Workbook + ":stream";
            plan.Project = plan.WorkbookPath;
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.WorkbookPath = Workbook;
            plan.Project = Workbook;
            plan.Steps[0].ExpectedSnapshotDirectory = Before + ":stream";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Steps[0].ExpectedSnapshotDirectory = Before;
            plan.Steps[0].TargetSnapshotDirectory = Changed + ":stream";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.Steps[0].TargetSnapshotDirectory = Changed;
            plan.FixtureRoot = Fixture + ":stream";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.FixtureRoot = Fixture;
            plan.EvidenceRoot = Evidence + ":stream";
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
            plan.EvidenceRoot = Evidence;
            Assert.AreEqual(3, OwnerGitQualificationManifest.Parse(Json(plan)).Steps.Length,
                "Rejecting malformed paths must preserve the valid fixed plan contract.");
        }

        [TestMethod]
        public void ClassicPathPreflightRejectsTheObservedTerminalBoundaryBeforeAnyReceipt()
        {
            // The failed native packet had EvidenceRoot=203, intent=258 and
            // terminal/mutation=260. A shorter root must be chosen before Excel.
            string longRoot = Path.Combine(@"C:\Evidence",
                new string('x', 203 - @"C:\Evidence".Length - 1 - ("LabelButton-" + Id).Length - 1),
                "LabelButton-" + Id);
            Assert.AreEqual(203, longRoot.Length);
            string prefix = Path.Combine(longRoot, "owner-git-" + Id);
            Assert.AreEqual(258, (prefix + ".intent.json").Length);
            Assert.AreEqual(260, (prefix + ".mutation.json").Length);
            Assert.AreEqual(260, (prefix + ".terminal.json").Length);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireEvidenceRootBudget(longRoot));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireReceiptPaths(longRoot, Id));
            var plan = Plan(Step("checkpoint_restore", 1, "invalid/truncated OLE"));
            plan.EvidenceRoot = longRoot;
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
        }

        [TestMethod]
        public void EveryReceiptSuffixAndRepositoryRefHasAClassicPathBudget()
        {
            string shortRoot = Path.Combine(@"C:\Evidence", "LabelButton-" + Id);
            OwnerGitQualificationManifest.RequireReceiptPaths(shortRoot, Id);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireReceiptPaths(shortRoot, "bad-id"));
            string root179 = Path.Combine(@"C:\Evidence",
                new string('x', OwnerGitQualificationManifest.MaxEvidenceRootLength - @"C:\Evidence".Length - 1 - ("LabelButton-" + Id).Length - 1),
                "LabelButton-" + Id);
            Assert.AreEqual(179, root179.Length);
            OwnerGitQualificationManifest.RequireEvidenceRootBudget(root179);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireEvidenceRootBudget(root179 + "x"));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireEvidenceRootBudget(null));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireEvidenceRootBudget(shortRoot + ":stream"));
            string nearLimit = Path.Combine(@"C:\", new string('a', 100), new string('b', 100), new string('c', 54));
            Assert.AreEqual(259, nearLimit.Length);
            OwnerGitQualificationManifest.RequireClassicFilePath(nearLimit);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireClassicFilePath(nearLimit + "x"));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireClassicFilePath(null));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireClassicFilePath(shortRoot + ":stream"));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireClassicDirectoryPath(nearLimit));
            var plan = Plan(Step("checkpoint_restore", 1, "invalid/truncated OLE"));
            plan.EvidenceRoot = root179;
            Assert.AreEqual(OwnerGitQualificationManifest.MaxEvidenceRootLength, plan.EvidenceRoot.Length);
            plan.Branch = new string('b', 128);
            // The branch fits its schema but its on-disk ref does not fit the
            // same direct net48 API used by the qualification repository.
            string repoRef = Path.Combine(plan.EvidenceRoot, plan.RepoRelativePath, "refs", "remotes", "origin", plan.Branch);
            Assert.IsTrue(repoRef.Length > OwnerGitQualificationManifest.MaxClassicFilePathLength);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.Parse(Json(plan)));
        }

        [TestMethod]
        public void BridgeRequestHasExactlyThreeDataFieldsAndNoCodeOrPaths()
        {
            string valid = "{\"Command\":\"diagnostic_userform_git\",\"Action\":\"" + Id + "\",\"ExpectedSha256\":\"" + Sha + "\"}";
            OwnerGitQualificationManifest.RequireExactRequest(valid);
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireExactRequest(valid.TrimEnd('}') + ",\"Path\":\"C:\\\\Private\"}"));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireExactRequest(valid.Replace("diagnostic_userform_git", "immediate_execute")));
            Assert.ThrowsException<ArgumentException>(() => OwnerGitQualificationManifest.RequireExactRequest(valid.Replace(Sha, "loose")));
        }

        [TestMethod]
        public void OwnerIdentityAndCurrentApprovalRefuseForeignStaOrPolicy()
        {
            var plan = Plan(Step("checkpoint_restore", 1), Step("controlled_interruption", 2), Step("rollback", 3));
            OwnerGitQualificationManifest.RequireOwnerIdentity(plan, 42, plan.OwnerBirthUtcTicks, 42, 100, 100, 200, true,
                System.Threading.ApartmentState.STA);
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                43, plan.OwnerBirthUtcTicks, 42, 100, 100, 200, true, System.Threading.ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                42, plan.OwnerBirthUtcTicks + 1, 42, 100, 100, 200, true, System.Threading.ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                42, plan.OwnerBirthUtcTicks, 42, 100, 101, 200, true, System.Threading.ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                42, plan.OwnerBirthUtcTicks, 42, 100, 100, 201, true, System.Threading.ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                42, plan.OwnerBirthUtcTicks, 42, 100, 100, 200, false, System.Threading.ApartmentState.STA));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequireOwnerIdentity(plan,
                42, plan.OwnerBirthUtcTicks, 42, 100, 100, 200, true, System.Threading.ApartmentState.MTA));
            OwnerGitQualificationManifest.RequirePolicy("Automatic");
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequirePolicy("ReadOnly"));
            Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.RequirePolicy("AskEachTime"));
        }

        [TestMethod]
        public void SnapshotDirectoryHashPinsRealSerializedBytesAndRejectsTamper()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-OwnerGit-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var snapshot = VbaGitSnapshotCoverageTests.LogicalForm(FormResourcePreflightTests.Resource());
                foreach (var file in snapshot.Serialize()) File.WriteAllBytes(Path.Combine(directory, file.Key), file.Value);
                string hash = OwnerGitQualificationManifest.SnapshotDirectoryHash(directory);
                Assert.AreEqual(OwnerGitQualificationManifest.SnapshotHash(snapshot), hash);
                Assert.IsTrue(OwnerGitQualificationManifest.ReadSnapshot(directory, hash).SameAs(snapshot));
                File.AppendAllText(Path.Combine(directory, "manifest.json"), " ");
                Assert.ThrowsException<InvalidOperationException>(() => OwnerGitQualificationManifest.ReadSnapshot(directory, hash));
            }
            finally { Directory.Delete(directory, true); }
        }
    }
}
