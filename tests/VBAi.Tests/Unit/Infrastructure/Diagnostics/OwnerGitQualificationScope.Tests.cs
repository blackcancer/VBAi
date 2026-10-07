using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises immutable publication, emission ordering and uncertain-outcome retention without Office.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OwnerGitQualificationScopeTests
    {
        private const string Sha = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        private string root, evidence, previous;

        [TestInitialize]
        public void Initialize()
        {
            previous = Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName);
            Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, null);
            root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            evidence = Path.Combine(root, "LabelButton-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(evidence, "before"));
            Directory.CreateDirectory(Path.Combine(evidence, "changed"));
            WriteSnapshot("before");
            WriteSnapshot("changed");
            Directory.CreateDirectory(Path.Combine(evidence, "local.git"));
        }

        private void WriteSnapshot(string name)
        {
            var snapshot = new VbaGitSnapshot(new VbaGitManifest
            {
                References = "[]",
                Components = new[] { new VbaGitComponent { Name = "Module1", Type = 1 } }
            }, new Dictionary<string, byte[]>
            {
                ["Module1.bas"] = System.Text.Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n' " + name + "\r\n")
            });
            foreach (var file in snapshot.Serialize()) File.WriteAllBytes(Path.Combine(evidence, name, file.Key), file.Value);
        }

        [TestCleanup]
        public void Cleanup()
        {
            Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, previous);
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }

        private OwnerGitQualificationManifest Plan()
        {
            string before = Path.Combine(evidence, "before"), changed = Path.Combine(evidence, "changed");
            return new OwnerGitQualificationManifest
            {
                Version = 1,
                OwnerPid = 42,
                OwnerBirthUtcTicks = 638000000000000000,
                OwnerNativeTid = 77,
                VbeHandle = 99,
                AssemblyMvid = "00000000-0000-0000-0000-000000000001",
                AssemblySha256 = Sha,
                FixtureRoot = root,
                WorkbookPath = Path.Combine(root, "owned.xlsm"),
                Project = Path.Combine(root, "owned.xlsm"),
                EvidenceRoot = evidence,
                RepoRelativePath = "local.git",
                Branch = "qualification-layout",
                Steps = new[] {
                    OwnerGitQualificationScope.Step("checkpoint_restore", changed, before, "20261006000000000-abcdef12"),
                    OwnerGitQualificationScope.Step("controlled_interruption", before, changed),
                    OwnerGitQualificationScope.Step("rollback", changed, before) }
            };
        }

        [TestMethod]
        public void ScopePublishesOnceAndRestoresEnvironmentWithoutDeletingEvidence()
        {
            string manifestPath;
            using (var scope = new OwnerGitQualificationScope(evidence))
            {
                manifestPath = scope.ManifestPath;
                Assert.AreEqual(manifestPath, Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName));
                Assert.IsFalse(File.Exists(manifestPath), "The child captures an opt-in path before PID-dependent publication.");
                var plan = Plan();
                scope.Publish(plan);
                var read = OwnerGitQualificationManifest.Parse(File.ReadAllText(manifestPath));
                Assert.AreEqual(plan.OwnerBirthUtcTicks, read.OwnerBirthUtcTicks);
                Assert.AreEqual(plan.OwnerNativeTid, read.OwnerNativeTid);
                CollectionAssert.AreEqual(File.ReadAllBytes(manifestPath),
                    new System.Text.UTF8Encoding(false).GetBytes(new JavaScriptSerializer().Serialize(plan)));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Publish(plan));
            }
            Assert.IsNull(Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName));
            Assert.IsTrue(File.Exists(manifestPath), "Diagnostic evidence survives normal scope disposal.");
        }

        [TestMethod]
        public void WrongOrMissingEvidenceAndOverlappingOptInFailBeforePublication()
        {
            Assert.ThrowsException<ArgumentException>(() => new OwnerGitQualificationScope("relative"));
            Assert.ThrowsException<DirectoryNotFoundException>(() => new OwnerGitQualificationScope(Path.Combine(root, "absent")));
            using (var scope = new OwnerGitQualificationScope(evidence))
            {
                Assert.ThrowsException<InvalidOperationException>(() => new OwnerGitQualificationScope(evidence));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Publish(null));
                var plan = Plan(); plan.EvidenceRoot = root;
                Assert.ThrowsException<InvalidOperationException>(() => scope.Publish(plan));
                Assert.IsFalse(File.Exists(scope.ManifestPath));
                plan = Plan(); plan.Steps[0].Verb = "push";
                Assert.ThrowsException<ArgumentException>(() => scope.Publish(plan));
                Assert.IsFalse(File.Exists(scope.ManifestPath));
            }
        }

        [TestMethod]
        public void LongEvidenceRootFailsBeforeOptInOrOwnedHostLaunch()
        {
            string leaf = "LabelButton-" + Guid.NewGuid().ToString("N");
            int pad = 203 - root.Length - leaf.Length - 2;
            Assert.IsTrue(pad > 0 && pad < 200, "Synthetic root must fit a single local directory component.");
            string longEvidence = Path.Combine(root, new string('x', pad), leaf);
            Assert.AreEqual(203, longEvidence.Length);
            Directory.CreateDirectory(longEvidence);
            Assert.ThrowsException<ArgumentException>(() => new OwnerGitQualificationScope(longEvidence));
            Assert.IsNull(Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName));
            Assert.AreEqual(0, Directory.GetFiles(longEvidence).Length, "No manifest or intent was published.");
        }

        [TestMethod]
        public void ExistingManifestCannotBeOverwritten()
        {
            using (var scope = new OwnerGitQualificationScope(evidence))
            {
                File.WriteAllText(scope.ManifestPath, "existing evidence");
                Assert.ThrowsException<IOException>(() => scope.Publish(Plan()));
                Assert.AreEqual("existing evidence", File.ReadAllText(scope.ManifestPath));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(Guid.NewGuid().ToString("N"), Sha));
            }
        }

        [TestMethod]
        public void ConcurrentOptInChangeIsPreservedAndPreventsPublication()
        {
            var scope = new OwnerGitQualificationScope(evidence);
            Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, "other campaign");
            Assert.ThrowsException<InvalidOperationException>(() => scope.Publish(Plan()));
            Assert.ThrowsException<InvalidOperationException>(() => scope.Dispose());
            Assert.AreEqual("other campaign", Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName));
            scope.Dispose();
        }

        [TestMethod]
        public void ClaimsRequirePublishedExactRevisionAndOrderedOneShotSteps()
        {
            var scope = new OwnerGitQualificationScope(evidence);
            try
            {
                var plan = Plan();
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(plan.Steps[0].Id, Sha));
                scope.Publish(plan);
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(plan.Steps[1].Id, Sha));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(plan.Steps[0].Id, "bad revision"));
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(Guid.NewGuid().ToString("N"), Sha));
                scope.Claim(plan.Steps[0].Id, Sha);
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(plan.Steps[0].Id, Sha));
                scope.Claim(plan.Steps[1].Id, Sha);
                scope.Claim(plan.Steps[2].Id, Sha);
                Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(plan.Steps[2].Id, Sha));
            }
            finally { scope.Dispose(); }
            Assert.ThrowsException<InvalidOperationException>(() => scope.Claim(Guid.NewGuid().ToString("N"), Sha));
        }

        [TestMethod]
        public void SnapshotPinsAreRawBytesAndTargetChangesCannotBeHidden()
        {
            string before = Path.Combine(evidence, "before"), changed = Path.Combine(evidence, "changed");
            var step = OwnerGitQualificationScope.Step("controlled_interruption", before, changed);
            Assert.AreNotEqual(step.ExpectedSnapshotSha256, step.TargetSnapshotSha256);
            File.AppendAllText(Path.Combine(changed, "manifest.json"), " ");
            Assert.AreNotEqual(step.TargetSnapshotSha256, OwnerGitQualificationManifest.SnapshotDirectoryHash(changed));
            Assert.AreEqual(step.ExpectedSnapshotSha256, OwnerGitQualificationManifest.SnapshotDirectoryHash(before));
        }

        [DataTestMethod]
        [DataRow(true, "Succeeded", true, false, false)]
        [DataRow(true, "Succeeded", false, false, true)]
        [DataRow(true, "ExpectedPrewriteRefusal", false, false, true)]
        [DataRow(false, "ExpectedPrewriteRefusal", false, false, false)]
        [DataRow(false, "ExpectedPrewriteRefusal", true, false, true)]
        [DataRow(false, "ExpectedPrewriteRefusal", false, true, true)]
        [DataRow(false, "FailedBeforeMutation", false, false, false)]
        [DataRow(false, "FailedBeforeMutation", false, true, true)]
        [DataRow(false, "FailedAfterMutationAdmission", true, false, true)]
        [DataRow(false, "Succeeded", true, false, true)]
        public void OnlyClassifiedSettledOutcomesAllowNormalCleanup(bool ok, string outcome, bool started, bool pending, bool retain)
        {
            var reply = new Dictionary<string, object> { ["Ok"] = ok };
            var terminal = new Dictionary<string, object>
            {
                ["Outcome"] = outcome,
                ["MutationStarted"] = started,
                ["RecoveryPending"] = pending
            };
            Assert.AreEqual(retain, OwnerGitQualificationScope.MustRetain(reply, terminal));
        }

        [TestMethod]
        public void MissingOrUnknownTerminalFieldsRetainOriginalHost()
        {
            var reply = new Dictionary<string, object> { ["Ok"] = false };
            var terminal = new Dictionary<string, object> { ["Outcome"] = "FailedBeforeMutation", ["MutationStarted"] = false };
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(null, terminal));
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(reply, null));
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(new Dictionary<string, object>(), terminal));
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(reply, terminal));
            terminal["RecoveryPending"] = null;
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(reply, terminal));
            terminal["RecoveryPending"] = false; terminal["MutationStarted"] = "false";
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(reply, terminal));
            terminal.Remove("Outcome");
            Assert.IsTrue(OwnerGitQualificationScope.MustRetain(reply, terminal));
        }

        [TestMethod]
        public void TerminalMustBelongToEmittedStepAndHaveTypedOutcome()
        {
            const string id = "0123456789abcdef0123456789abcdef";
            var terminal = new Dictionary<string, object> { ["StepId"] = id, ["Outcome"] = "Succeeded", ["MutationStarted"] = true };
            OwnerGitQualificationScope.ValidateTerminal(id, terminal);
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateTerminal("other", terminal));
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateTerminal(id, null));
            terminal["MutationStarted"] = "true";
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateTerminal(id, terminal));
            terminal["MutationStarted"] = true; terminal["Outcome"] = true;
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateTerminal(id, terminal));
        }

        [TestMethod]
        public void SuccessfulBridgePayloadMustMatchStepReceiptAndActualResult()
        {
            const string id = "0123456789abcdef0123456789abcdef", receipt = "exact terminal path";
            var data = new Dictionary<string, object>
            {
                ["StepId"] = id,
                ["Outcome"] = "Succeeded",
                ["ReceiptPath"] = receipt,
                ["Result"] = new Dictionary<string, object>()
            };
            var reply = new Dictionary<string, object> { ["Ok"] = true, ["Data"] = data };
            Assert.AreSame(data, OwnerGitQualificationScope.ValidateSuccess(id, receipt, reply));
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateSuccess("other", receipt, reply));
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateSuccess(id, "wrong path", reply));
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateSuccess(id, receipt, null));
            data["Outcome"] = "ExpectedPrewriteRefusal";
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateSuccess(id, receipt, reply));
            data["Outcome"] = "Succeeded"; data["Result"] = null;
            Assert.ThrowsException<IOException>(() => OwnerGitQualificationScope.ValidateSuccess(id, receipt, reply));
        }
    }
}
