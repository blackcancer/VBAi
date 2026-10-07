using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class MacroGitOperationsRevisionObservationTests
    {
        private static string Sha(byte[] value)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(value)).Replace("-", "").ToLowerInvariant(); }
        private static string Sha(string value) => Sha(Encoding.UTF8.GetBytes(value));
        private static MacroGitOperations Create(out MacroGitRepository repository, out VbaGitProject project,
            out global::FakeProject host, List<string> calls)
        {
            string root = Path.Combine(Path.GetTempPath(), "RevisionObservation-" + Guid.NewGuid().ToString("N"));
            repository = new MacroGitRepository(root, "main");
            repository.CommandOverride = (args, input, use, allow) =>
            {
                calls.Add(string.Join(" ", args));
                string reference = args.Length == 4 ? args[3] : null;
                string value = reference == "refs/heads/main" ? "head-used" : reference == "refs/codex/baseline" ? "baseline-used" :
                    reference == "refs/codex/backup" ? "backup-used" : reference == "refs/codex/after-import" ? "after-used" : null;
                return new MacroGitRepository.Result { Bytes = Encoding.UTF8.GetBytes(value ?? ""), ExitCode = value == null ? 1 : 0 };
            };
            host = new global::FakeProject { FileName = Path.Combine(root, "revision-observation.xlsm") };
            host.VBComponents.Add(new global::FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nPublic Const Value = 1\n"));
            var bound = host;
            project = new VbaGitProject(() => bound, host.FileName);
            return new MacroGitOperations(project, repository);
        }

        [TestMethod]
        public void ObservationReportsExactlyHashedFilesAndOperandsWithoutMoreReferenceReads()
        {
            var calls = new List<string>();
            MacroGitRepository repository; VbaGitProject project; global::FakeProject host;
            using (var operations = Create(out repository, out project, out host, calls))
            {
                Assert.IsNull(operations.ObserveRevision);
                var snapshot = project.Capture();
                string expected = operations.Revision(snapshot);
                string[] expectedCalls = calls.ToArray(); calls.Clear();
                MacroGitOperations.RevisionDiagnostic observed = null;
                operations.ObserveRevision = value => observed = value;
                Assert.AreEqual(expected, operations.Revision(snapshot));
                CollectionAssert.AreEqual(expectedCalls, calls.ToArray());
                CollectionAssert.AreEqual(new[] {
                    "rev-parse --verify --quiet refs/codex/backup", "rev-parse --verify --quiet refs/codex/after-import",
                    "rev-parse --verify --quiet refs/heads/main", "rev-parse --verify --quiet refs/codex/baseline"
                }, expectedCalls);
                Assert.IsNotNull(observed); Assert.AreEqual(expected, observed.Revision);
                var files = snapshot.ComparisonFiles().ToArray();
                Assert.AreEqual(files.Length, observed.TotalFiles); Assert.IsFalse(observed.FilesTruncated);
                for (int i = 0; i < files.Length; i++)
                {
                    byte[] key = Encoding.UTF8.GetBytes(files[i].Key + "\0" + files[i].Value.Length + "\0");
                    Assert.AreEqual(Sha(key), observed.Files[i].KeySha256);
                    Assert.AreEqual(key.Length, observed.Files[i].KeyBytes);
                    Assert.AreEqual(Sha(files[i].Value), observed.Files[i].ContentSha256);
                    Assert.AreEqual(files[i].Value.Length, observed.Files[i].ContentBytes);
                }
                Assert.AreEqual(Sha("main"), observed.Branch.Sha256);
                Assert.AreEqual(Sha("head-used"), observed.Head.Sha256);
                Assert.AreEqual(Sha("baseline-used"), observed.Baseline.Sha256);
                Assert.AreEqual(Sha("null"), observed.PendingMerge.Sha256);
                Assert.AreEqual(Sha("backup-used"), observed.Backup.Sha256);
                Assert.AreEqual(Sha("after-used"), observed.After.Sha256);
                Assert.IsTrue(observed.Marker.IsNull); Assert.AreEqual(0, observed.Marker.Bytes);
                Assert.AreEqual(Sha("main\0head-used\0baseline-used\0null\0backup-used\0after-used\0absent"), observed.StateSha256);
            }
        }

        [TestMethod]
        public void ThrowingObserverCannotChangeRevisionOrReferenceReadSequence()
        {
            var calls = new List<string>();
            MacroGitRepository repository; VbaGitProject project; global::FakeProject host;
            using (var operations = Create(out repository, out project, out host, calls))
            {
                var snapshot = project.Capture(); string expected = operations.Revision(snapshot);
                string[] original = calls.ToArray(); calls.Clear(); int callbacks = 0;
                operations.ObserveRevision = value => { callbacks++; throw new IOException("Diagnostic sink unavailable"); };
                Assert.AreEqual(expected, operations.Revision(snapshot));
                Assert.AreEqual(1, callbacks); CollectionAssert.AreEqual(original, calls.ToArray());
            }
        }

        [TestMethod]
        public async Task ChangedStateStillRefusesBeforeBranchMutationWhenObserverThrows()
        {
            var calls = new List<string>();
            MacroGitRepository repository; VbaGitProject project; global::FakeProject host;
            using (var operations = Create(out repository, out project, out host, calls))
            {
                string expected = operations.Revision(project.Capture()); calls.Clear();
                host.VBComponents.Item("Module1").CodeModule.Text += "Public Const Changed = 2\n";
                operations.ObserveRevision = value => { throw new IOException("Synthetic recorder failure"); };
                var failure = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() =>
                    operations.ExecuteAsync("branch_create", expectedState: expected, name: "must-not-exist"));
                Assert.AreEqual(UiText.Get("The Git/VBA state changed. Read git_status again before making changes."), failure.Message);
                Assert.AreEqual(4, calls.Count);
                Assert.IsTrue(calls.All(x => x.StartsWith("rev-parse ", StringComparison.Ordinal)));
            }
        }

        [TestMethod]
        public void FileEvidenceIsBoundedWithoutTruncatingTheActualRevisionInput()
        {
            var calls = new List<string>();
            MacroGitRepository repository; VbaGitProject project; global::FakeProject host;
            using (var operations = Create(out repository, out project, out host, calls))
            {
                var components = new List<VbaGitComponent>(); var files = new Dictionary<string, byte[]>();
                for (int i = 0; i <= MacroGitOperations.RevisionDiagnostic.MaximumFiles; i++)
                {
                    string name = "Module" + i;
                    components.Add(new VbaGitComponent { Name = name, Type = 1 });
                    files.Add(name + ".bas", Encoding.UTF8.GetBytes("Attribute VB_Name = \"" + name + "\"\n"));
                }
                var snapshot = new VbaGitSnapshot(new VbaGitManifest { Components = components.ToArray(), References = "" }, files);
                string expected = operations.Revision(snapshot); MacroGitOperations.RevisionDiagnostic observed = null;
                operations.ObserveRevision = value => observed = value;
                Assert.AreEqual(expected, operations.Revision(snapshot));
                Assert.AreEqual(128, observed.Files.Length); Assert.AreEqual(130, observed.TotalFiles);
                Assert.IsTrue(observed.FilesTruncated);
                files["Module99.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module99\"\nPublic Const Changed = 1\n");
                var changed = new VbaGitSnapshot(snapshot.Manifest, files);
                Assert.AreNotEqual(expected, operations.Revision(changed));
            }
        }
    }
}