using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in synthetic UserForm transfer through the retained private qualification repository.</summary>
    [TestClass, TestCategory("AuthenticatedIntegration"), DoNotParallelize]
    public sealed class NativeUserFormGitHubTests
    {
        public TestContext TestContext { get; set; }
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        [STATestMethod]
        public void OwnedUserFormPushFetchImportPreservesControlsCodeAndFrx()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_GITHUB_TESTS") != "1")
                Assert.Inconclusive("Explicit disposable Excel UserForm and retained synthetic GitHub repository opt-in required.");
            string manifestPath = Environment.GetEnvironmentVariable("VBAi_TEST_GITHUB_MANIFEST");
            Assert.IsFalse(string.IsNullOrWhiteSpace(manifestPath));
            var manifest = Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(manifestPath));
            string remote = Convert.ToString(manifest["repositoryUrl"]) + ".git";
            const string allowedRepository = "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git";
            Assert.AreEqual(allowedRepository, remote, "Only the explicitly retained disposable repository is authorized.");
            Assert.AreEqual("1396566119", Convert.ToString(manifest["repositoryId"]));
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_GIT_OUTPUT");
            Assert.IsTrue(!string.IsNullOrWhiteSpace(output) && Path.IsPathRooted(output));
            Directory.CreateDirectory(output);
            string branch = "qualification-userform-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var report = new Dictionary<string, object> { ["Stage"] = "preflight", ["Branch"] = branch, ["Remote"] = remote,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ["PublishedMain"] = false,
                ["ProductionMacroExecuted"] = false, ["StartedUtc"] = DateTime.UtcNow.ToString("o") };
            var previousContext = SynchronizationContext.Current;
            using (var dispatcher = new Control())
            using (var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
            {
                dispatcher.CreateControl();
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                try
                {
                    var accounts = Await(new GitHubAccountService().ListAsync(deadline.Token));
                    string account = LlmSettings.Load().GitHubAccount;
                    if (string.IsNullOrWhiteSpace(account)) account = accounts.Single();
                    Assert.AreEqual("blackcancer", account, true);
                    using (var api = new GitHubApi(account))
                    {
                        VerifyRepository(api, manifest, deadline.Token);
                        Assert.IsFalse(Await(api.Branches(remote, deadline.Token)).Any(item => item.name == branch));
                        ExcelVbeFixture.Run(source => ExcelVbeFixture.Run(target => {
                            report["SourcePid"] = source.ProcessId; report["TargetPid"] = target.ProcessId;
                            string sourcePath = source.File("userform-source.xlsm"), targetPath = target.File("userform-target.xlsm");
                            const string form = "QualificationForm";
                            string marker = "USERFORM_" + Guid.NewGuid().ToString("N");
                            source.PrepareGitForm(form, "Synthetic remote form", marker, sourcePath);
                            target.PrepareGitForm(form, "Existing target sentinel", "SENTINEL", targetPath);
                            // Preserve real host files before any import. No helper save is used as an import oracle.
                            File.Copy(sourcePath, Path.Combine(output, "source-before-import.xlsm"));
                            File.Copy(targetPath, Path.Combine(output, "target-before-import.xlsm"));
                            source.WithGitProject(sourcePath, sourceProject => target.WithGitProject(targetPath, targetProject => {
                                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                                var captured = sourceProject.Capture(); var before = targetProject.Capture();
                                Assert.IsTrue(captured.Manifest.Components.Single(item => item.Name == form).HasResources);
                                Assert.IsTrue(captured.Files[form + ".frx"].Length > 0);
                                report["ExportedFiles"] = Describe(captured);
                                SaveSnapshot(output, "source-export", captured); SaveSnapshot(output, "target-backup", before);
                                var sourceView = source.ReadGitForm(form);
                                source.CaptureGitFormDesigner(form, Path.Combine(output, "source-designer.png"));
                                report["DesignerCaptureAcceptance"] = "CAPTURED_PENDING_VISUAL_REVIEW";
                                report["Stage"] = "missing-frx-refusal"; WriteReport(output, report);
                                var incomplete = captured.Serialize(); incomplete.Remove(form + ".frx");
                                Assert.ThrowsException<InvalidOperationException>(() => VbaGitSnapshot.Read(incomplete));
                                Assert.IsTrue(targetProject.Capture().SameAs(before));
                                report["MissingFrxRejectedBeforeImport"] = true;
                                // Record opaque-resource validation separately; do not feed corrupt bytes to Office in this valid-transfer test.
                                var corrupt = captured.Serialize(); corrupt[form + ".frx"] = new byte[0];
                                bool rejected = false;
                                try { VbaGitSnapshot.Read(corrupt); } catch (InvalidOperationException) { rejected = true; }
                                report["EmptyFrxRejectedBySnapshotPreflight"] = rejected;
                                var repository = NewRepository(Path.Combine(output, "source.git"), branch, remote, account);
                                RunGit(repository, new[] { "fetch", "--no-tags", "origin", "refs/heads/main:refs/remotes/origin/baseline" });
                                string parent = repository.Resolve("refs/remotes/origin/baseline");
                                Assert.AreEqual(Convert.ToString(manifest["mainCommit"]), parent);
                                string commit = repository.Commit(captured, parent, "Synthetic native UserForm qualification");
                                repository.SetRef(repository.Head, commit);
                                VerifyRepository(api, manifest, deadline.Token);
                                report["Stage"] = "push-once"; report["Commit"] = commit; WriteReport(output, report);
                                repository.Push(commit);
                                report["BranchPublished"] = true;
                                var fetchedRepository = NewRepository(Path.Combine(output, "target.git"), branch, remote, account);
                                Assert.AreEqual(commit, fetchedRepository.Fetch());
                                var fetched = fetchedRepository.Read(commit);
                                Assert.IsTrue(captured.SameAs(fetched));
                                CollectionAssert.AreEqual(captured.Files[form + ".frx"], fetched.Files[form + ".frx"]);
                                report["RemoteFiles"] = Describe(fetched);
                                report["Stage"] = "production-pull-once"; WriteReport(output, report);
                                using (var operations = new MacroGitOperations(targetProject, fetchedRepository))
                                {
                                    Await(operations.ExecuteAsync("pull", operations.Revision(before)));
                                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                                    Assert.IsFalse(fetchedRepository.RecoveryPending);
                                    Assert.IsTrue(fetchedRepository.Read(fetchedRepository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                                }
                                var imported = targetProject.Capture();
                                Assert.IsTrue(imported.SameAs(captured));
                                CollectionAssert.AreEqual(captured.ComparisonFiles()[form + ".frx"], imported.ComparisonFiles()[form + ".frx"]);
                                // Native re-export may change CFB allocation/timestamps and documented
                                // padding. The fetched transport bytes above still must be exact.
                                report["RawReexportFrxMatchesTransport"] = captured.Files[form + ".frx"].SequenceEqual(imported.Files[form + ".frx"]);
                                report["LogicalReexportFrxVerified"] = true;
                                var targetView = target.ReadGitForm(form);
                                target.CaptureGitFormDesigner(form, Path.Combine(output, "target-designer.png"));
                                foreach (string key in sourceView.Keys) Assert.AreEqual(sourceView[key], targetView[key], "Native form mismatch: " + key);
                                StringAssert.Contains(Convert.ToString(targetView["Code"]), marker);
                                Assert.AreEqual(2, Convert.ToInt32(targetView["ControlCount"]));
                                SaveSnapshot(output, "target-reexport", imported);
                                report["ImportedFiles"] = Describe(imported); report["NativeControlsAndCodeVerified"] = true;
                                report["BackupVerified"] = true; report["RecoveryPending"] = fetchedRepository.RecoveryPending;
                                VerifyRepository(api, manifest, deadline.Token);
                                report["MainUnchanged"] = true;
                                report["Stage"] = "native-roundtrip-verified-awaiting-shutdown"; WriteReport(output, report);
                            }));
                        }));
                    }
                    report["NormalShutdownVerified"] = true; report["Stage"] = "PASS";
                    TestContext.WriteLine("PASS: native form export, dedicated private branch, fetch/import, controls/code/FRX and normal shutdown. Main unchanged.");
                }
                catch (Exception error) { report["Failure"] = error.ToString(); throw; }
                finally { report["FinishedUtc"] = DateTime.UtcNow.ToString("o"); WriteReport(output, report); SynchronizationContext.SetSynchronizationContext(previousContext); }
            }
        }

        [STATestMethod]
        public void CorruptFrxRefusesBeforeReplacingOwnedFormOrReportsRecovery()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_CORRUPTION_TESTS") != "1")
                Assert.Inconclusive("Separate explicit disposable-form corruption diagnostic opt-in required.");
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_GIT_OUTPUT");
            Assert.IsTrue(!string.IsNullOrWhiteSpace(output) && Path.IsPathRooted(output));
            output = Path.Combine(output, "corruption"); Directory.CreateDirectory(output);
            var report = new Dictionary<string, object> { ["Stage"] = "preflight", ["RemoteMutation"] = false,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ["Corruption"] = "Empty existing FRX companion" };
            var previousContext = SynchronizationContext.Current;
            using (var dispatcher = new Control())
            {
                dispatcher.CreateControl(); SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                try
                {
                    ExcelVbeFixture.Run(host => {
                        report["Pid"] = host.ProcessId;
                        string path = host.File("corruption-sentinel.xlsm"); const string form = "QualificationForm";
                        host.PrepareGitForm(form, "Preserve this sentinel", "CORRUPTION_SENTINEL", path);
                        File.Copy(path, Path.Combine(output, "sentinel-before-import.xlsm"));
                        host.WithGitProject(path, project => {
                            var before = project.Capture(); SaveSnapshot(output, "backup-export", before);
                            var files = before.Serialize(); files[form + ".frx"] = new byte[0];
                            VbaGitSnapshot corrupt;
                            try { corrupt = VbaGitSnapshot.Read(files); }
                            catch (InvalidOperationException)
                            {
                                Assert.IsTrue(project.Capture().SameAs(before)); report["PreflightRejected"] = true;
                                report["ChangedBeforeRecovery"] = false; return;
                            }
                            report["PreflightRejected"] = false;
                            var repository = NewRepository(Path.Combine(output, "diagnostic.git"), "qualification-corrupt-local-only",
                                "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git", null);
                            string checkpoint = repository.Checkpoint(corrupt, "Local-only empty FRX diagnostic").Id;
                            Exception importFailure = null;
                            using (var operations = new MacroGitOperations(project, repository))
                            {
                                report["Stage"] = "corrupt-import-once-backup-prepared-by-production"; WriteReport(output, report);
                                try { Await(operations.ExecuteAsync("checkpoint_restore", operations.Revision(before), name: checkpoint)); }
                                catch (Exception error) { importFailure = error; report["ImportError"] = error.ToString(); }
                                report["RecoveryPending"] = repository.RecoveryPending;
                                var backup = repository.Read(repository.Resolve(MacroGitRepository.Backup));
                                report["BackupVerified"] = backup != null && backup.SameAs(before);
                                var actual = project.Capture();
                                bool changed = !actual.SameAs(before); report["ChangedBeforeRecovery"] = changed;
                                SaveSnapshot(output, "actual-after-attempt", actual); WriteReport(output, report);
                                if (repository.RecoveryPending)
                                {
                                    var recorded = repository.Read(repository.Resolve(MacroGitRepository.AfterImport));
                                    Assert.IsTrue(recorded != null && recorded.SameAs(actual), "Unknown current state; recovery was not attempted.");
                                    Assert.IsTrue(backup != null && backup.SameAs(before), "Backup mismatch; recovery was not attempted.");
                                    // Explicit recovery of the measured post-import state, never a retry of the corrupt import.
                                    Await(operations.ExecuteAsync("rollback", operations.Revision(actual)));
                                    report["ExplicitRecoveryVerified"] = project.Capture().SameAs(before);
                                    Assert.AreEqual(true, report["ExplicitRecoveryVerified"]);
                                }
                                Assert.IsNotNull(importFailure, "Corrupt resource import was accepted; preflight is insufficient.");
                                Assert.IsFalse(changed, "Corrupt FRX changed the live project before recovery. Recoverability does not satisfy refusal-before-overwrite.");
                            }
                        });
                    });
                    report["NormalShutdownVerified"] = true; report["Stage"] = "PASS";
                }
                catch (Exception error) { report["Failure"] = error.ToString(); throw; }
                finally { WriteReport(output, report); SynchronizationContext.SetSynchronizationContext(previousContext); }
            }
        }

        private static MacroGitRepository NewRepository(string path, string branch, string remote, string account)
        {
            var repository = new MacroGitRepository(path, branch, account); repository.Initialize(remote);
            RunGit(repository, new[] { "config", "user.name", "VBAi Qualification" });
            RunGit(repository, new[] { "config", "user.email", "qualification@example.invalid" });
            return repository;
        }
        private static void RunGit(MacroGitRepository repository, string[] arguments)
        {
            typeof(MacroGitRepository).GetMethod("Run", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(repository, new object[] { arguments, null, true, false });
        }
        private static void VerifyRepository(GitHubApi api, IDictionary<string, object> manifest, CancellationToken token)
        {
            const string route = "/repos/blackcancer/vbai-qualification-20260929203712-7267b1e6";
            var repository = Await(api.Request<Dictionary<string, object>>(HttpMethod.Get, route, null, token));
            Assert.AreEqual(true, repository["private"]); Assert.AreEqual(Convert.ToString(manifest["repositoryId"]), Convert.ToString(repository["id"]));
            Assert.AreEqual("blackcancer/vbai-qualification-20260929203712-7267b1e6", repository["full_name"]);
            var main = Await(api.Request<Dictionary<string, object>>(HttpMethod.Get, route + "/git/ref/heads/main", null, token));
            Assert.AreEqual(Convert.ToString(manifest["mainCommit"]), Convert.ToString(VbeBridgeClient.Object(main["object"])["sha"]));
        }
        private static T Await<T>(Task<T> task)
        {
            var watch = Stopwatch.StartNew();
            while (!task.IsCompleted) { if (watch.Elapsed > TimeSpan.FromMinutes(3)) throw new TimeoutException("Qualification operation timed out; not retried."); Application.DoEvents(); Thread.Sleep(10); }
            return task.GetAwaiter().GetResult();
        }
        private static object[] Describe(VbaGitSnapshot snapshot)
        {
            return snapshot.Serialize().Select(file => { using (var hash = SHA256.Create()) return (object)new {
                Path = file.Key, Bytes = file.Value.Length, Sha256 = BitConverter.ToString(hash.ComputeHash(file.Value)).Replace("-", "") }; }).ToArray();
        }
        private static void SaveSnapshot(string output, string name, VbaGitSnapshot snapshot)
        {
            string root = Path.Combine(output, name); Directory.CreateDirectory(root);
            foreach (var file in snapshot.Serialize()) File.WriteAllBytes(Path.Combine(root, file.Key), file.Value);
        }
        private static void WriteReport(string output, IDictionary<string, object> report)
        { File.WriteAllText(Path.Combine(output, "userform-github.json"), Json.Serialize(report)); }
    }
}
