using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormGitHubTests
    {
        // Keep exact owned RCWs reachable when emitted native work has an uncertain outcome.
        private static readonly List<ExcelVbeFixture> RetainedSequentialHosts = new List<ExcelVbeFixture>();

        /// <summary>Qualifies the fixed synthetic remote with sequential owned hosts and native save/reopen, without concurrent Excel instances.</summary>
        [STATestMethod, TestCategory("NativeUserFormExplicitBootstrap")]
        public void ExplicitSequentialUserFormPushFetchImportAndReopenPreserveNativeState()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_GITHUB_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_EXPLICIT_BOOTSTRAP") != "1")
                Assert.Inconclusive("Separate authenticated synthetic repository and explicit owned Excel launch opt-ins are required.");
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Owned Excel automation requires its separate opt-in.");
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName)))
                throw new InvalidOperationException("This matrix does not allow a path/token diagnostic manifest.");
            string manifestPath = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_TEST_GITHUB_MANIFEST"));
            var manifest = Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(manifestPath));
            string remote = Convert.ToString(manifest["repositoryUrl"]) + ".git";
            Assert.AreEqual("https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git", remote);
            Assert.AreEqual("1396566119", Convert.ToString(manifest["repositoryId"]));
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_GIT_OUTPUT"));
            string output = Path.Combine(root, "explicit-sequential-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            string branch = "qualification-userform-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var report = new Dictionary<string, object> {
                ["Stage"] = "preflight", ["Branch"] = branch, ["Remote"] = remote,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["PublishedMain"] = false, ["ProductionMacroExecuted"] = false,
                ["Scope"] = "Sequential explicit fresh owned Excel; production capture, exact synthetic GitHub transfer, guarded import/backup, helper save/reopen and normal exit. No embedded Git UI claim.",
                ["StartedUtc"] = DateTime.UtcNow.ToString("o")
            };
            var previousContext = SynchronizationContext.Current;
            using (var dispatcher = new Control())
            using (var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
            {
                dispatcher.CreateControl(); SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                Exception primaryFailure = null;
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
                        const string form = "QualificationForm";
                        VbaGitSnapshot captured = null;
                        IDictionary<string, object> sourceView = null;
                        report["Stage"] = "source-native-capture"; WriteReport(output, report);
                        SequentialHost(source => {
                            report["SourcePid"] = source.ProcessId;
                            report["SourceFixtureRoot"] = source.Root;
                            string sourcePath = source.File("explicit-userform-source.xlsm");
                            source.PrepareGitForm(form, "Synthetic remote form", "USERFORM_" + Guid.NewGuid().ToString("N"), sourcePath);
                            File.Copy(sourcePath, Path.Combine(output, "source-before-import.xlsm"));
                            source.WithGitProject(sourcePath, project => {
                                captured = project.Capture();
                                Assert.IsTrue(captured.Manifest.Components.Single(item => item.Name == form).HasResources);
                                Assert.IsTrue(captured.SameAs(project.Capture()), "Unchanged source revision must be stable before transport.");
                                SaveSnapshot(output, "source-export", captured);
                                sourceView = source.ReadGitForm(form);
                                source.CaptureGitFormDesigner(form, Path.Combine(output, "source-designer.png"));
                            });
                        }, shutdown => { report["SourceShutdown"] = shutdown; });
                        report["SourceNormalShutdownVerified"] = true;
                        var incomplete = captured.Serialize(); incomplete.Remove(form + ".frx");
                        Assert.ThrowsException<InvalidOperationException>(() => VbaGitSnapshot.Read(incomplete));
                        var corrupt = captured.Serialize(); corrupt[form + ".frx"] = new byte[0];
                        Assert.ThrowsException<InvalidOperationException>(() => VbaGitSnapshot.Read(corrupt));
                        report["MissingAndEmptyFrxRejectedBeforeTargetLaunch"] = true;
                        report["ExportedFiles"] = Describe(captured);
                        var repository = NewRepository(Path.Combine(output, "source.git"), branch, remote, account);
                        RunGit(repository, new[] { "fetch", "--no-tags", "origin", "refs/heads/main:refs/remotes/origin/baseline" });
                        string parent = repository.Resolve("refs/remotes/origin/baseline");
                        Assert.AreEqual(Convert.ToString(manifest["mainCommit"]), parent);
                        string commit = repository.Commit(captured, parent, "Synthetic sequential native UserForm qualification");
                        repository.SetRef(repository.Head, commit);
                        VerifyRepository(api, manifest, deadline.Token);
                        report["Stage"] = "push-once"; report["Commit"] = commit; WriteReport(output, report);
                        repository.Push(commit);
                        report["BranchPublished"] = true;
                        var fetchedRepository = NewRepository(Path.Combine(output, "target.git"), branch, remote, account);
                        Assert.AreEqual(commit, fetchedRepository.Fetch());
                        var fetched = fetchedRepository.Read(commit);
                        CollectionAssert.AreEquivalent(captured.Files.Keys.ToArray(), fetched.Files.Keys.ToArray());
                        foreach (var file in captured.Files) CollectionAssert.AreEqual(file.Value, fetched.Files[file.Key], "Exact remote bytes: " + file.Key);
                        Assert.IsTrue(captured.SameAs(fetched));
                        SaveSnapshot(output, "fetched", fetched); report["RemoteFiles"] = Describe(fetched);
                        report["Stage"] = "target-native-import"; WriteReport(output, report);
                        SequentialHost(target => {
                            report["TargetPid"] = target.ProcessId;
                            report["TargetFixtureRoot"] = target.Root;
                            Assert.AreNotEqual(report["SourceFixtureRoot"], target.Root, "Distinct owned bootstrap identities are required, independently of PID reuse.");
                            string targetPath = target.File("explicit-userform-target.xlsm");
                            target.PrepareGitForm(form, "Existing target sentinel", "SENTINEL", targetPath);
                            File.Copy(targetPath, Path.Combine(output, "target-before-import.xlsm"));
                            target.WithGitProject(targetPath, project => {
                                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                                var before = project.Capture(); SaveSnapshot(output, "target-backup", before);
                                using (var operations = new MacroGitOperations(project, fetchedRepository))
                                {
                                    ExecuteGuardedImport(() => Await(operations.ExecuteAsync("pull", operations.Revision(before))),
                                        () => fetchedRepository.RecoveryPending, () => RetainSequentialHost(target));
                                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                                    Assert.IsFalse(fetchedRepository.RecoveryPending);
                                    Assert.IsTrue(fetchedRepository.Read(fetchedRepository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                                    Assert.IsTrue(project.Capture().SameAs(captured));
                                    SaveSnapshot(output, "imported", project.Capture());
                                }
                                AssertExactNativeForm(sourceView, target.ReadGitForm(form));
                            });
                            report["ImportAndBackupVerified"] = true;
                            report["Stage"] = "helper-save-reopen"; WriteReport(output, report);
                            target.SaveAndReopenGitLayout(targetPath);
                            AssertExactNativeForm(sourceView, target.ReadGitForm(form));
                            target.WithGitProject(targetPath, project => {
                                var reopened = project.Capture(); Assert.IsTrue(reopened.SameAs(captured));
                                SaveSnapshot(output, "reopened", reopened);
                            });
                            File.Copy(targetPath, Path.Combine(output, "saved-reopened.xlsm"));
                            target.CaptureGitFormDesigner(form, Path.Combine(output, "reopened-designer.png"));
                            report["HelperSaveReopenVerified"] = true;
                        }, shutdown => { report["TargetShutdown"] = shutdown; });
                        report["TargetNormalShutdownVerified"] = true;
                        VerifyRepository(api, manifest, deadline.Token);
                        report["MainUnchanged"] = true;
                        report["DesignerCaptureAcceptance"] = "CAPTURED_PENDING_VISUAL_REVIEW";
                    }
                    report["Stage"] = "PASS";
                }
                catch (Exception error)
                {
                    primaryFailure = error;
                    report["Failure"] = error.ToString();
                    if (Equals(report["Stage"], "push-once")) report["RemotePushOutcomeMayBeUncertain"] = true;
                }
                finally
                {
                    CompleteSequentialEvidence(primaryFailure,
                        () => { report["FinishedUtc"] = DateTime.UtcNow.ToString("o"); WriteReport(output, report); },
                        () => SynchronizationContext.SetSynchronizationContext(previousContext));
                }
            }
        }

        private static void AssertExactNativeForm(IDictionary<string, object> expected, IDictionary<string, object> actual)
        {
            CollectionAssert.AreEquivalent(expected.Keys.ToArray(), actual.Keys.ToArray());
            foreach (var value in expected) Assert.AreEqual(value.Value, actual[value.Key], "Native form: " + value.Key);
        }

        private static void SequentialHost(Action<ExcelVbeFixture> scenario, Action<object> shutdown)
        {
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"));
            Directory.CreateDirectory(root);
            var host = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(root, "unused-inspection-" + Guid.NewGuid().ToString("N") + ".jsonl"));
            Exception failure = null;
            try { scenario(host); }
            catch (Exception error)
            {
                failure = error;
                if (host.PreserveForDiagnosticRecovery || HasUncertainSequentialDelivery(error))
                {
                    RetainSequentialHost(host);
                    ExceptionDispatchInfo.Capture(error).Throw(); // No further native call or cleanup.
                }
            }
            try { host.Dispose(); shutdown(host.ShutdownDiagnostics); }
            catch (Exception cleanup)
            {
                if (failure != null) throw new AggregateException("Sequential native scenario and shutdown failed; neither is replayed.", failure, cleanup);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private static void RetainSequentialHost(ExcelVbeFixture host)
        {
            host.PreserveForDiagnosticRecovery = true;
            if (!RetainedSequentialHosts.Contains(host)) RetainedSequentialHosts.Add(host);
        }

        internal static void ExecuteGuardedImport(Action operation, Func<bool> recoveryPending, Action retain)
        {
            try { operation(); }
            catch (Exception primary)
            {
                bool mustRetain = HasUncertainSequentialDelivery(primary);
                var errors = new List<Exception> { primary };
                if (!mustRetain)
                    try { mustRetain = recoveryPending(); }
                    catch (Exception markerFailure) { mustRetain = true; errors.Add(markerFailure); }
                if (mustRetain)
                    try { retain(); }
                    catch (Exception retentionFailure) { errors.Add(retentionFailure); }
                if (errors.Count > 1) throw new AggregateException("Import, recovery observation and retention errors remain separate; no native replay.", errors);
                throw;
            }
        }

        internal static void CompleteSequentialEvidence(Exception primary, Action persist, Action restoreContext)
        {
            var errors = new List<Exception>();
            if (primary != null) errors.Add(primary);
            try { persist(); } catch (Exception recording) { errors.Add(recording); }
            finally { try { restoreContext(); } catch (Exception contextFailure) { errors.Add(contextFailure); } }
            if (errors.Count > 1) throw new AggregateException("Scenario, final evidence and context restoration errors remain separate.", errors);
            if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
        }

        internal static bool HasUncertainSequentialDelivery(Exception error)
        {
            if (error == null) return false;
            if (error is TimeoutException || error is IOException || error is OperationCanceledException) return true;
            if (error is AggregateException aggregate && aggregate.InnerExceptions.Any(HasUncertainSequentialDelivery)) return true;
            return HasUncertainSequentialDelivery(error.InnerException);
        }
    }
}
