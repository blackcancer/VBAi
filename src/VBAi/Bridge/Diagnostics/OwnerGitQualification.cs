using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>One explicitly armed, disposable owner-STA Git qualification. No caller-supplied verbs or paths.</summary>
    internal sealed class OwnerGitQualification
    {

        /// <summary>Connected VBE session used to revalidate the live owner and approved project before each step.</summary>
        private readonly VbeSession session;

        /// <summary>Host process ID captured when the bridge session connected; PID reuse is rejected by the manifest birth-time check.</summary>
        private readonly int connectedPid;

        /// <summary>Environment-configured manifest path captured at connection; the fixture publishes plan contents only after native setup.</summary>
        private readonly string manifestPath;

        /// <summary>Parsed frozen plan retained after its one allowed load and validation.</summary>
        private OwnerGitQualificationManifest manifest;

        /// <summary>Digest of the manifest bytes loaded once; later steps require the source manifest to remain unchanged.</summary>
        private string manifestHash;

        /// <summary>Index of the next frozen step; advances only after the preceding step reaches a verified terminal receipt.</summary>
        private int nextStep;

        /// <summary>Prevents further plan execution after uncertain or post-mutation failure.</summary>
        private bool quarantined;

        /// <summary>Creates a qualification runner bound to the connected VBE session and host process identity.</summary>
        /// <param name="session">Live session used to verify the exact project and VBE owner.</param>
        /// <param name="connectedPid">Host process ID recorded when this bridge session was established.</param>
        internal OwnerGitQualification(VbeSession session, int connectedPid)
        {
            this.session = session;
            this.connectedPid = connectedPid;
            // Only the path is captured at connection; the owned fixture publishes the plan after native setup.
            manifestPath = Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName);
        }

        /// <summary>Executes the next frozen Git step after revalidating the owner, project, revision, and snapshot hashes.
        /// It records intent before dispatch, admits at most one native mutation, and writes a terminal receipt. An uncertain
        /// or post-mutation failure quarantines the plan instead of retrying the native operation.</summary>
        /// <param name="request">Bridge request whose action ID must equal the next unconsumed step and whose expected revision must match the live project.</param>
        /// <param name="requestJson">Original JSON checked for the exact diagnostic command, action ID, and expected revision fields.</param>
        /// <returns>Operation outcome and the path of its durable terminal receipt.</returns>
        internal async Task<object> ExecuteAsync(Request request, string requestJson)
        {
            OwnerGitQualificationManifest.RequireExactRequest(requestJson);
            if (string.IsNullOrWhiteSpace(manifestPath)) throw new InvalidOperationException("Owner Git qualification is disabled.");
            LoadOnce();
            RequireContext();
            var step = OwnerGitQualificationManifest.RequireStep(manifest, nextStep, quarantined, request.Action);

            var expected = OwnerGitQualificationManifest.ReadSnapshot(step.ExpectedSnapshotDirectory, step.ExpectedSnapshotSha256);
            var target = OwnerGitQualificationManifest.ReadSnapshot(step.TargetSnapshotDirectory, step.TargetSnapshotSha256);
            string repoPath = Path.Combine(manifest.EvidenceRoot, manifest.RepoRelativePath);
            OwnerGitQualificationManifest.RequireChild(manifest.EvidenceRoot, repoPath, true);
            string account = step.Verb == "pull" ? LlmSettings.Load().GitHubAccount : null;
            if (step.Verb == "pull" && string.IsNullOrWhiteSpace(account))
            {
                using (var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
                    account = (await new GitHubAccountService().ListAsync(deadline.Token)).Single();
                RequireContext();
            }
            if (step.Verb == "pull" && !string.Equals(account, "blackcancer", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Frozen synthetic GitHub account is not selected.");
            var repository = new MacroGitRepository(repoPath, manifest.Branch, account);
            if (manifest.RemoteUrl != null) repository.Initialize(manifest.RemoteUrl);
            else if (!File.Exists(Path.Combine(repoPath, "HEAD"))) throw new InvalidOperationException("Frozen local Git repository is absent.");
            if (repository.Branch != manifest.Branch) throw new InvalidOperationException("Frozen repository branch changed.");
            var project = session.GitProject(manifest.Project, manifest.WorkbookPath);
            using (var operations = new MacroGitOperations(project, repository))
            {
                var before = project.Capture();
                if (!before.SameAs(expected) ||
                    !string.Equals(operations.Revision(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Frozen VBA/Git state changed before owner step.");
                string backupBefore = repository.Resolve(MacroGitRepository.Backup);
                string afterImportBefore = repository.Resolve(MacroGitRepository.AfterImport);
                bool recoveryBefore = repository.RecoveryPending;
                if (step.Verb == "pull") operations.ExpectedIncomingCommit = manifest.RemoteCommit;

                string prefix = Path.Combine(manifest.EvidenceRoot, "owner-git-" + step.Id);
                var intent = new { StepId = step.Id, Step = nextStep, step.Verb, ManifestSha256 = manifestHash,
                    ExpectedState = request.ExpectedSha256, BeforeSnapshotSha256 = OwnerGitQualificationManifest.SnapshotHash(before),
                    CandidateMvid = manifest.AssemblyMvid, CandidateSha256 = manifest.AssemblySha256,
                    OwnerPid = manifest.OwnerPid, OwnerNativeTid = manifest.OwnerNativeTid, StartedUtc = DateTime.UtcNow.ToString("o") };
                WriteNew(prefix + ".intent.json", intent);
                bool mutationStarted = false;
                Exception primary = null;
                object result = null;
                try
                {
                    if (step.Verb == "checkpoint_restore" && step.ExpectedErrorSubstring == null)
                    {
                        var checkpoint = repository.Read(repository.CheckpointCommit(step.CheckpointId));
                        if (checkpoint == null || !checkpoint.SameAs(target))
                            throw new InvalidOperationException("Frozen checkpoint differs from target snapshot.");
                    }
                    if (step.Verb == "rollback")
                    {
                        var backup = repository.Read(repository.Resolve(MacroGitRepository.Backup));
                        if (!repository.RecoveryPending || backup == null || !backup.SameAs(target))
                            throw new InvalidOperationException("Frozen rollback backup changed.");
                    }
                    Action boundary = () => {
                        RequireContext();
                        OwnerGitQualificationManifest.ReadSnapshot(step.ExpectedSnapshotDirectory, step.ExpectedSnapshotSha256);
                        OwnerGitQualificationManifest.ReadSnapshot(step.TargetSnapshotDirectory, step.TargetSnapshotSha256);
                        OwnerGitQualificationManifest.RequireNoReparse(repoPath);
                        if (mutationStarted) throw new InvalidOperationException("Native mutation was already admitted.");
                        mutationStarted = true;
                        WriteNew(prefix + ".mutation.json", new { StepId = step.Id, NativeMutationAdmittedUtc = DateTime.UtcNow.ToString("o") });
                    };
                    if (step.Verb == "controlled_interruption")
                    {
                        repository.PrepareRecovery(expected);
                        project.Apply(target, expected, boundary);
                        var measured = project.Capture();
                        // Record the real state even if it differs from the target. Never retry Apply.
                        repository.RecordImportedState(measured);
                        if (!measured.SameAs(target)) throw new InvalidOperationException("Measured native after-state differs from frozen target.");
                        result = new { MeasuredAfterSha256 = OwnerGitQualificationManifest.SnapshotHash(measured) };
                    }
                    else
                    {
                        operations.ImportOwnerPreflight = boundary;
                        await operations.ExecuteAsync(step.Verb, request.ExpectedSha256, name: step.CheckpointId);
                    }
                    if (step.ExpectedErrorSubstring != null)
                        throw new InvalidOperationException("Declared malformed checkpoint was accepted unexpectedly.");
                    RequireContext();
                    var after = project.Capture();
                    if (!after.SameAs(target)) throw new InvalidOperationException("Native readback differs from frozen target.");
                    result = new { AfterSnapshotSha256 = OwnerGitQualificationManifest.SnapshotHash(after),
                        BackupCommit = repository.Resolve(MacroGitRepository.Backup),
                        AfterImportCommit = repository.Resolve(MacroGitRepository.AfterImport),
                        RecoveryPending = repository.RecoveryPending,
                        Measured = result };
                    nextStep++;
                    WriteNew(prefix + ".terminal.json", new { StepId = step.Id, Outcome = "Succeeded", MutationStarted = mutationStarted,
                        Result = result, FinishedUtc = DateTime.UtcNow.ToString("o") });
                    return new { StepId = step.Id, Outcome = "Succeeded", Result = result, ReceiptPath = prefix + ".terminal.json" };
                }
                catch (Exception error)
                {
                    primary = error;
                    bool expectedRefusal = false;
                    string refusalObservationError = null;
                    if (!mutationStarted && step.ExpectedErrorSubstring != null && error.Message.Contains(step.ExpectedErrorSubstring))
                    {
                        try
                        {
                            RequireContext();
                            expectedRefusal = IsProvenPrewriteRefusal(step.ExpectedErrorSubstring, error.Message, mutationStarted,
                                recoveryBefore, repository.RecoveryPending, backupBefore, repository.Resolve(MacroGitRepository.Backup),
                                afterImportBefore, repository.Resolve(MacroGitRepository.AfterImport), project.Capture().SameAs(before));
                        }
                        catch (Exception observation) { refusalObservationError = observation.ToString(); }
                    }
                    quarantined = !expectedRefusal;
                    if (expectedRefusal) nextStep++;
                    bool? pending = null;
                    string pendingError = null;
                    try { pending = repository.RecoveryPending; } catch (Exception observation) { pendingError = observation.ToString(); quarantined = true; }
                    try { WriteNew(prefix + ".terminal.json", new { StepId = step.Id,
                        Outcome = expectedRefusal ? "ExpectedPrewriteRefusal" : mutationStarted ? "FailedAfterMutationAdmission" : "FailedBeforeMutation",
                        MutationStarted = mutationStarted, Error = primary.ToString(), RecoveryPending = pending,
                        RecoveryObservationError = pendingError, RefusalObservationError = refusalObservationError,
                        BackupCommitBefore = backupBefore, AfterImportCommitBefore = afterImportBefore,
                        BackupCommitAfter = repository.Resolve(MacroGitRepository.Backup),
                        AfterImportCommitAfter = repository.Resolve(MacroGitRepository.AfterImport),
                        FinishedUtc = DateTime.UtcNow.ToString("o") }); }
                    catch (Exception receiptError) { quarantined = true; throw new AggregateException(primary, receiptError); }
                    throw;
                }
            }
        }

        /// <summary>Loads the GUID-named manifest once, pins its hash, and verifies it belongs to the connected host.</summary>
        private void LoadOnce()
        {
            if (manifest != null) return;
            string name = Path.GetFileName(manifestPath);
            Guid id;
            if (!name.EndsWith(".owner-git.json", StringComparison.Ordinal) ||
                !Guid.TryParseExact(name.Substring(0, name.Length - ".owner-git.json".Length), "N", out id))
                throw new InvalidOperationException("GUID-named owner Git manifest required.");
            OwnerGitQualificationManifest.RequireClassicFilePath(manifestPath);
            OwnerGitQualificationManifest.RequireNoReparse(manifestPath);
            byte[] bytes = File.ReadAllBytes(manifestPath);
            if (bytes.Length == 0 || bytes.Length > 32768) throw new InvalidOperationException("Bounded manifest required.");
            var parsed = OwnerGitQualificationManifest.Parse(new System.Text.UTF8Encoding(false, true).GetString(bytes));
            OwnerGitQualificationManifest.RequireChild(parsed.EvidenceRoot, manifestPath, false);
            if (parsed.OwnerPid != connectedPid) throw new InvalidOperationException("Connected host differs from frozen owner.");
            manifestHash = OwnerGitQualificationManifest.Hash(bytes);
            manifest = parsed;
        }

        /// <summary>Rechecks the manifest hash, original process and VBE STA identity, active project, scope, and loaded assembly candidate.</summary>
        private void RequireContext()
        {
            if (manifest == null || OwnerGitQualificationManifest.HashFile(manifestPath) != manifestHash)
                throw new InvalidOperationException("Pinned owner Git manifest changed.");
            int pid; long birth;
            using (var process = Process.GetCurrentProcess()) { pid = process.Id; birth = process.StartTime.ToUniversalTime().Ticks; }
            uint windowPid;
            uint windowTid = OwnerGitQualificationManifest.GetWindowThreadProcessId(new IntPtr(manifest.VbeHandle), out windowPid);
            OwnerGitQualificationManifest.RequireOwnerIdentity(manifest, pid, birth, windowPid, windowTid,
                OwnerGitQualificationManifest.GetCurrentThreadId(), session.GitOwnerHandle(manifest.Project),
                session.GitActiveProjectMatches(manifest.Project, manifest.WorkbookPath) &&
                    string.Equals(session.GitScope(manifest.Project), manifest.WorkbookPath, StringComparison.OrdinalIgnoreCase),
                Thread.CurrentThread.GetApartmentState());
            if (typeof(VbeSession).Module.ModuleVersionId.ToString("D") != manifest.AssemblyMvid ||
                !string.Equals(OwnerGitQualificationManifest.HashFile(typeof(VbeSession).Assembly.Location),
                    manifest.AssemblySha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Installed product candidate changed.");
            OwnerGitQualificationManifest.RequirePolicy(LlmSettings.Load().VbeEditApproval);
            OwnerGitQualificationManifest.RequireNoReparse(manifest.FixtureRoot);
            OwnerGitQualificationManifest.RequireNoReparse(manifest.EvidenceRoot);
            OwnerGitQualificationManifest.RequireNoReparse(manifest.WorkbookPath);
            if (!File.Exists(manifest.WorkbookPath) || !Directory.Exists(manifest.EvidenceRoot))
                throw new InvalidOperationException("Disposable workbook or evidence root disappeared.");
        }

        /// <summary>Creates a new durable receipt without replacing an existing intent or terminal record.</summary>
        /// <param name="path">Destination path for a receipt under the frozen evidence root.</param>
        /// <param name="value">Receipt object serialized as UTF-8 JSON before the file is created.</param>
        private static void WriteNew(string path, object value)
        {
            OwnerGitQualificationManifest.RequireClassicFilePath(path);
            byte[] bytes = new System.Text.UTF8Encoding(false).GetBytes(new JavaScriptSerializer().Serialize(value) + "\n");
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { file.Write(bytes, 0, bytes.Length); file.Flush(true); }
        }

        /// <summary>Proves a declared malformed-checkpoint refusal occurred before mutation by checking recovery state, Git refs, and live project state.</summary>
        /// <param name="expectedError">Exact error substring declared for the frozen malformed-checkpoint scenario.</param>
        /// <param name="actualError">Error message returned by the attempted preflight.</param>
        /// <param name="mutationStarted">Whether the one-shot native mutation boundary was entered.</param>
        /// <param name="recoveryBefore">Recovery marker observed before preflight.</param>
        /// <param name="recoveryAfter">Recovery marker observed after refusal.</param>
        /// <param name="backupBefore">Backup commit resolved before preflight.</param>
        /// <param name="backupAfter">Backup commit resolved after refusal.</param>
        /// <param name="afterImportBefore">After-import commit resolved before preflight.</param>
        /// <param name="afterImportAfter">After-import commit resolved after refusal.</param>
        /// <param name="projectUnchanged">Whether a fresh live project capture still matches the preflight snapshot.</param>
        /// <returns>True only when the expected error is present, no mutation began, and all recovery, ref, and project observations are unchanged.</returns>
        internal static bool IsProvenPrewriteRefusal(string expectedError, string actualError, bool mutationStarted,
            bool recoveryBefore, bool recoveryAfter, string backupBefore, string backupAfter,
            string afterImportBefore, string afterImportAfter, bool projectUnchanged)
        {
            return expectedError != null && actualError != null && actualError.Contains(expectedError) && !mutationStarted &&
                !recoveryBefore && !recoveryAfter && projectUnchanged &&
                string.Equals(backupBefore, backupAfter, StringComparison.Ordinal) &&
                string.Equals(afterImportBefore, afterImportAfter, StringComparison.Ordinal);
        }
    }

    /// <summary>Exposes narrowly scoped owner-window and active-project checks for the frozen Git qualification runner.</summary>
    internal sealed partial class VbeSession
    {

        /// <summary>Resolves the selected project's VBE MainWindow handle while releasing every acquired COM reference.</summary>
        /// <param name="project">Exact project selector already authorized by this session.</param>
        /// <returns>Native VBE main-window handle as a 64-bit integer.</returns>
        internal long GitOwnerHandle(string project)
        {
            object selected = null, editor = null, main = null;
            Exception primary = null;
            try
            {
                selected = (object)GetProject(project);
                editor = (object)((dynamic)selected).VBE;
                main = (object)((dynamic)editor).MainWindow;
                return Convert.ToInt64(((dynamic)main).HWnd);
            }
            catch (Exception error) { primary = error; throw; }
            finally { FormFontRestoration.ReleaseOwnedReferences(new[] { selected, editor, main }, ReleaseOwnerReference, primary); }
        }

        /// <summary>Checks that the authorized project remains the active VBProject and its live path still matches the frozen path.</summary>
        /// <param name="project">Exact project selector previously authorized by this session.</param>
        /// <param name="path">Canonical project path recorded by the frozen Git step.</param>
        /// <returns><see langword="true"/> when object identity and case-insensitive path comparison both match.</returns>
        internal bool GitActiveProjectMatches(string project, string path)
        {
            object selected = null, active = null;
            Exception primary = null;
            try
            {
                selected = (object)GetProject(project);
                active = (object)((dynamic)vbe).ActiveVBProject;
                return VbeProjectHostPath.SameProject(selected, active) &&
                    string.Equals(VbeProjectHostPath.Read(active), path, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception error) { primary = error; throw; }
            finally { FormFontRestoration.ReleaseOwnedReferences(new[] { selected, active }, ReleaseOwnerReference, primary); }
        }

        /// <summary>Releases an acquired COM reference after owner-window or active-project inspection.</summary>
        /// <param name="value">COM object to release; non-COM and null values are ignored.</param>
        private static void ReleaseOwnerReference(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
