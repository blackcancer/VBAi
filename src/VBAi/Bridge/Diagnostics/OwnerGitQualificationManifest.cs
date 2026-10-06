using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>A single disposable native qualification, never a general Git bridge.</summary>
    internal sealed class OwnerGitQualificationManifest
    {

        /// <summary>Environment variable captured when the add-in connects to arm the owned fixture.</summary>
        internal const string EnvironmentName = "VBAi_TEST_OWNER_GIT_MANIFEST";

        /// <summary>Only bridge command accepted by this diagnostic plan.</summary>
        internal const string CommandName = "diagnostic_userform_git";
        // net48 FileStream uses the classic unprefixed Windows path contract here.
        // Reserve 80 characters below MAX_PATH for receipts (57), Git objects/refs
        // (at least 60), and fixed snapshot directories plus VBA component names.
        /// <summary>Maximum evidence-root path length that leaves room for every bounded receipt and Git object path.</summary>
        internal const int MaxEvidenceRootLength = 179;

        /// <summary>Maximum path length accepted by the classic net48 file API used by the fixture.</summary>
        internal const int MaxClassicFilePathLength = 259;

        /// <summary>Maximum repository-directory path length accepted by the classic net48 directory API.</summary>
        internal const int MaxClassicDirectoryPathLength = 247;

        /// <summary>Gets or sets the serialized manifest schema version.</summary>
        /// <value>Must be <c>1</c>; other schema versions are refused.</value>
        public int Version { get; set; }

        /// <summary>Gets or sets the process ID of the host that owns the armed fixture.</summary>
        /// <value>Positive PID captured from the original connected host, checked with its creation time to detect PID reuse.</value>
        public int OwnerPid { get; set; }

        /// <summary>Gets or sets the UTC process-start ticks for the original host process.</summary>
        /// <value>Positive UTC ticks used with <see cref="OwnerPid"/> to prove the same process still owns the fixture.</value>
        public long OwnerBirthUtcTicks { get; set; }

        /// <summary>Gets or sets the native thread ID of the host's owning VBE STA.</summary>
        /// <value>Nonzero TID captured during setup; native work must still execute on this thread.</value>
        public uint OwnerNativeTid { get; set; }

        /// <summary>Gets or sets the original VBE root window handle.</summary>
        /// <value>Nonzero native handle whose process, thread, and root ancestry are revalidated before each step.</value>
        public long VbeHandle { get; set; }

        /// <summary>Gets or sets the candidate add-in module version ID frozen for this qualification.</summary>
        /// <value>Canonical dashed GUID of the exact assembly admitted by the fixture.</value>
        public string AssemblyMvid { get; set; }

        /// <summary>Gets or sets the SHA-256 digest of the candidate add-in assembly.</summary>
        /// <value>Lowercase or uppercase hexadecimal digest checked against the frozen candidate.</value>
        public string AssemblySha256 { get; set; }

        /// <summary>Gets or sets the GUID-named root of the owned disposable workbook fixture.</summary>
        /// <value>Absolute fixture directory that must contain <see cref="WorkbookPath"/>.</value>
        public string FixtureRoot { get; set; }

        /// <summary>Gets or sets the absolute path of the disposable workbook bound to this plan.</summary>
        /// <value>Classic-length child path under <see cref="FixtureRoot"/> matching the authorized project identity.</value>
        public string WorkbookPath { get; set; }

        /// <summary>Gets or sets the isolated directory for repository data, immutable snapshots, and operation receipts.</summary>
        /// <value>GUID-suffixed absolute directory validated for path bounds and reparse points before use.</value>
        public string EvidenceRoot { get; set; }

        /// <summary>Gets or sets the single repository directory name relative to the evidence root.</summary>
        /// <value>One nonrooted path component; separators, dot segments, and paths outside the evidence root are refused.</value>
        public string RepoRelativePath { get; set; }

        /// <summary>Gets or sets the canonical VBA project identity authorized for Git operations.</summary>
        /// <value>Must match the disposable workbook path under the plan's case-insensitive identity check.</value>
        public string Project { get; set; }

        /// <summary>Gets or sets the exact local branch expected by the fixture.</summary>
        /// <value>Nonempty branch name bounded to 128 characters and rechecked before each operation.</value>
        public string Branch { get; set; }

        /// <summary>Gets or sets the optional qualification remote URL.</summary>
        /// <value>Null for local banks; otherwise must equal the single frozen qualification repository URL.</value>
        public string RemoteUrl { get; set; }

        /// <summary>Gets or sets the exact incoming commit for the optional remote pull step.</summary>
        /// <value>Required for a pull plan and absent from local plans.</value>
        public string RemoteCommit { get; set; }

        /// <summary>Gets or sets the fixed ordered steps allowed for this fixture.</summary>
        /// <value>One to four validated steps from the local, malformed-checkpoint, or frozen remote bank.</value>
        public OwnerGitQualificationStep[] Steps { get; set; }

        /// <summary>Parses and validates the complete frozen owner-Git plan before any repository operation.</summary>
        /// <param name="json">Serialized manifest, limited to 32 KiB and required to contain exactly the known fields.</param>
        /// <returns>A validated plan bound to one host process, VBE STA, candidate assembly, workbook, evidence root, and ordered step bank.</returns>
        internal static OwnerGitQualificationManifest Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || Encoding.UTF8.GetByteCount(json) > 32768)
                throw new ArgumentException("Bounded owner Git manifest required.");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 32768 };
            var raw = serializer.DeserializeObject(json) as IDictionary<string, object>;
            RequireKeys(raw, new[] { "Version", "OwnerPid", "OwnerBirthUtcTicks", "OwnerNativeTid", "VbeHandle",
                "AssemblyMvid", "AssemblySha256", "FixtureRoot", "WorkbookPath", "EvidenceRoot", "RepoRelativePath",
                "Project", "Branch", "RemoteUrl", "RemoteCommit", "Steps" });
            if (!(raw["Version"] is int) || !(raw["OwnerPid"] is int) || !Number(raw["OwnerBirthUtcTicks"]) ||
                !Number(raw["OwnerNativeTid"]) || !Number(raw["VbeHandle"]) ||
                new[] { "AssemblyMvid", "AssemblySha256", "FixtureRoot", "WorkbookPath", "EvidenceRoot",
                    "RepoRelativePath", "Project", "Branch" }.Any(key => !(raw[key] is string)) ||
                raw["RemoteUrl"] != null && !(raw["RemoteUrl"] is string) ||
                raw["RemoteCommit"] != null && !(raw["RemoteCommit"] is string))
                throw new ArgumentException("Owner Git manifest has noncanonical field types.");
            var steps = raw["Steps"] as object[];
            if (steps == null || steps.Length < 1 || steps.Length > 4)
                throw new ArgumentException("One to four fixed owner Git steps are required.");
            foreach (var item in steps)
            {
                RequireKeys(item as IDictionary<string, object>, new[] { "Id", "Verb", "ExpectedSnapshotDirectory",
                    "ExpectedSnapshotSha256", "TargetSnapshotDirectory", "TargetSnapshotSha256", "CheckpointId", "ExpectedErrorSubstring" });
                var fields = (IDictionary<string, object>)item;
                if (new[] { "Id", "Verb", "ExpectedSnapshotDirectory", "ExpectedSnapshotSha256",
                        "TargetSnapshotDirectory", "TargetSnapshotSha256" }.Any(key => !(fields[key] is string)) ||
                    fields["CheckpointId"] != null && !(fields["CheckpointId"] is string) ||
                    fields["ExpectedErrorSubstring"] != null && !(fields["ExpectedErrorSubstring"] is string))
                    throw new ArgumentException("Owner Git step has noncanonical field types.");
            }
            var manifest = serializer.Deserialize<OwnerGitQualificationManifest>(json);
            if (manifest.Version != 1 || manifest.OwnerPid <= 0 || manifest.OwnerBirthUtcTicks <= 0 ||
                manifest.OwnerNativeTid == 0 || manifest.VbeHandle == 0 || manifest.Steps == null || manifest.Steps.Length != steps.Length)
                throw new ArgumentException("Invalid owner Git identity or version.");
            Guid mvid;
            if (!Guid.TryParseExact(manifest.AssemblyMvid, "D", out mvid) || !Sha(manifest.AssemblySha256))
                throw new ArgumentException("Exact candidate identity required.");
            if (string.IsNullOrWhiteSpace(manifest.Project) || manifest.Project.Length > 1024 ||
                !string.Equals(manifest.Project, manifest.WorkbookPath, StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(manifest.Branch) || manifest.Branch.Length > 128)
                throw new ArgumentException("Exact project and branch required.");
            RequireGuidRoot(manifest.FixtureRoot);
            RequireChild(manifest.FixtureRoot, manifest.WorkbookPath, false);
            RequireClassicFilePath(manifest.WorkbookPath);
            RequireEvidenceRootBudget(manifest.EvidenceRoot);
            if (!Guid.TryParseExact(Path.GetFileName(manifest.EvidenceRoot).Split('-').Last(), "N", out _))
                throw new ArgumentException("GUID-suffixed evidence root required.");
            RequireNoReparse(manifest.EvidenceRoot);
            RequireEvidenceRootBudget(manifest.EvidenceRoot);
            if (string.IsNullOrWhiteSpace(manifest.RepoRelativePath) || Path.IsPathRooted(manifest.RepoRelativePath) ||
                manifest.RepoRelativePath.IndexOfAny(new[] { ':', '/', '\\' }) >= 0 ||
                manifest.RepoRelativePath == "." || manifest.RepoRelativePath == "..")
                throw new ArgumentException("Repository must be one named child of evidence root.");
            RequireChild(manifest.EvidenceRoot, Path.Combine(manifest.EvidenceRoot, manifest.RepoRelativePath), true);
            string repoPath = Path.Combine(manifest.EvidenceRoot, manifest.RepoRelativePath);
            RequireClassicDirectoryPath(repoPath);
            RequireClassicFilePath(Path.Combine(repoPath, "objects", "aa", new string('a', 38)));
            RequireClassicFilePath(Path.Combine(repoPath, "refs", "codex", "checkpoints", "20261006000000000-abcdef12"));
            RequireClassicFilePath(Path.Combine(repoPath, "refs", "heads", manifest.Branch));
            RequireClassicFilePath(Path.Combine(repoPath, "refs", "remotes", "origin", manifest.Branch));
            if (!string.IsNullOrEmpty(manifest.RemoteUrl))
            {
                if (manifest.RemoteUrl != "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git" ||
                    !manifest.Branch.StartsWith("qualification-userform-", StringComparison.Ordinal) || !Commit(manifest.RemoteCommit))
                    throw new ArgumentException("Exact authenticated remote and commit required.");
            }
            else if (!string.IsNullOrEmpty(manifest.RemoteCommit)) throw new ArgumentException("Remote commit without remote URL.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var step in manifest.Steps)
            {
                Guid id;
                if (!Guid.TryParseExact(step.Id, "N", out id) || !ids.Add(step.Id) ||
                    !new[] { "checkpoint_restore", "controlled_interruption", "rollback", "pull" }.Contains(step.Verb))
                    throw new ArgumentException("Unknown or repeated owner Git step.");
                RequireChild(manifest.EvidenceRoot, step.ExpectedSnapshotDirectory, true);
                RequireChild(manifest.EvidenceRoot, step.TargetSnapshotDirectory, true);
                RequireClassicDirectoryPath(step.ExpectedSnapshotDirectory);
                RequireClassicDirectoryPath(step.TargetSnapshotDirectory);
                foreach (string directory in new[] { step.ExpectedSnapshotDirectory, step.TargetSnapshotDirectory })
                    if (Directory.Exists(directory))
                        foreach (string file in Directory.GetFiles(directory)) RequireClassicFilePath(file);
                RequireReceiptPaths(manifest.EvidenceRoot, step.Id);
                if (!Sha(step.ExpectedSnapshotSha256) || !Sha(step.TargetSnapshotSha256))
                    throw new ArgumentException("Exact source and target snapshot hashes required.");
                if (step.Verb == "checkpoint_restore" &&
                    !System.Text.RegularExpressions.Regex.IsMatch(step.CheckpointId ?? "", @"^\d{17}-[0-9a-f]{8}$"))
                    throw new ArgumentException("An exact checkpoint id is required.");
                if (step.Verb != "checkpoint_restore" && step.CheckpointId != null)
                    throw new ArgumentException("Unexpected checkpoint id.");
                if (step.Verb == "pull" && string.IsNullOrEmpty(manifest.RemoteCommit))
                    throw new ArgumentException("Pull requires frozen remote commit.");
                if (step.ExpectedErrorSubstring != null &&
                    (step.Verb != "checkpoint_restore" || step.ExpectedErrorSubstring.Length < 4 || step.ExpectedErrorSubstring.Length > 256))
                    throw new ArgumentException("Only a declared malformed checkpoint can expect preflight refusal.");
            }
            string sequence = string.Join(",", manifest.Steps.Select(step => step.Verb));
            if (sequence != "checkpoint_restore,controlled_interruption,rollback" &&
                sequence != "checkpoint_restore" && sequence != "pull")
                throw new ArgumentException("Only the frozen local, corruption, or remote bank sequence is permitted.");
            if ((sequence == "pull") != (manifest.RemoteCommit != null) ||
                (sequence == "checkpoint_restore") != (manifest.Steps[0].ExpectedErrorSubstring != null))
                throw new ArgumentException("Bank and preflight outcome mismatch.");
            return manifest;
        }

        /// <summary>Requires an object to contain exactly the manifest fields named by the frozen schema.</summary>
        /// <param name="value">Deserialized JSON object to validate.</param>
        /// <param name="keys">Expected field names; missing or extra keys are refused.</param>
        private static void RequireKeys(IDictionary<string, object> value, string[] keys)
        {
            if (value == null || value.Count != keys.Length || value.Keys.Except(keys, StringComparer.Ordinal).Any() ||
                keys.Any(key => !value.ContainsKey(key))) throw new ArgumentException("Exact owner Git manifest fields required.");
        }

        /// <summary>Checks whether the JSON serializer returned an integer representation accepted for tick, TID, or handle fields.</summary>
        /// <param name="value">Deserialized JSON value to check.</param>
        /// <returns>True when the value is an <see cref="int"/> or <see cref="long"/>.</returns>
        private static bool Number(object value) => value is int || value is long;

        /// <summary>Accepts only the diagnostic command, one canonical step GUID, and a SHA-256 coordinator revision.</summary>
        /// <param name="json">Raw bridge request JSON; additional or missing fields are rejected.</param>
        internal static void RequireExactRequest(string json)
        {
            var value = new JavaScriptSerializer().DeserializeObject(json) as IDictionary<string, object>;
            RequireKeys(value, new[] { "Command", "Action", "ExpectedSha256" });
            Guid id;
            if (!string.Equals(value["Command"] as string, CommandName, StringComparison.Ordinal) ||
                !Guid.TryParseExact(value["Action"] as string, "N", out id) || !Sha(value["ExpectedSha256"] as string))
                throw new ArgumentException("Exact diagnostic command, step id and coordinator revision required.");
        }

        /// <summary>Returns only the next ordered step and refuses a repeated, reordered, exhausted, or quarantined plan.</summary>
        /// <param name="manifest">Validated frozen plan containing the allowed step sequence.</param>
        /// <param name="next">Zero-based index of the next step to execute.</param>
        /// <param name="quarantined">Whether a prior uncertain outcome has disabled further execution.</param>
        /// <param name="stepId">Action GUID supplied by the current bridge request.</param>
        /// <returns>The next step when its ID exactly matches <paramref name="stepId"/>.</returns>
        internal static OwnerGitQualificationStep RequireStep(OwnerGitQualificationManifest manifest, int next,
            bool quarantined, string stepId)
        {
            if (manifest == null || quarantined || next < 0 || next >= manifest.Steps.Length ||
                !string.Equals(manifest.Steps[next].Id, stepId, StringComparison.Ordinal))
                throw new InvalidOperationException("Owner Git step was repeated, reordered, or quarantined.");
            return manifest.Steps[next];
        }

        /// <summary>Requires the same host process instance, original VBE window, and owner STA captured when the disposable fixture was armed.</summary>
        /// <param name="plan">Frozen identity fields loaded from the pinned manifest.</param>
        /// <param name="actualPid">Current host process ID.</param>
        /// <param name="actualBirth">Current process start time in UTC ticks, preventing PID reuse from matching the original host.</param>
        /// <param name="windowPid">Process ID that owns the frozen VBE root handle.</param>
        /// <param name="windowTid">Native thread ID that owns the frozen VBE root handle.</param>
        /// <param name="currentTid">Native thread ID of the thread performing this validation.</param>
        /// <param name="currentHandle">Current VBE owner handle retained by the session.</param>
        /// <param name="activeProject">Whether the authorized workbook is the active project and remains the session scope.</param>
        /// <param name="apartment">Apartment state of the calling thread; native qualification requires STA.</param>
        internal static void RequireOwnerIdentity(OwnerGitQualificationManifest plan, int actualPid, long actualBirth,
            uint windowPid, uint windowTid, uint currentTid, long currentHandle, bool activeProject, System.Threading.ApartmentState apartment)
        {
            if (plan == null || plan.OwnerPid != actualPid || plan.OwnerBirthUtcTicks != actualBirth ||
                windowPid != actualPid || windowTid != plan.OwnerNativeTid || currentTid != windowTid ||
                currentHandle != plan.VbeHandle || !activeProject || apartment != System.Threading.ApartmentState.STA)
                throw new InvalidOperationException("Exact active VBE owner/project changed.");
        }

        /// <summary>Requires policy for owner git qualification manifest.</summary>
        /// <param name="policy">Text that supplies the policy value. Use the format required by the calling operation.</param>
        internal static void RequirePolicy(string policy)
        {
            if (policy != "Automatic")
                throw new InvalidOperationException("Qualification mutation requires current Automatic VBE edit policy.");
        }

        /// <summary>Handles sha for owner git qualification manifest.</summary>
        /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for sha on owner git qualification manifest.</returns>
        internal static bool Sha(string value) => value != null && value.Length == 64 && value.All(c =>
            c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F');

        /// <summary>Handles commit for owner git qualification manifest.</summary>
        /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for commit on owner git qualification manifest.</returns>
        private static bool Commit(string value) => value != null && value.Length == 40 && value.All(c =>
            c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F');

        /// <summary>Requires guid root for owner git qualification manifest.</summary>
        /// <param name="root">Text that supplies the root value. Use the format required by the calling operation.</param>
        internal static void RequireGuidRoot(string root)
        {
            RequireClassicDirectoryPath(root);
            if (!Guid.TryParseExact(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar)), "N", out _))
                throw new ArgumentException("Canonical local GUID fixture root required.");
            RequireNoReparse(root);
        }

        /// <summary>Requires child for owner git qualification manifest.</summary>
        /// <param name="parent">Text that supplies the parent value. Use the format required by the calling operation.</param>
        /// <param name="child">Text that supplies the child value. Use the format required by the calling operation.</param>
        /// <param name="directory">Indicates whether directory is enabled.</param>
        internal static void RequireChild(string parent, string child, bool directory)
        {
            if (directory) RequireClassicDirectoryPath(child);
            else RequireClassicFilePath(child);
            if (!child.StartsWith(parent.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar,
                    StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Path must be a canonical child of the frozen disposable root.");
            RequireNoReparse(child);
            if (directory && File.Exists(child) || !directory && Directory.Exists(child))
                throw new ArgumentException("Unexpected filesystem object type.");
        }

        /// <summary>Requires canonical local path for owner git qualification manifest.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        private static void RequireCanonicalLocalPath(string path)
        {
            // Reject NTFS streams before GetFullPath, which throws NotSupportedException
            // for their syntax on net48. All manifest path fields share this contract.
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                path.StartsWith(@"\\", StringComparison.Ordinal) || path.IndexOf(':', 2) >= 0 ||
                Path.GetFullPath(path) != path)
                throw new ArgumentException("Canonical local path without alternate streams required.");
        }

        /// <summary>Requires evidence root budget for owner git qualification manifest.</summary>
        /// <param name="root">Text that supplies the root value. Use the format required by the calling operation.</param>
        internal static void RequireEvidenceRootBudget(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || root.Length > MaxEvidenceRootLength)
                throw new ArgumentException("Disposable owner Git evidence root exceeds the net48 path budget.");
            RequireClassicDirectoryPath(root);
        }

        /// <summary>Requires receipt paths for owner git qualification manifest.</summary>
        /// <param name="evidenceRoot">Text that supplies the evidence root value. Use the format required by the calling operation.</param>
        /// <param name="stepId">Text that supplies the step id value. Use the format required by the calling operation.</param>
        internal static void RequireReceiptPaths(string evidenceRoot, string stepId)
        {
            RequireEvidenceRootBudget(evidenceRoot);
            Guid id;
            if (!Guid.TryParseExact(stepId, "N", out id)) throw new ArgumentException("Exact owner Git receipt id required.");
            string prefix = Path.Combine(evidenceRoot, "owner-git-" + stepId);
            foreach (string suffix in new[] { ".intent.json", ".mutation.json", ".terminal.json" })
                RequireClassicFilePath(prefix + suffix);
        }

        /// <summary>Requires classic file path for owner git qualification manifest.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        internal static void RequireClassicFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length > MaxClassicFilePathLength)
                throw new ArgumentException("Owner Git file path exceeds the net48 path budget.");
            RequireCanonicalLocalPath(path);
            RequireClassicDirectoryPath(Path.GetDirectoryName(path));
        }

        /// <summary>Requires classic directory path for owner git qualification manifest.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        internal static void RequireClassicDirectoryPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length > MaxClassicDirectoryPathLength)
                throw new ArgumentException("Owner Git directory path exceeds the net48 path budget.");
            RequireCanonicalLocalPath(path);
        }

        /// <summary>Requires no reparse for owner git qualification manifest.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        internal static void RequireNoReparse(string path)
        {
            string current = Path.GetPathRoot(path);
            foreach (string part in path.Substring(current.Length).Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries))
            {
                current = Path.Combine(current, part);
                if ((File.Exists(current) || Directory.Exists(current)) &&
                    (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new ArgumentException("Reparse point in owner Git path.");
            }
        }

        /// <summary>Determines whether it has h for owner git qualification manifest.</summary>
        /// <param name="data">byte[] that supplies the data for this operation.</param>
        /// <returns>Text produced by the operation for hash on owner git qualification manifest.</returns>
        internal static string Hash(byte[] data)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
        }

        /// <summary>Determines whether it has h file for owner git qualification manifest.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>Text produced by the operation for hash file on owner git qualification manifest.</returns>
        internal static string HashFile(string path) => Hash(File.ReadAllBytes(path));

        /// <summary>Handles snapshot hash for owner git qualification manifest.</summary>
        /// <param name="snapshot">vba git snapshot that supplies the snapshot for this operation.</param>
        /// <returns>Text produced by the operation for snapshot hash on owner git qualification manifest.</returns>
        internal static string SnapshotHash(VbaGitSnapshot snapshot)
        {
            using (var stream = new MemoryStream())
            {
                foreach (var file in snapshot.Serialize())
                {
                    byte[] name = Encoding.UTF8.GetBytes(file.Key + "\0" + file.Value.Length + "\0");
                    stream.Write(name, 0, name.Length); stream.Write(file.Value, 0, file.Value.Length);
                }
                return Hash(stream.ToArray());
            }
        }

        /// <summary>Reads snapshot for owner git qualification manifest.</summary>
        /// <param name="directory">Text that supplies the directory value. Use the format required by the calling operation.</param>
        /// <param name="expectedHash">Text that supplies the expected hash value. Use the format required by the calling operation.</param>
        /// <returns>vba git snapshot produced by the operation for read snapshot on owner git qualification manifest.</returns>
        internal static VbaGitSnapshot ReadSnapshot(string directory, string expectedHash)
        {
            RequireNoReparse(directory);
            var paths = Directory.GetFiles(directory);
            if (paths.Length < 2 || paths.Length > 2049 || Directory.GetDirectories(directory).Length != 0 ||
                paths.Sum(path => new FileInfo(path).Length) > VbaGitSnapshot.MaxBytes)
                throw new InvalidOperationException("Bounded snapshot files required.");
            foreach (var path in paths) RequireNoReparse(path);
            if (!string.Equals(HashFiles(paths), expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Frozen snapshot files changed.");
            var files = paths.ToDictionary(path => Path.GetFileName(path), File.ReadAllBytes, StringComparer.Ordinal);
            return VbaGitSnapshot.Read(files);
        }

        /// <summary>Handles snapshot directory hash for owner git qualification manifest.</summary>
        /// <param name="directory">Text that supplies the directory value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for snapshot directory hash on owner git qualification manifest.</returns>
        internal static string SnapshotDirectoryHash(string directory)
        {
            RequireNoReparse(directory);
            var paths = Directory.GetFiles(directory);
            if (paths.Length < 2 || paths.Length > 2049 || Directory.GetDirectories(directory).Length != 0 ||
                paths.Sum(path => new FileInfo(path).Length) > VbaGitSnapshot.MaxBytes)
                throw new InvalidOperationException("Bounded snapshot files required.");
            foreach (var path in paths) RequireNoReparse(path);
            var files = paths.ToDictionary(path => Path.GetFileName(path), File.ReadAllBytes, StringComparer.Ordinal);
            VbaGitSnapshot.Read(files);
            return HashFiles(paths);
        }

        /// <summary>Determines whether it has h files for owner git qualification manifest.</summary>
        /// <param name="paths">i enumerable&lt;string&gt; that supplies the paths for this operation.</param>
        /// <returns>Text produced by the operation for hash files on owner git qualification manifest.</returns>
        private static string HashFiles(IEnumerable<string> paths)
        {
            using (var stream = new MemoryStream())
            {
                foreach (string path in paths.OrderBy(path => Path.GetFileName(path), StringComparer.Ordinal))
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    byte[] name = Encoding.UTF8.GetBytes(Path.GetFileName(path) + "\0" + bytes.Length + "\0");
                    stream.Write(name, 0, name.Length); stream.Write(bytes, 0, bytes.Length);
                }
                return Hash(stream.ToArray());
            }
        }

        /// <summary>Returns current thread id for owner git qualification manifest.</summary>
        /// <returns>uint produced by the operation for get current thread id on owner git qualification manifest.</returns>
        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();

        /// <summary>Returns window thread process id for owner git qualification manifest.</summary>
        /// <param name="hwnd">Native handle that supplies the hwnd for this operation.</param>
        /// <param name="pid">uint that supplies the pid for this operation.</param>
        /// <returns>uint produced by the operation for get window thread process id on owner git qualification manifest.</returns>
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    }

    /// <summary>Owns the owner git qualification step state and operations.</summary>
    internal sealed class OwnerGitQualificationStep
    {

        /// <summary>Gets or sets the id.</summary>
        /// <value>Current id exposed by owner git qualification step.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the verb.</summary>
        /// <value>Current verb exposed by owner git qualification step.</value>
        public string Verb { get; set; }

        /// <summary>Gets or sets the expected snapshot directory.</summary>
        /// <value>Current expected snapshot directory exposed by owner git qualification step.</value>
        public string ExpectedSnapshotDirectory { get; set; }

        /// <summary>Gets or sets the expected snapshot sha256.</summary>
        /// <value>Current expected snapshot sha256 exposed by owner git qualification step.</value>
        public string ExpectedSnapshotSha256 { get; set; }

        /// <summary>Gets or sets the target snapshot directory.</summary>
        /// <value>Current target snapshot directory exposed by owner git qualification step.</value>
        public string TargetSnapshotDirectory { get; set; }

        /// <summary>Gets or sets the target snapshot sha256.</summary>
        /// <value>Current target snapshot sha256 exposed by owner git qualification step.</value>
        public string TargetSnapshotSha256 { get; set; }

        /// <summary>Gets or sets the checkpoint id.</summary>
        /// <value>Current checkpoint id exposed by owner git qualification step.</value>
        public string CheckpointId { get; set; }

        /// <summary>Gets or sets the expected error substring.</summary>
        /// <value>Current expected error substring exposed by owner git qualification step.</value>
        public string ExpectedErrorSubstring { get; set; }
    }
}
