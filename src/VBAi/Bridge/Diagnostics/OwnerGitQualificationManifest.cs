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
        internal const string EnvironmentName = "VBAi_TEST_OWNER_GIT_MANIFEST";
        internal const string CommandName = "diagnostic_userform_git";
        // net48 FileStream uses the classic unprefixed Windows path contract here.
        // Reserve 80 characters below MAX_PATH for receipts (57), Git objects/refs
        // (at least 60), and fixed snapshot directories plus VBA component names.
        internal const int MaxEvidenceRootLength = 179;
        internal const int MaxClassicFilePathLength = 259;
        internal const int MaxClassicDirectoryPathLength = 247;
        public int Version { get; set; }
        public int OwnerPid { get; set; }
        public long OwnerBirthUtcTicks { get; set; }
        public uint OwnerNativeTid { get; set; }
        public long VbeHandle { get; set; }
        public string AssemblyMvid { get; set; }
        public string AssemblySha256 { get; set; }
        public string FixtureRoot { get; set; }
        public string WorkbookPath { get; set; }
        public string EvidenceRoot { get; set; }
        public string RepoRelativePath { get; set; }
        public string Project { get; set; }
        public string Branch { get; set; }
        public string RemoteUrl { get; set; }
        public string RemoteCommit { get; set; }
        public OwnerGitQualificationStep[] Steps { get; set; }

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

        private static void RequireKeys(IDictionary<string, object> value, string[] keys)
        {
            if (value == null || value.Count != keys.Length || value.Keys.Except(keys, StringComparer.Ordinal).Any() ||
                keys.Any(key => !value.ContainsKey(key))) throw new ArgumentException("Exact owner Git manifest fields required.");
        }
        private static bool Number(object value) => value is int || value is long;

        internal static void RequireExactRequest(string json)
        {
            var value = new JavaScriptSerializer().DeserializeObject(json) as IDictionary<string, object>;
            RequireKeys(value, new[] { "Command", "Action", "ExpectedSha256" });
            Guid id;
            if (!string.Equals(value["Command"] as string, CommandName, StringComparison.Ordinal) ||
                !Guid.TryParseExact(value["Action"] as string, "N", out id) || !Sha(value["ExpectedSha256"] as string))
                throw new ArgumentException("Exact diagnostic command, step id and coordinator revision required.");
        }

        internal static OwnerGitQualificationStep RequireStep(OwnerGitQualificationManifest manifest, int next,
            bool quarantined, string stepId)
        {
            if (manifest == null || quarantined || next < 0 || next >= manifest.Steps.Length ||
                !string.Equals(manifest.Steps[next].Id, stepId, StringComparison.Ordinal))
                throw new InvalidOperationException("Owner Git step was repeated, reordered, or quarantined.");
            return manifest.Steps[next];
        }

        internal static void RequireOwnerIdentity(OwnerGitQualificationManifest plan, int actualPid, long actualBirth,
            uint windowPid, uint windowTid, uint currentTid, long currentHandle, bool activeProject, System.Threading.ApartmentState apartment)
        {
            if (plan == null || plan.OwnerPid != actualPid || plan.OwnerBirthUtcTicks != actualBirth ||
                windowPid != actualPid || windowTid != plan.OwnerNativeTid || currentTid != windowTid ||
                currentHandle != plan.VbeHandle || !activeProject || apartment != System.Threading.ApartmentState.STA)
                throw new InvalidOperationException("Exact active VBE owner/project changed.");
        }

        internal static void RequirePolicy(string policy)
        {
            if (policy != "Automatic")
                throw new InvalidOperationException("Qualification mutation requires current Automatic VBE edit policy.");
        }

        internal static bool Sha(string value) => value != null && value.Length == 64 && value.All(c =>
            c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F');
        private static bool Commit(string value) => value != null && value.Length == 40 && value.All(c =>
            c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F');

        internal static void RequireGuidRoot(string root)
        {
            RequireClassicDirectoryPath(root);
            if (!Guid.TryParseExact(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar)), "N", out _))
                throw new ArgumentException("Canonical local GUID fixture root required.");
            RequireNoReparse(root);
        }

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

        private static void RequireCanonicalLocalPath(string path)
        {
            // Reject NTFS streams before GetFullPath, which throws NotSupportedException
            // for their syntax on net48. All manifest path fields share this contract.
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                path.StartsWith(@"\\", StringComparison.Ordinal) || path.IndexOf(':', 2) >= 0 ||
                Path.GetFullPath(path) != path)
                throw new ArgumentException("Canonical local path without alternate streams required.");
        }

        internal static void RequireEvidenceRootBudget(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || root.Length > MaxEvidenceRootLength)
                throw new ArgumentException("Disposable owner Git evidence root exceeds the net48 path budget.");
            RequireClassicDirectoryPath(root);
        }

        internal static void RequireReceiptPaths(string evidenceRoot, string stepId)
        {
            RequireEvidenceRootBudget(evidenceRoot);
            Guid id;
            if (!Guid.TryParseExact(stepId, "N", out id)) throw new ArgumentException("Exact owner Git receipt id required.");
            string prefix = Path.Combine(evidenceRoot, "owner-git-" + stepId);
            foreach (string suffix in new[] { ".intent.json", ".mutation.json", ".terminal.json" })
                RequireClassicFilePath(prefix + suffix);
        }

        internal static void RequireClassicFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length > MaxClassicFilePathLength)
                throw new ArgumentException("Owner Git file path exceeds the net48 path budget.");
            RequireCanonicalLocalPath(path);
            RequireClassicDirectoryPath(Path.GetDirectoryName(path));
        }

        internal static void RequireClassicDirectoryPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length > MaxClassicDirectoryPathLength)
                throw new ArgumentException("Owner Git directory path exceeds the net48 path budget.");
            RequireCanonicalLocalPath(path);
        }

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

        internal static string Hash(byte[] data)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
        }
        internal static string HashFile(string path) => Hash(File.ReadAllBytes(path));

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

        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    }

    internal sealed class OwnerGitQualificationStep
    {
        public string Id { get; set; }
        public string Verb { get; set; }
        public string ExpectedSnapshotDirectory { get; set; }
        public string ExpectedSnapshotSha256 { get; set; }
        public string TargetSnapshotDirectory { get; set; }
        public string TargetSnapshotSha256 { get; set; }
        public string CheckpointId { get; set; }
        public string ExpectedErrorSubstring { get; set; }
    }
}
