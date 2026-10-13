using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>One file-specific consent for a reviewed synthetic coverage copy; never changes Office trust policy.</summary>
    internal sealed class OfficeCoverageCopyConsent
    {
        /// <summary>Fresh, read-only native dialog evidence and its addressed activation boundary.</summary>
        internal sealed class Dialog
        {
            internal long Handle;
            internal int ProcessId;
            internal string Title, ClassName, RuntimeIdentity, ButtonIdentity;
            internal string[] Paths;
            internal int EnableButtons;
            internal bool Enabled;
            internal Action Invoke;
        }

        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private readonly int processId;
        private readonly string title, original, originalHash, coverageRoot, copyName;
        private readonly HashSet<string> priorFolders;
        private readonly Action verifyOwner;
        private readonly Func<Dialog[]> inspect;
        private readonly Action<string, object> record;
        internal bool Attempted { get; private set; }
        internal bool Returned { get; private set; }
        internal string ConsentedPath { get; private set; }

        /// <summary>Captures the saved original and preexisting copy folders before the sole coverage dispatch.</summary>
        internal OfficeCoverageCopyConsent(string host, int processId, string original, string coverageRoot,
            IEnumerable<string> priorFolders, Action verifyOwner, Func<Dialog[]> inspect, Action<string, object> record)
        {
            if (host != "Word" && host != "PowerPoint") throw new ArgumentException("Word or PowerPoint is required.");
            if (processId <= 0) throw new ArgumentException("An owned process is required.");
            this.processId = processId;
            this.original = Path.GetFullPath(original);
            this.coverageRoot = Path.GetFullPath(coverageRoot);
            this.priorFolders = new HashSet<string>(priorFolders.Select(Path.GetFullPath), StringComparer.OrdinalIgnoreCase);
            this.verifyOwner = verifyOwner ?? throw new ArgumentNullException(nameof(verifyOwner));
            this.inspect = inspect ?? throw new ArgumentNullException(nameof(inspect));
            this.record = record ?? throw new ArgumentNullException(nameof(record));
            title = "Avis de sécurité pour Microsoft " + host;
            copyName = "coverage" + (host == "PowerPoint" ? ".pptm" : Path.GetExtension(this.original));
            RequireOwner();
            RequireOrdinaryPath(this.original);
            originalHash = Hash(this.original);
        }

        /// <summary>Observes, revalidates and invokes at most one exact copy warning on the qualification STA.</summary>
        internal void Poll()
        {
            RequireOwner();
            if (Attempted) return;
            var windows = inspect();
            if (windows == null) throw new InvalidOperationException("Security dialog enumeration is unverified.");
            if (windows.Length == 0) return;
            if (windows.Length != 1) throw new InvalidOperationException("Exactly one owned copy warning is required.");
            var first = windows[0];
            string path = Validate(first);
            var fresh = inspect();
            if (fresh == null || fresh.Length != 1 || !SameDialog(first, fresh[0]) || Validate(fresh[0]) != path)
                throw new InvalidOperationException("The addressed copy warning changed before consent.");
            // Latch before durable intent. A lost receipt or an uncertain UI invocation cannot cause a retry.
            Attempted = true;
            ConsentedPath = path;
            record("coverage-copy-consent-intent", new { ProcessId = processId, Dialog = first.Handle,
                Copy = path, Original = original, OriginalSha256 = originalHash,
                FileSpecific = true, TrustPolicyChanged = false, InvocationRetried = false });
            var final = inspect();
            if (final == null || final.Length != 1 || !SameDialog(first, final[0]) || Validate(final[0]) != path)
                throw new InvalidOperationException("The addressed copy warning changed after consent intent; no invocation was attempted.");
            final[0].Invoke();
            Returned = true;
            record("coverage-copy-consent-returned", new { ProcessId = processId, Copy = path, Returned });
        }

        private string Validate(Dialog dialog)
        {
            RequireOwner();
            if (dialog == null || dialog.Handle == 0 || dialog.ProcessId != processId || dialog.Title != title
                || (dialog.ClassName != "NUIDialog" && dialog.ClassName != "#32770")
                || string.IsNullOrEmpty(dialog.RuntimeIdentity) || string.IsNullOrEmpty(dialog.ButtonIdentity)
                || dialog.Paths == null || dialog.Paths.Length != 1 || dialog.EnableButtons != 1
                || !dialog.Enabled || dialog.Invoke == null)
                throw new InvalidOperationException("The native file warning identity or unique enabled action is unverified.");
            string path = dialog.Paths[0];
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
                throw new InvalidOperationException("The warning must identify an absolute copy path.");
            string normalized = Path.GetFullPath(path);
            string folder = Path.GetDirectoryName(normalized);
            Guid id;
            if (!string.Equals(path, normalized, StringComparison.Ordinal) || !string.Equals(Path.GetFileName(path), copyName, StringComparison.Ordinal)
                || !string.Equals(Path.GetDirectoryName(folder), coverageRoot, StringComparison.OrdinalIgnoreCase)
                || !Guid.TryParseExact(Path.GetFileName(folder), "N", out id) || priorFolders.Contains(folder)
                || string.Equals(path, original, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only a newly created coverage copy in the frozen root may receive consent.");
            RequireOrdinaryPath(normalized);
            var freshFolders = Directory.GetDirectories(coverageRoot)
                .Where(candidate => !priorFolders.Contains(candidate)).ToArray();
            if (freshFolders.Length != 1 || !string.Equals(freshFolders[0], folder, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The coverage dispatch must own a single newly created copy folder.");
            if (Hash(original) != originalHash) throw new InvalidOperationException("The saved synthetic original changed before consent.");
            return normalized;
        }

        private void RequireOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Copy consent must remain on the original qualification thread.");
            verifyOwner();
        }

        private static bool SameDialog(Dialog first, Dialog second) => second != null && first.Handle == second.Handle
            && first.RuntimeIdentity == second.RuntimeIdentity && first.ButtonIdentity == second.ButtonIdentity;

        /// <summary>Rejects redirected files and ancestor directories without modifying their attributes.</summary>
        private static void RequireOrdinaryPath(string path)
        {
            if (!File.Exists(path)) throw new InvalidOperationException("The owned file is unavailable.");
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidOperationException("A reparse point cannot receive synthetic copy consent.");
        }

        /// <summary>Reads only the explicitly owned file; it never unlocks, saves or unblocks it.</summary>
        private static string Hash(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(stream));
        }
    }
}
