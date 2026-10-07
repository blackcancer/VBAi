using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using VBAi;

namespace VBAi.Tests.Unit
{
    /// <summary>Test-only, explicitly armed observation of actual revision operands and command results.</summary>
    internal sealed class GitRevisionFixtureTrace : IDisposable
    {
        internal const string EnvironmentVariable = "VBAi_GIT_REVISION_DIAGNOSTIC_ROOT";
        internal const int MaximumEvents = 512;
        internal const int MaximumBytes = 512 * 1024;
        private readonly object gate = new object();
        private readonly FileStream file;
        private readonly MacroGitRepository repository;
        private readonly Func<string[], byte[], bool, bool, MacroGitRepository.Result> previous;
        private int events, operation;
        private bool unavailable, disposed;

        private GitRevisionFixtureTrace(FileStream file, MacroGitRepository repository)
        {
            this.file = file; this.repository = repository; previous = repository.CommandOverride;
            repository.CommandOverride = ForwardCommand;
            Record(new { Kind = "Armed", ProductMvid = typeof(MacroGitOperations).Module.ModuleVersionId,
                TestMvid = typeof(GitRevisionFixtureTrace).Module.ModuleVersionId,
                MaximumEvents, MaximumBytes,
                Limitation = "Opt-in diagnostic hashing and synchronous journal writes add scheduling overhead." });
        }

        internal static GitRevisionFixtureTrace TryBegin(MacroGitRepository repository)
        {
            try
            {
                string root = Environment.GetEnvironmentVariable(EnvironmentVariable);
                return string.IsNullOrEmpty(root) ? null : TryBeginAtRoot(root, repository);
            }
            catch { return null; }
        }

        internal static GitRevisionFixtureTrace TryBeginAtRoot(string root, MacroGitRepository repository)
        {
            FileStream file = null;
            try
            {
                if (repository.CommandOverride != null) return null;
                RequireOwnedRoot(root);
                file = new FileStream(Path.Combine(root, "diagnostic.jsonl"), FileMode.CreateNew,
                    FileAccess.Write, FileShare.Read);
                return new GitRevisionFixtureTrace(file, repository);
            }
            catch { try { file?.Dispose(); } catch { } return null; }
        }

        private static void RequireOwnedRoot(string root)
        {
            string full = Path.GetFullPath(root);
            string temp = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
            Guid id;
            string name = Path.GetFileName(full);
            if (!Path.IsPathRooted(root) || !string.Equals(root, full, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetDirectoryName(full), temp, StringComparison.OrdinalIgnoreCase) ||
                !name.StartsWith("VBAiGitRevision-", StringComparison.Ordinal) ||
                !Guid.TryParseExact(name.Substring("VBAiGitRevision-".Length), "N", out id) || !Directory.Exists(full))
                throw new IOException("A precreated direct TEMP VBAiGitRevision-<guid> directory is required.");
            var drive = new DriveInfo(Path.GetPathRoot(full));
            if (drive.DriveType != DriveType.Fixed) throw new IOException("A fixed local drive is required.");
            for (string ancestor = full; ancestor != null; ancestor = Path.GetDirectoryName(ancestor))
                if ((File.GetAttributes(ancestor) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Diagnostic roots cannot traverse filesystem links.");
        }

        internal void Observe(MacroGitOperations operations)
        {
            int owner = Interlocked.Increment(ref operation);
            operations.ObserveRevision = value => Record(new { Kind = "Revision", Operation = owner, Data = value });
        }

        private MacroGitRepository.Result ForwardCommand(string[] args, byte[] input, bool use, bool allow)
        {
            try
            {
                // Native clears the override only for this one existing Run, then restores it in finally.
                var result = MacroGitRepositoryTests.Native(repository, args, input, use, allow);
                RecordCommand(args, input, use, allow, result, null);
                return result;
            }
            catch (Exception failure)
            {
                RecordCommand(args, input, use, allow, null, failure.GetType().FullName);
                throw;
            }
        }

        private void RecordCommand(string[] args, byte[] input, bool use, bool allow,
            MacroGitRepository.Result result, string failureType)
        {
            try
            {
                string command = args.Length > 0 && args[0].Length <= 32 && args[0].All(c => c >= 'a' && c <= 'z' || c == '-')
                    ? args[0] : "Other";
                Record(new { Kind = "Command", Operation = Volatile.Read(ref operation), Command = command,
                    ArgumentsSha256 = Sha(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(args))),
                    InputBytes = input?.Length, InputSha256 = input == null ? null : Sha(input),
                    UseRepository = use, AllowFailure = allow, ExitCode = result == null ? (int?)null : result.ExitCode,
                    ResultBytes = result?.Bytes?.Length,
                    ResultSha256 = result?.Bytes == null ? null : Sha(result.Bytes), FailureType = failureType });
            }
            catch { /* Preserve the result or original command exception even if diagnostics fail. */ }
        }

        internal void Record(object value)
        {
            lock (gate)
            {
                if (disposed || unavailable || events >= MaximumEvents) return;
                try
                {
                    byte[] line = new UTF8Encoding(false, true).GetBytes(new JavaScriptSerializer().Serialize(new
                    { Sequence = events + 1, Utc = DateTime.UtcNow.ToString("O"), Thread = Thread.CurrentThread.ManagedThreadId,
                        Event = value }) + "\n");
                    if (line.Length > MaximumBytes - file.Length) { unavailable = true; return; }
                    file.Write(line, 0, line.Length); file.Flush(true); events++;
                }
                catch { unavailable = true; }
            }
        }

        private static string Sha(byte[] bytes)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }

        public void Dispose()
        {
            lock (gate)
            {
                if (disposed) return;
                Record(new { Kind = "Closed", Operation = operation });
                disposed = true; repository.CommandOverride = previous;
                try { file.Dispose(); } catch { /* Cleanup must not replace the original test failure. */ }
            }
        }
    }
}