using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns one editor-only profile; retirement never deletes another window's or a legacy profile.</summary>
    internal sealed class EditorBrowserProfile
    {

        /// <summary>Serializes browser lifecycle observations and cleanup scheduling.</summary>
        private readonly object gate = new object();

        /// <summary>Tracks whether the environment was requested, its controller retired, its browser exited, and cleanup was queued.</summary>
        private bool requested, retired, exited, scheduled;

        /// <summary>Process identifier reported by this profile's browser environment; zero means no process has been observed.</summary>
        private uint browserProcessId;

        /// <summary>The unique user data folder supplied only to this environment.</summary>
        /// <value>Current path exposed by editor browser profile.</value>
        internal string Path { get; }

        /// <summary>The asynchronous filesystem cleanup, scheduled only after retirement and proven runtime exit.</summary>
        /// <value>Current cleanup exposed by editor browser profile.</value>
        internal Task Cleanup { get; private set; } = Task.CompletedTask;

        /// <summary>Creates an operation-owned profile below the supplied editor cache root.</summary>
        /// <param name="root">Existing or new cache directory that will contain this operation's uniquely named profile. Reparse-point roots are rejected.</param>
        internal EditorBrowserProfile(string root)
        {
            root = System.IO.Path.GetFullPath(root);
            Directory.CreateDirectory(root);
            if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("The editor profile root must not be a reparse point.");
            Path = System.IO.Path.Combine(root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        /// <summary>Prevents cleanup until the environment reports exit, even if initialization later fails.</summary>
        internal void BrowserRequested() { lock (gate) requested = true; }

        /// <summary>Records the actual browser owning this environment, invalidating an older exit observation.</summary>
        /// <param name="processId">Nonzero process identifier reported by the browser controller.</param>
        internal void ObserveBrowser(uint processId)
        {
            if (processId == 0) throw new ArgumentOutOfRangeException(nameof(processId));
            lock (gate)
            {
                if (browserProcessId != processId) exited = false;
                browserProcessId = processId;
            }
        }

        /// <summary>Accepts only this environment's exit notification; a live editor retains its profile.</summary>
        /// <param name="processId">Process identifier from the browser-exited notification; a mismatched or zero identifier is ignored.</param>
        internal void BrowserExited(uint processId)
        {
            lock (gate)
            {
                if (processId == 0 || (browserProcessId != 0 && processId != browserProcessId)) return;
                browserProcessId = processId;
                exited = true;
                ScheduleCleanup();
            }
        }

        /// <summary>Called after controller disposal; missing exit evidence preserves the profile without blocking the UI.</summary>
        internal void Retire()
        {
            lock (gate) { retired = true; ScheduleCleanup(); }
        }

        /// <summary>Schedules filesystem work once; no browser or COM object crosses to the worker.</summary>
        private void ScheduleCleanup()
        {
            if (!retired || (requested && !exited) || scheduled) return;
            scheduled = true;
            Cleanup = Task.Run(() =>
            {
                try
                {
                    if (!Directory.Exists(Path)) return;
                    var files = new List<string>();
                    var directories = new List<string>();
                    CollectOwnedTree(Path, files, directories);
                    foreach (string file in files) File.Delete(file);
                    foreach (string directory in directories) Directory.Delete(directory, false);
                }
                catch (Exception error) when (error is IOException || error is UnauthorizedAccessException)
                {
                    // A locked file, link or uncertain shutdown is retained, never forced or retried.
                    LoadLog.Write("Editor profile retained: " + error.GetType().Name);
                }
            });
        }

        /// <summary>Checks every entry before deletion; links are refused before traversing them.</summary>
        /// <param name="directory">Profile directory to inspect. Any reparse point aborts traversal before its target is followed.</param>
        /// <param name="files">Receives regular file paths found beneath <paramref name="directory"/>.</param>
        /// <param name="directories">Receives directories in child-before-parent order for safe removal after files are deleted.</param>
        private static void CollectOwnedTree(string directory, List<string> files, List<string> directories)
        {
            if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("The editor profile contains a reparse point.");
            foreach (string entry in Directory.GetFileSystemEntries(directory))
            {
                FileAttributes attributes = File.GetAttributes(entry);
                if ((attributes & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("The editor profile contains a reparse point.");
                if ((attributes & FileAttributes.Directory) != 0) CollectOwnedTree(entry, files, directories);
                else files.Add(entry);
            }
            directories.Add(directory);
        }
    }
}
