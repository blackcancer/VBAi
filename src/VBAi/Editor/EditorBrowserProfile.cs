using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns one editor-only profile; retirement never deletes another window's or a legacy profile.</summary>
    internal sealed class EditorBrowserProfile
    {

        /// <summary>Maintains the gate state for editor browser profile.</summary>
        private readonly object gate = new object();

        /// <summary>Maintains the requested and retired and exited and scheduled state for editor browser profile.</summary>
        private bool requested, retired, exited, scheduled;

        /// <summary>Identifies the browser process id associated with editor browser profile.</summary>
        private uint browserProcessId;

        /// <summary>The unique user data folder supplied only to this environment.</summary>
        /// <value>Current path exposed by editor browser profile.</value>
        internal string Path { get; }

        /// <summary>The asynchronous filesystem cleanup, scheduled only after retirement and proven runtime exit.</summary>
        /// <value>Current cleanup exposed by editor browser profile.</value>
        internal Task Cleanup { get; private set; } = Task.CompletedTask;

        /// <summary>Creates an operation-owned profile below the supplied editor cache root.</summary>
        /// <param name="root">Text that supplies the root value. Use the format required by the calling operation.</param>
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
        /// <param name="processId">uint that supplies the process id for this operation.</param>
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
        /// <param name="processId">uint that supplies the process id for this operation.</param>
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
            Cleanup = Task.Run(() => {
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
        /// <param name="directory">Text that supplies the directory value. Use the format required by the calling operation.</param>
        /// <param name="files">list&lt;string&gt; that supplies the files for this operation.</param>
        /// <param name="directories">list&lt;string&gt; that supplies the directories for this operation.</param>
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
