using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Serializes SQLite writes without moving live chat, UI or COM objects off their thread.</summary>
    internal sealed class ChatPersistenceWorker : IDisposable
    {

        /// <summary>Immutable persistence payload and optimistic-concurrency version captured on the chat's owning thread.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Writer identity, session key/scope/title, serialized payload, and storage version used by one queued write.</summary>
            internal readonly string Writer, Id, Scope, Title, Payload, ExpectedVersion;

            /// <summary>Completion source present only for a deletion snapshot; resolves with a recovery-file warning or faults on delete failure.</summary>
            internal readonly System.Threading.Tasks.TaskCompletionSource<Exception> Deletion;

            /// <summary>Captures a chat session's storage identity and current serialized payload for background persistence.</summary>
            /// <param name="session">Session state already snapshotted on the UI thread.</param>
            /// <param name="payload">Serialized session data to enqueue.</param>
            internal Snapshot(ChatSessionState session, string payload)
            {
                Writer = session.WriterId; Id = session.Id; Scope = session.Scope;
                Title = session.Title; Payload = payload; ExpectedVersion = session.StorageVersion;
            }

            /// <summary>Creates the deletion marker from a save snapshot and gives it an asynchronous completion source.</summary>
            /// <param name="source">Existing snapshot whose writer, session key, scope, title, payload, and expected version identify the queued conversation.</param>
            internal Snapshot(Snapshot source)
            {
                Writer = source.Writer; Id = source.Id; Scope = source.Scope;
                Title = source.Title; Payload = source.Payload; ExpectedVersion = source.ExpectedVersion;
                Deletion = new System.Threading.Tasks.TaskCompletionSource<Exception>(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously);
            }
        }

        /// <summary>Tracks the latest storage version and a sticky failure for one writer until that conversation is reopened.</summary>
        private sealed class Cursor
        {

            /// <summary>Version returned by the most recent successful write for this writer.</summary>
            internal string Version;

            /// <summary>First storage failure; blocks later queued writes to avoid overwriting an unknown database state.</summary>
            internal Exception Failure;
        }

        /// <summary>Protects the pending queue, writer deletion claims, cursors, and worker lifecycle flags.</summary>
        private readonly object gate = new object();

        /// <summary>Newest coalesced save snapshot for each writer, or its ordered deletion marker.</summary>
        private readonly Dictionary<string, Snapshot> pending = new Dictionary<string, Snapshot>();

        /// <summary>Writer IDs in first-enqueued order; repeated saves replace payloads without duplicating queue entries.</summary>
        private readonly Queue<string> order = new Queue<string>();

        /// <summary>Per-writer expected-version and failure state, retained to serialize optimistic-concurrency writes.</summary>
        private readonly Dictionary<string, Cursor> cursors = new Dictionary<string, Cursor>();

        /// <summary>Writers whose delete is queued or committed; new saves for them are rejected.</summary>
        private readonly HashSet<string> deletingWriters = new HashSet<string>();

        /// <summary>Single background thread that owns SQLite persistence operations.</summary>
        private readonly Thread thread;

        /// <summary>Canonical full path to the chat database opened by the worker thread.</summary>
        private readonly string path;

        /// <summary>UI notification invoked after a save attempt with its captured snapshot, resulting version, and failure.</summary>
        private readonly Action<Snapshot, string, Exception> completed;

        /// <summary>Lifecycle flags protected by <see cref="gate"/>: shutdown has started or the worker is processing a snapshot.</summary>
        private bool stopping, working;

        /// <summary>Snapshot currently being written or deleted, protected by <see cref="gate"/>.</summary>
        private Snapshot active;

        /// <summary>Gets the local recovery directory used to retain failed or shutdown-interrupted save snapshots.</summary>
        /// <value>Database path with the <c>.recovery</c> suffix.</value>
        internal string RecoveryDirectory { get; private set; }

        /// <summary>Starts the dedicated SQLite worker and derives its adjacent local recovery directory.</summary>
        /// <param name="databasePath">Chat database path; stored as a full path before the worker starts.</param>
        /// <param name="completed">Callback for completed save attempts; deletion completion is returned through the snapshot task.</param>
        internal ChatPersistenceWorker(string databasePath, Action<Snapshot, string, Exception> completed)
        {
            path = Path.GetFullPath(databasePath);
            RecoveryDirectory = path + ".recovery";
            this.completed = completed;
            thread = new Thread(Run) { IsBackground = true, Name = "VBAi chat persistence" };
            thread.Start();
        }

        /// <summary>Coalesces the newest save per writer and wakes the worker; transient scopes are never persisted.</summary>
        /// <param name="snapshot">Immutable UI-thread snapshot containing the serialized chat and expected storage version.</param>
        internal void Enqueue(Snapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (ChatSessionStore.IsTransientScope(snapshot.Scope)) return;
            lock (gate)
            {
                if (stopping) throw new ObjectDisposedException(nameof(ChatPersistenceWorker));
                if (deletingWriters.Contains(snapshot.Writer)) throw new InvalidOperationException("This conversation is being deleted or has been deleted.");
                if (!pending.ContainsKey(snapshot.Writer))
                {
                    if (pending.Count >= 64) throw new IOException("The history save queue is full. Your draft remains in this window.");
                    order.Enqueue(snapshot.Writer);
                }
                pending[snapshot.Writer] = snapshot;
                Monitor.PulseAll(gate);
            }
        }

        /// <summary>Orders deletion after any in-flight write and replaces queued writes for this conversation.</summary>
        /// <param name="snapshot">Snapshot identifying the conversation to delete; its current pending save is replaced by this ordered delete.</param>
        /// <returns>Task that completes after deletion, with an exception value if only recovery-file cleanup failed.</returns>
        internal System.Threading.Tasks.Task<Exception> DeleteAsync(Snapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (ChatSessionStore.IsTransientScope(snapshot.Scope)) return System.Threading.Tasks.Task.FromResult<Exception>(null);
            lock (gate)
            {
                if (stopping) throw new ObjectDisposedException(nameof(ChatPersistenceWorker));
                if (deletingWriters.Contains(snapshot.Writer)) throw new InvalidOperationException("This conversation is already being deleted.");
                if (!pending.ContainsKey(snapshot.Writer))
                {
                    if (pending.Count >= 64) throw new IOException("The history save queue is full. Your draft remains in this window.");
                    order.Enqueue(snapshot.Writer);
                }
                snapshot = new Snapshot(snapshot);
                deletingWriters.Add(snapshot.Writer);
                pending[snapshot.Writer] = snapshot;
                Monitor.PulseAll(gate);
                return snapshot.Deletion.Task;
            }
        }

        /// <summary>Drains queued writer snapshots in order, retaining local recovery before each save and serializing SQLite access.</summary>
        private void Run()
        {
            ChatSessionStore store = null;
            try
            {
                while (true)
                {
                    Snapshot snapshot;
                    lock (gate)
                    {
                        while (order.Count == 0 && !stopping) Monitor.Wait(gate);
                        if (order.Count == 0) return;
                        string writer = order.Dequeue(); snapshot = pending[writer]; pending.Remove(writer); working = true; active = snapshot;
                    }
                    string version = null; Exception failure = null, recoveryWarning = null;
                    try
                    {
                        // A distinct recovery file per loaded conversation prevents cross-host overwrites.
                        if (snapshot.Deletion == null) WriteRecovery(snapshot);
                        if (!cursors.TryGetValue(snapshot.Writer, out var cursor))
                            cursors.Add(snapshot.Writer, cursor = new Cursor { Version = snapshot.ExpectedVersion });
                        if (cursor.Failure != null) throw new IOException("Reopen this conversation before saving again. Its local recovery copy has been retained.", cursor.Failure);
                        try
                        {
                            if (store == null) store = new ChatSessionStore(path);
                            if (snapshot.Deletion != null) store.Delete(snapshot.Id, snapshot.Scope, cursor.Version);
                            else
                            {
                                version = store.SavePayload(snapshot.Id, snapshot.Scope, snapshot.Title, snapshot.Payload, cursor.Version);
                                cursor.Version = version;
                            }
                        }
                        catch (Exception error) { cursor.Failure = error; throw; }
                        if (snapshot.Deletion == null) File.Delete(RecoveryPath(snapshot));
                        else
                        {
                            try { File.Delete(RecoveryPath(snapshot)); }
                            catch (Exception error) when (error is IOException || error is UnauthorizedAccessException)
                            {
                                recoveryWarning = error;
                                LoadLog.Write("Conversation deleted, but its local recovery copy could not be removed: " + error.GetType().Name);
                            }
                        }
                    }
                    catch (Exception error) { failure = error; }
                    try
                    {
                        if (snapshot.Deletion == null) completed?.Invoke(snapshot, version, failure);
                        else
                        {
                            if (failure == null) snapshot.Deletion.TrySetResult(recoveryWarning);
                            else
                            {
                                lock (gate) deletingWriters.Remove(snapshot.Writer);
                                snapshot.Deletion.TrySetException(failure);
                            }
                        }
                    }
                    catch (Exception error) { LoadLog.Write("Chat persistence notification failed: " + error.GetType().Name); }
                    finally
                    {
                        lock (gate) { working = false; active = null; Monitor.PulseAll(gate); }
                    }
                }
            }
            finally { store?.Dispose(); }
        }

        /// <summary>Builds the stable recovery filename for one writer's conversation snapshot.</summary>
        /// <param name="snapshot">Snapshot whose writer ID determines the recovery filename.</param>
        /// <returns>Path under <see cref="RecoveryDirectory"/> ending in the writer ID and <c>.json</c>.</returns>
        private string RecoveryPath(Snapshot snapshot) => Path.Combine(RecoveryDirectory, snapshot.Writer + ".json");

        /// <summary>Atomically writes a local recovery snapshot; this file is retained for user recovery and is never replayed automatically.</summary>
        /// <param name="snapshot">Captured conversation data to serialize; transient scopes are skipped.</param>
        /// <param name="destination">Optional shutdown-specific filename; null uses the stable per-writer recovery path.</param>
        private void WriteRecovery(Snapshot snapshot, string destination = null)
        {
            if (ChatSessionStore.IsTransientScope(snapshot.Scope)) return;
            // Same local-data privacy boundary as chat.db; never replay this file automatically.
            Directory.CreateDirectory(RecoveryDirectory);
            var serializer = new JavaScriptSerializer { MaxJsonLength = 64 * 1024 * 1024 };
            UpdatePaths.WriteAtomic(destination ?? RecoveryPath(snapshot), serializer.Serialize(new
            {
                FormatVersion = 1,
                snapshot.Id,
                snapshot.Scope,
                snapshot.ExpectedVersion,
                SavedUtc = DateTime.UtcNow.ToString("O"),
                snapshot.Payload
            }));
        }

        /// <summary>Waits for active and queued work to drain, returning false when the time budget expires.</summary>
        /// <param name="milliseconds">Maximum wait time in milliseconds.</param>
        /// <returns><see langword="true"/> when no write or delete remains active or queued.</returns>
        internal bool Flush(int milliseconds)
        {
            var elapsed = System.Diagnostics.Stopwatch.StartNew();
            lock (gate)
            {
                while (working || order.Count != 0)
                {
                    int remaining = milliseconds - (int)elapsed.ElapsedMilliseconds;
                    if (remaining <= 0) return false;
                    Monitor.Wait(gate, remaining);
                }
                return true;
            }
        }

        /// <summary>Stops the worker and waits up to two seconds; remaining save snapshots are retained as uniquely named local recovery files.</summary>
        public void Dispose()
        {
            lock (gate) { stopping = true; Monitor.PulseAll(gate); }
            if (!thread.Join(2000))
            {
                Snapshot[] remaining;
                lock (gate)
                    remaining = (active == null ? pending.Values : pending.Values.Concat(new[] { active }))
                        .GroupBy(item => item.Writer).Select(group => group.First()).ToArray();
                // Only shutdown may perform this fallback I/O on the caller. Separate names prevent
                // an in-flight older commit from deleting a newer emergency snapshot.
                foreach (var snapshot in remaining.Where(item => item.Deletion == null))
                    WriteRecovery(snapshot, Path.Combine(RecoveryDirectory, "closing-" + Guid.NewGuid().ToString("N") + ".json"));
                LoadLog.Write("Chat persistence close exceeded its drain budget. Remaining save snapshots were retained as local recovery copies; pending deletions continue on the worker.");
            }
        }
    }
}
