using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Web.Script.Serialization;
using System.Linq;

namespace VBAi
{

    /// <summary>Serializes SQLite writes without moving live chat, UI or COM objects off their thread.</summary>
    internal sealed class ChatPersistenceWorker : IDisposable
    {

        /// <summary>Owns the snapshot state and operations.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Identifies the writer and id and scope and title and payload and expected version associated with snapshot.</summary>
            internal readonly string Writer, Id, Scope, Title, Payload, ExpectedVersion;

            /// <summary>Maintains the deletion state for snapshot.</summary>
            internal readonly System.Threading.Tasks.TaskCompletionSource<Exception> Deletion;

            /// <summary>Initializes a Snapshot instance with the supplied state.</summary>
            /// <param name="session">chat session state that supplies the session for this operation.</param>
            /// <param name="payload">Text that supplies the payload value. Use the format required by the calling operation.</param>
            internal Snapshot(ChatSessionState session, string payload)
            {
                Writer = session.WriterId; Id = session.Id; Scope = session.Scope;
                Title = session.Title; Payload = payload; ExpectedVersion = session.StorageVersion;
            }

            /// <summary>Initializes a Snapshot instance with the supplied state.</summary>
            /// <param name="source">snapshot that supplies the source for this operation.</param>
            internal Snapshot(Snapshot source)
            {
                Writer = source.Writer; Id = source.Id; Scope = source.Scope;
                Title = source.Title; Payload = source.Payload; ExpectedVersion = source.ExpectedVersion;
                Deletion = new System.Threading.Tasks.TaskCompletionSource<Exception>(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously);
            }
        }

        /// <summary>Owns the cursor state and operations.</summary>
        private sealed class Cursor
        {

            /// <summary>Maintains the version state for cursor.</summary>
            internal string Version;

            /// <summary>Maintains the failure state for cursor.</summary>
            internal Exception Failure;
        }

        /// <summary>Maintains the gate state for chat persistence worker.</summary>
        private readonly object gate = new object();

        /// <summary>Maintains the pending state for chat persistence worker.</summary>
        private readonly Dictionary<string, Snapshot> pending = new Dictionary<string, Snapshot>();

        /// <summary>Maintains the order state for chat persistence worker.</summary>
        private readonly Queue<string> order = new Queue<string>();

        /// <summary>Maintains the cursors state for chat persistence worker.</summary>
        private readonly Dictionary<string, Cursor> cursors = new Dictionary<string, Cursor>();

        /// <summary>Maintains the deleting writers state for chat persistence worker.</summary>
        private readonly HashSet<string> deletingWriters = new HashSet<string>();

        /// <summary>Maintains the thread state for chat persistence worker.</summary>
        private readonly Thread thread;

        /// <summary>Keeps the path path available to chat persistence worker.</summary>
        private readonly string path;

        /// <summary>Maintains the completed state for chat persistence worker.</summary>
        private readonly Action<Snapshot, string, Exception> completed;

        /// <summary>Maintains the stopping and working state for chat persistence worker.</summary>
        private bool stopping, working;

        /// <summary>Maintains the active state for chat persistence worker.</summary>
        private Snapshot active;

        /// <summary>Gets or sets the recovery directory.</summary>
        /// <value>Current recovery directory exposed by chat persistence worker.</value>
        internal string RecoveryDirectory { get; private set; }

        /// <summary>Initializes a ChatPersistenceWorker instance with the supplied state.</summary>
        /// <param name="databasePath">Path used for the database path being processed.</param>
        /// <param name="completed">Exception describing the completed failure.</param>
        internal ChatPersistenceWorker(string databasePath, Action<Snapshot, string, Exception> completed)
        {
            path = Path.GetFullPath(databasePath);
            RecoveryDirectory = path + ".recovery";
            this.completed = completed;
            thread = new Thread(Run) { IsBackground = true, Name = "VBAi chat persistence" };
            thread.Start();
        }

        /// <summary>Handles enqueue for chat persistence worker.</summary>
        /// <param name="snapshot">snapshot that supplies the snapshot for this operation.</param>
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
        /// <param name="snapshot">snapshot that supplies the snapshot for this operation.</param>
        /// <returns>task&lt;exception&gt; produced by the operation for delete async on chat persistence worker.</returns>
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

        /// <summary>Runs  for chat persistence worker.</summary>
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

        /// <summary>Handles recovery path for chat persistence worker.</summary>
        /// <param name="snapshot">snapshot that supplies the snapshot for this operation.</param>
        /// <returns>Text produced by the operation for recovery path on chat persistence worker.</returns>
        private string RecoveryPath(Snapshot snapshot) => Path.Combine(RecoveryDirectory, snapshot.Writer + ".json");

        /// <summary>Writes recovery for chat persistence worker.</summary>
        /// <param name="snapshot">snapshot that supplies the snapshot for this operation.</param>
        /// <param name="destination">Text that supplies the destination value. Use the format required by the calling operation.</param>
        private void WriteRecovery(Snapshot snapshot, string destination = null)
        {
            if (ChatSessionStore.IsTransientScope(snapshot.Scope)) return;
            // Same local-data privacy boundary as chat.db; never replay this file automatically.
            Directory.CreateDirectory(RecoveryDirectory);
            var serializer = new JavaScriptSerializer { MaxJsonLength = 64 * 1024 * 1024 };
            UpdatePaths.WriteAtomic(destination ?? RecoveryPath(snapshot), serializer.Serialize(new {
                FormatVersion = 1, snapshot.Id, snapshot.Scope, snapshot.ExpectedVersion,
                SavedUtc = DateTime.UtcNow.ToString("O"), snapshot.Payload
            }));
        }

        /// <summary>Handles flush for chat persistence worker.</summary>
        /// <param name="milliseconds">int that supplies the milliseconds for this operation.</param>
        /// <returns>Boolean indicating the result of the check for flush on chat persistence worker.</returns>
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

        /// <summary>Disposes  for chat persistence worker.</summary>
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
