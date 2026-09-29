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
        internal sealed class Snapshot
        {
            internal readonly string Writer, Id, Scope, Title, Payload, ExpectedVersion;
            internal Snapshot(ChatSessionState session, string payload)
            {
                Writer = session.WriterId; Id = session.Id; Scope = session.Scope;
                Title = session.Title; Payload = payload; ExpectedVersion = session.StorageVersion;
            }
        }
        private sealed class Cursor
        {
            internal string Version;
            internal Exception Failure;
        }
        private readonly object gate = new object();
        private readonly Dictionary<string, Snapshot> pending = new Dictionary<string, Snapshot>();
        private readonly Queue<string> order = new Queue<string>();
        private readonly Dictionary<string, Cursor> cursors = new Dictionary<string, Cursor>();
        private readonly Thread thread;
        private readonly string path;
        private readonly Action<Snapshot, string, Exception> completed;
        private bool stopping, working;
        private Snapshot active;
        internal string RecoveryDirectory { get; private set; }

        internal ChatPersistenceWorker(string databasePath, Action<Snapshot, string, Exception> completed)
        {
            path = Path.GetFullPath(databasePath);
            RecoveryDirectory = path + ".recovery";
            this.completed = completed;
            thread = new Thread(Run) { IsBackground = true, Name = "VBAi chat persistence" };
            thread.Start();
        }

        internal void Enqueue(Snapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            lock (gate)
            {
                if (stopping) throw new ObjectDisposedException(nameof(ChatPersistenceWorker));
                if (!pending.ContainsKey(snapshot.Writer))
                {
                    if (pending.Count >= 64) throw new IOException("The history save queue is full. Your draft remains in this window.");
                    order.Enqueue(snapshot.Writer);
                }
                pending[snapshot.Writer] = snapshot;
                Monitor.PulseAll(gate);
            }
        }

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
                    string version = null; Exception failure = null;
                    try
                    {
                        // A distinct recovery file per loaded conversation prevents cross-host overwrites.
                        WriteRecovery(snapshot);
                        if (!cursors.TryGetValue(snapshot.Writer, out var cursor))
                            cursors.Add(snapshot.Writer, cursor = new Cursor { Version = snapshot.ExpectedVersion });
                        if (cursor.Failure != null) throw new IOException("Reopen this conversation before saving again. Its local recovery copy has been retained.", cursor.Failure);
                        try
                        {
                            if (store == null) store = new ChatSessionStore(path);
                            version = store.SavePayload(snapshot.Id, snapshot.Scope, snapshot.Title, snapshot.Payload, cursor.Version);
                            cursor.Version = version;
                        }
                        catch (Exception error) { cursor.Failure = error; throw; }
                        File.Delete(RecoveryPath(snapshot));
                    }
                    catch (Exception error) { failure = error; }
                    try { completed?.Invoke(snapshot, version, failure); }
                    catch (Exception error) { LoadLog.Write("Chat persistence notification failed: " + error.GetType().Name); }
                    finally
                    {
                        lock (gate) { working = false; active = null; Monitor.PulseAll(gate); }
                    }
                }
            }
            finally { store?.Dispose(); }
        }

        private string RecoveryPath(Snapshot snapshot) => Path.Combine(RecoveryDirectory, snapshot.Writer + ".json");

        private void WriteRecovery(Snapshot snapshot, string destination = null)
        {
            // Same local-data privacy boundary as chat.db; never replay this file automatically.
            Directory.CreateDirectory(RecoveryDirectory);
            var serializer = new JavaScriptSerializer { MaxJsonLength = 64 * 1024 * 1024 };
            UpdatePaths.WriteAtomic(destination ?? RecoveryPath(snapshot), serializer.Serialize(new {
                FormatVersion = 1, snapshot.Id, snapshot.Scope, snapshot.ExpectedVersion,
                SavedUtc = DateTime.UtcNow.ToString("O"), snapshot.Payload
            }));
        }

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
                foreach (var snapshot in remaining)
                    WriteRecovery(snapshot, Path.Combine(RecoveryDirectory, "closing-" + Guid.NewGuid().ToString("N") + ".json"));
                LoadLog.Write("Chat persistence close exceeded its drain budget. Remaining snapshots were saved as local recovery copies.");
            }
        }
    }
}
