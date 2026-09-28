using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed class EditorSyncPlan
    {
        internal readonly string Before, After;
        internal readonly Tuple<int, int, string> Patch;
        internal readonly int WorkerThreadId;
        internal EditorSyncPlan(string before, string after)
        {
            Before = before; After = after; EditorDocument.Validate(after);
            Patch = EditorDocument.Difference(before, after);
            WorkerThreadId = Thread.CurrentThread.ManagedThreadId;
        }
    }
    /// <summary>One background worker per editor. Only immutable text snapshots cross this boundary.</summary>
    internal sealed class EditorSyncWorker : IDisposable
    {
        private readonly BlockingCollection<Action> work = new BlockingCollection<Action>();
        private readonly Thread thread;
        internal EditorSyncWorker()
        {
            thread = new Thread(Run) { IsBackground = true, Name = "VBAi editor synchronization" };
            thread.Start();
        }
        private void Run()
        {
            try { foreach (var action in work.GetConsumingEnumerable()) action(); }
            finally { work.Dispose(); }
        }
        internal Task<T> Evaluate<T>(Func<T> action)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            try { work.Add(() => { try { completion.SetResult(action()); } catch (Exception error) { completion.SetException(error); } }); }
            catch (Exception error) { completion.SetException(error); }
            return completion.Task;
        }
        internal Task<EditorSyncPlan> Prepare(EditorDocument document, EditorDraftStore drafts)
        {
            // Capture on the UI thread. The worker never dereferences a document, COM proxy or control.
            var snapshot = new EditorDraft { Key = document.RecoveryKey, Baseline = document.Baseline, Text = document.Text };
            string id = document.Id;
            var completion = new TaskCompletionSource<EditorSyncPlan>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                work.Add(() =>
                {
                    try
                    {
                        drafts.SaveSnapshot(id, snapshot);
                        completion.SetResult(new EditorSyncPlan(snapshot.Baseline, snapshot.Text));
                    }
                    catch (Exception error) { completion.SetException(error); }
                });
            }
            catch (Exception error) { completion.SetException(error); }
            return completion.Task;
        }
        public void Dispose()
        {
            // Never wait on a background job from the host UI. Queued snapshots finish safely.
            try { work.CompleteAdding(); } catch (ObjectDisposedException) { }
        }
    }
}
