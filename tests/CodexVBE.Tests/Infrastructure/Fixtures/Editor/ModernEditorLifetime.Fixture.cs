using System;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
namespace CodexVBE.Tests.Unit.Editor
{
    internal sealed class OwnedEditorDispatcher : SynchronizationContext, IDisposable
    {
        private readonly SynchronizationContext previous = Current;
        private int active;
        private readonly System.Collections.Generic.List<Exception> failures = new System.Collections.Generic.List<Exception>();
        internal OwnedEditorDispatcher() { SetSynchronizationContext(this); }
        public override void Post(SendOrPostCallback callback, object value) => previous.Post(state => { try { callback(state); } catch (Exception error) { failures.Add(error); } }, value);
        public override void Send(SendOrPostCallback callback, object value) => previous.Send(callback, value);
        public override void OperationStarted() { Interlocked.Increment(ref active); }
        public override void OperationCompleted() { Interlocked.Decrement(ref active); }
        internal void Drain(Type expectedFailure = null)
        {
            ModernEditorDebugFixture.Wait(System.Threading.Tasks.Task.Run(() => { var limit = System.Diagnostics.Stopwatch.StartNew(); while (Volatile.Read(ref active) != 0) { if (limit.ElapsedMilliseconds > 5000) throw new TimeoutException("Owned editor events did not complete"); Thread.Sleep(1); } }));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(expectedFailure == null ? 0 : 1, failures.Count, "Unexpected asynchronous editor failure");
            if (expectedFailure != null) Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsInstanceOfType(failures[0], expectedFailure);
            failures.Clear();
        }
        public void Dispose() { Drain(); SetSynchronizationContext(previous); }
    }
    internal sealed class EditorLifetimeModule : IEditorModule
    {
        internal string Code = "Public Sub Owned()\nEnd Sub";
        internal bool FailName;
        internal readonly string Recovery = Guid.NewGuid().ToString();
        public string Name => FailName ? throw new InvalidOperationException("owned module name unavailable") : "OwnedLifetime";
        public string Key => Recovery;
        public bool CanWrite => true;
        public string Read() => Code;
        public string Write(string expected, string text) { if (expected != Code) throw new InvalidOperationException("owned stale write"); return Code = text; }
        public void ShowNative(int line, int column) { }
    }
}