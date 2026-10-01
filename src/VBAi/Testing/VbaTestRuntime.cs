using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi
{
    /// <summary>Late-bound test-result contract, available without a host-specific VBE accessor.</summary>
    [ComVisible(true), Guid("6905357A-66DF-4A81-9902-4FA2675AC470"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVbaTestRuntime
    {
        /// <summary>Returns the seven-field pending call for this owning STA and support signature.</summary>
        [DispId(1)] object Request(string supportVersion, string supportSignature);
        /// <summary>Publishes one verdict bound to the claimed nonce, revision, run and test.</summary>
        [DispId(2)] bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber);
    }

    /// <summary>A factory-created callback handle; it cannot arm, dispatch, cancel or schedule a test.</summary>
    [ComVisible(true), Guid("5AF2F40B-939B-4CC6-A06C-F0C79841C031"), ProgId("VBAi.TestRuntime"),
        ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestRuntime))]
    public sealed class VbaTestRuntime : IVbaTestRuntime
    {
        [ThreadStatic] private static VbaTestResultSink pending;
        private readonly VbaTestResultSink sink;
        private readonly string binding;

        /// <summary>Binds only to an already dispatched attempt on the calling VBA thread.</summary>
        public VbaTestRuntime()
        {
            sink = pending ?? throw new InvalidOperationException("No authorized test attempt exists on this VBA thread.");
            binding = sink.BindRuntime();
        }

        /// <inheritdoc/>
        public object Request(string supportVersion, string supportSignature) => sink.RequestLocal(binding, supportVersion, supportSignature);
        /// <inheritdoc/>
        public bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber)
            => sink.PublishLocal(binding, nonce, revision, runId, testId, status, message, errorNumber);

        internal static IDisposable Expose(VbaTestResultSink value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (pending != null) throw new InvalidOperationException("Another test runtime is already exposed on this VBA thread.");
            value.BindRuntime();
            pending = value;
            return new Exposure(value);
        }

        private sealed class Exposure : IDisposable
        {
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            private VbaTestResultSink value;
            internal Exposure(VbaTestResultSink value) { this.value = value; }
            public void Dispose()
            {
                if (owner != Thread.CurrentThread.ManagedThreadId)
                    throw new InvalidOperationException("The runtime exposure must be revoked on its owning thread.");
                if (value == null) return;
                if (!ReferenceEquals(pending, value)) throw new InvalidOperationException("The runtime exposure changed unexpectedly.");
                pending = null; value = null;
            }
        }
    }
}
