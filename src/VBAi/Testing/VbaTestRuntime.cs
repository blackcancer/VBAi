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
        /// <param name="supportVersion">Version reported by the generated test support module.</param>
        /// <param name="supportSignature">Signature reported by that module for this test attempt.</param>
        /// <returns>Pending attempt fields for the calling VBA thread, or an empty result when the support contract does not match.</returns>
        [DispId(1)] object Request(string supportVersion, string supportSignature);

        /// <summary>Publishes one verdict bound to the claimed nonce, revision, run and test.</summary>
        /// <param name="nonce">One-use nonce returned by <see cref="Request"/> for this attempt.</param>
        /// <param name="revision">Project revision attached to the dispatched attempt.</param>
        /// <param name="runId">Run identifier attached to the dispatched attempt.</param>
        /// <param name="testId">Discovered test identifier being reported.</param>
        /// <param name="status">Supported terminal verdict such as Passed, Failed, or Skipped.</param>
        /// <param name="message">Bounded diagnostic text associated with the verdict.</param>
        /// <param name="errorNumber">VBA error number, or zero when no VBA error occurred.</param>
        /// <returns>True only when the complete verdict matches the pending attempt and is accepted once.</returns>
        [DispId(2)] bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber);
    }

    /// <summary>A factory-created callback handle; it cannot arm, dispatch, cancel or schedule a test.</summary>
    [ComVisible(true), Guid("5AF2F40B-939B-4CC6-A06C-F0C79841C031"), ProgId("VBAi.TestRuntime"),
        ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestRuntime))]
    public sealed class VbaTestRuntime : IVbaTestRuntime
    {

        /// <summary>Sink currently exposed to the active test on this VBA thread; never shared across threads.</summary>
        [ThreadStatic] private static VbaTestResultSink pending;

        /// <summary>Sink bound to this callback handle's single dispatched attempt.</summary>
        private readonly VbaTestResultSink sink;

        /// <summary>Opaque binding token minted by the sink when this handle is constructed.</summary>
        private readonly string binding;

        /// <summary>Binds only to an already dispatched attempt on the calling VBA thread.</summary>
        public VbaTestRuntime()
        {
            sink = pending ?? throw new InvalidOperationException("No authorized test attempt exists on this VBA thread.");
            binding = sink.BindRuntime();
        }

        /// <inheritdoc/>
        /// <summary>Reads pending attempt data only when generated support version and signature match.</summary>
        /// <param name="supportVersion">Version reported by the generated support module.</param>
        /// <param name="supportSignature">Signature reported by the same support module.</param>
        /// <returns>Seven-field pending call data, or an empty response when the contract is stale or mismatched.</returns>
        public object Request(string supportVersion, string supportSignature) => sink.RequestLocal(binding, supportVersion, supportSignature);

        /// <inheritdoc/>
        /// <summary>Accepts a verdict only for this handle's current nonce, revision, run, and test.</summary>
        /// <param name="nonce">One-use nonce supplied by the pending call.</param>
        /// <param name="revision">Revision supplied by the pending call.</param>
        /// <param name="runId">Run ID supplied by the pending call.</param>
        /// <param name="testId">Test ID supplied by the pending call.</param>
        /// <param name="status">Terminal status token from the supported verdict set.</param>
        /// <param name="message">Diagnostic text to attach to the result.</param>
        /// <param name="errorNumber">VBA error number, or zero for a non-error result.</param>
        /// <returns>True when the sink accepts this verdict; duplicate or mismatched callbacks return false.</returns>
        public bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber)
            => sink.PublishLocal(binding, nonce, revision, runId, testId, status, message, errorNumber);

        /// <summary>Temporarily exposes a result sink so a COM callback can bind on the current VBA thread.</summary>
        /// <param name="value">Sink for the single active test attempt.</param>
        /// <returns>Scope that clears the thread-local exposure on disposal by the same thread.</returns>
        internal static IDisposable Expose(VbaTestResultSink value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (pending != null) throw new InvalidOperationException("Another test runtime is already exposed on this VBA thread.");
            value.BindRuntime();
            pending = value;
            return new Exposure(value);
        }

        /// <summary>Revokes one thread-local sink exposure and enforces same-thread disposal.</summary>
        private sealed class Exposure : IDisposable
        {

            /// <summary>Managed thread ID that created this exposure scope.</summary>
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;

            /// <summary>Sink whose pending thread-local exposure this scope owns.</summary>
            private VbaTestResultSink value;

            /// <summary>Captures the sink whose exposure this scope must revoke.</summary>
            /// <param name="value">Sink installed in the thread-local pending slot.</param>
            internal Exposure(VbaTestResultSink value) { this.value = value; }

            /// <summary>Disposes  for exposure.</summary>
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
