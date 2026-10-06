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
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <param name="supportSignature">Text that supplies the support signature value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for request on i vba test runtime.</returns>
        [DispId(1)] object Request(string supportVersion, string supportSignature);

        /// <summary>Publishes one verdict bound to the claimed nonce, revision, run and test.</summary>
        /// <param name="nonce">Text that supplies the nonce value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="runId">Text that supplies the run id value. Use the format required by the calling operation.</param>
        /// <param name="testId">Text that supplies the test id value. Use the format required by the calling operation.</param>
        /// <param name="status">Text that supplies the status value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <param name="errorNumber">int that supplies the error number for this operation.</param>
        /// <returns>Boolean indicating the result of the check for publish on i vba test runtime.</returns>
        [DispId(2)] bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber);
    }

    /// <summary>A factory-created callback handle; it cannot arm, dispatch, cancel or schedule a test.</summary>
    [ComVisible(true), Guid("5AF2F40B-939B-4CC6-A06C-F0C79841C031"), ProgId("VBAi.TestRuntime"),
        ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestRuntime))]
    public sealed class VbaTestRuntime : IVbaTestRuntime
    {

        /// <summary>Maintains the pending state for vba test runtime.</summary>
        [ThreadStatic] private static VbaTestResultSink pending;

        /// <summary>Maintains the sink state for vba test runtime.</summary>
        private readonly VbaTestResultSink sink;

        /// <summary>Maintains the binding state for vba test runtime.</summary>
        private readonly string binding;

        /// <summary>Binds only to an already dispatched attempt on the calling VBA thread.</summary>
        public VbaTestRuntime()
        {
            sink = pending ?? throw new InvalidOperationException("No authorized test attempt exists on this VBA thread.");
            binding = sink.BindRuntime();
        }

        /// <inheritdoc/>
        /// <summary>Handles request for vba test runtime.</summary>
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <param name="supportSignature">Text that supplies the support signature value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for request on vba test runtime.</returns>
        public object Request(string supportVersion, string supportSignature) => sink.RequestLocal(binding, supportVersion, supportSignature);

        /// <inheritdoc/>
        /// <summary>Handles publish for vba test runtime.</summary>
        /// <param name="nonce">Text that supplies the nonce value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="runId">Text that supplies the run id value. Use the format required by the calling operation.</param>
        /// <param name="testId">Text that supplies the test id value. Use the format required by the calling operation.</param>
        /// <param name="status">Text that supplies the status value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <param name="errorNumber">int that supplies the error number for this operation.</param>
        /// <returns>Boolean indicating the result of the check for publish on vba test runtime.</returns>
        public bool Publish(string nonce, string revision, string runId, string testId, string status, string message, int errorNumber)
            => sink.PublishLocal(binding, nonce, revision, runId, testId, status, message, errorNumber);

        /// <summary>Handles expose for vba test runtime.</summary>
        /// <param name="value">vba test result sink that supplies the value for this operation.</param>
        /// <returns>i disposable produced by the operation for expose on vba test runtime.</returns>
        internal static IDisposable Expose(VbaTestResultSink value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (pending != null) throw new InvalidOperationException("Another test runtime is already exposed on this VBA thread.");
            value.BindRuntime();
            pending = value;
            return new Exposure(value);
        }

        /// <summary>Owns the exposure state and operations.</summary>
        private sealed class Exposure : IDisposable
        {

            /// <summary>Maintains the owner state for exposure.</summary>
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;

            /// <summary>Maintains the value state for exposure.</summary>
            private VbaTestResultSink value;

            /// <summary>Initializes a Exposure instance with the supplied state.</summary>
            /// <param name="value">vba test result sink that supplies the value for this operation.</param>
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
