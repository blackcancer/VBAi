using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;

namespace VBAi
{

    /// <summary>Versioned project-local callback; it never executes VBA or grants execution authority.</summary>
    [ComVisible(true), Guid("DB1B3DAD-A7EC-48CD-A6BE-AB4667524D47"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVbaTestResultSink
    {

        /// <summary>Claims the pending call once; returns nonce, module, procedure, phase, revision, run and test.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="hostPath">Path used for the host path being processed.</param>
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for request on i vba test result sink.</returns>
        [DispId(1)] object Request(object project, string hostPath, string supportVersion);

        /// <summary>Publishes one bounded verdict for the claimed call; native completion remains mandatory.</summary>
        /// <param name="nonce">Text that supplies the nonce value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="status">Text that supplies the status value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <param name="errorNumber">int that supplies the error number for this operation.</param>
        /// <returns>Boolean indicating the result of the check for publish on i vba test result sink.</returns>
        [DispId(2)] bool Publish(string nonce, string revision, string status, string message, int errorNumber);
    }

    /// <summary>STA-owned bounded callback state, published through the existing VBIDE AddIn.Object property.</summary>
    [ComVisible(true), Guid("744E3C39-BF87-49EA-926A-8B6D07F1A0DF"), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestResultSink))]
    public sealed class VbaTestResultSink : IVbaTestResultSink, IDisposable
    {

        /// <summary>Maintains the owner thread state for vba test result sink.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Maintains the same project state for vba test result sink.</summary>
        private readonly Func<object, object, bool> sameProject;

        /// <summary>Maintains the gate state for vba test result sink.</summary>
        private readonly object gate = new object();

        /// <summary>Maintains the active state for vba test result sink.</summary>
        private Attempt active;

        /// <summary>Maintains the disposed and uncertain state for vba test result sink.</summary>
        private bool disposed, uncertain;

        /// <summary>Owns the attempt state and operations.</summary>
        private sealed class Attempt
        {

            /// <summary>Maintains the project state for attempt.</summary>
            internal object Project;

            /// <summary>Identifies the path and version and signature and nonce and revision and run id and module and procedure and test id and phase and fault associated with attempt.</summary>
            internal string Path, Version, Signature, Nonce, Revision, RunId, Module, Procedure, TestId, Phase, Fault;

            /// <summary>Maintains the started and claimed and published state for attempt.</summary>
            internal bool Started, Claimed, Published;

            /// <summary>Maintains the verdict state for attempt.</summary>
            internal object[] Verdict;
        }

        /// <summary>Initializes a VbaTestResultSink instance with the supplied state.</summary>
        /// <param name="sameProject">func&lt;object, object, bool&gt; that supplies the same project for this operation.</param>
        internal VbaTestResultSink(Func<object, object, bool> sameProject = null)
        { this.sameProject = sameProject ?? VbeDebug.NativeProcedureValuesHost.SameComIdentity; }

        /// <summary>Handles arm for vba test result sink.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="runId">Text that supplies the run id value. Use the format required by the calling operation.</param>
        /// <param name="test">vba test descriptor that supplies the test for this operation.</param>
        /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
        /// <param name="supportSignature">Text that supplies the support signature value. Use the format required by the calling operation.</param>
        internal void Arm(object project, string path, string supportVersion, string revision,
            string runId, VbaTestDescriptor test, string phase, string supportSignature = null)
        {
            RequireOwner();
            lock (gate)
            {
                if (disposed || uncertain) throw new InvalidOperationException("The callback channel is closed or its preceding outcome is uncertain.");
                if (active != null) throw new InvalidOperationException("A callback attempt is already pending.");
                if (project == null || test == null || !Bound(path, 32768) || !Bound(supportVersion, 32)
                    || !Bound(revision, 128) || !Bound(runId, 128) || !Bound(test.Id, 128)
                    || !Bound(test.Module, 255) || !Bound(test.Procedure, 255) || !ValidPhase(phase)
                    || (supportSignature != null && !Bound(supportSignature, 128)))
                    throw new ArgumentException("An exact bounded callback plan is required.");
                var bytes = new byte[32];
                using (var random = RandomNumberGenerator.Create()) random.GetBytes(bytes);
                active = new Attempt { Project = project, Path = path, Version = supportVersion,
                    Signature = supportSignature, Revision = revision, RunId = runId, TestId = test.Id, Module = test.Module,
                    Procedure = test.Procedure, Phase = phase, Nonce = BitConverter.ToString(bytes).Replace("-", "") };
            }
        }

        /// <summary>Handles bind runtime for vba test result sink.</summary>
        /// <returns>Text produced by the operation for bind runtime on vba test result sink.</returns>
        internal string BindRuntime()
        {
            lock (gate) { RequireCallback(); return active.Nonce; }
        }

        /// <summary>Handles request local for vba test result sink.</summary>
        /// <param name="binding">Text that supplies the binding value. Use the format required by the calling operation.</param>
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <param name="supportSignature">Text that supplies the support signature value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for request local on vba test result sink.</returns>
        internal object RequestLocal(string binding, string supportVersion, string supportSignature)
        {
            lock (gate)
            {
                RequireCallback();
                if (active.Claimed || binding != active.Nonce || !Bound(supportSignature, 128)
                    || supportVersion != active.Version || supportSignature != active.Signature)
                    return Reject("The runtime support signature, version or claim does not match the pending attempt.");
                active.Claimed = true;
                return new object[] { active.Nonce, active.Module, active.Procedure, active.Phase,
                    active.Revision, active.RunId, active.TestId };
            }
        }

        /// <summary>Handles publish local for vba test result sink.</summary>
        /// <param name="binding">Text that supplies the binding value. Use the format required by the calling operation.</param>
        /// <param name="nonce">Text that supplies the nonce value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="runId">Text that supplies the run id value. Use the format required by the calling operation.</param>
        /// <param name="testId">Text that supplies the test id value. Use the format required by the calling operation.</param>
        /// <param name="status">Text that supplies the status value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <param name="errorNumber">int that supplies the error number for this operation.</param>
        /// <returns>Boolean indicating the result of the check for publish local on vba test result sink.</returns>
        internal bool PublishLocal(string binding, string nonce, string revision, string runId, string testId,
            string status, string message, int errorNumber)
        {
            lock (gate)
            {
                RequireCallback();
                if (binding != active.Nonce || runId != active.RunId || testId != active.TestId)
                { Reject("The runtime instance, run or test does not match the pending attempt."); return false; }
                return Publish(nonce, revision, status, message, errorNumber);
            }
        }

        /// <summary>Gets the has verdict.</summary>
        /// <value>Current has verdict exposed by vba test result sink.</value>
        internal bool HasVerdict { get { lock (gate) { return active?.Published == true; } } }

        /// <summary>Gets the has fault.</summary>
        /// <value>Current has fault exposed by vba test result sink.</value>
        internal bool HasFault { get { lock (gate) { return active?.Fault != null; } } }

        /// <summary>Handles begin native for vba test result sink.</summary>
        internal void BeginNative()
        {
            RequireOwner();
            lock (gate)
            {
                if (disposed || active == null || active.Started) throw new InvalidOperationException("No undispatched callback attempt exists.");
                active.Started = true;
            }
        }

        /// <inheritdoc/>
        /// <summary>Handles request for vba test result sink.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="hostPath">Path used for the host path being processed.</param>
        /// <param name="supportVersion">Text that supplies the support version value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for request on vba test result sink.</returns>
        public object Request(object project, string hostPath, string supportVersion)
        {
            lock (gate)
            {
                RequireCallback();
                if (active.Claimed || !Bound(hostPath, 32768) || !Bound(supportVersion, 32)
                    || !string.Equals(active.Path, hostPath, StringComparison.OrdinalIgnoreCase)
                    || active.Version != supportVersion || !sameProject(active.Project, project))
                    return Reject("The callback project, path, version or claim does not match the pending attempt.");
                active.Claimed = true;
                return new object[] { active.Nonce, active.Module, active.Procedure, active.Phase,
                    active.Revision, active.RunId, active.TestId };
            }
        }

        /// <inheritdoc/>
        /// <summary>Handles publish for vba test result sink.</summary>
        /// <param name="nonce">Text that supplies the nonce value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="status">Text that supplies the status value. Use the format required by the calling operation.</param>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <param name="errorNumber">int that supplies the error number for this operation.</param>
        /// <returns>Boolean indicating the result of the check for publish on vba test result sink.</returns>
        public bool Publish(string nonce, string revision, string status, string message, int errorNumber)
        {
            lock (gate)
            {
                RequireCallback();
                if (!active.Claimed || active.Published || nonce != active.Nonce || revision != active.Revision
                    || message == null || message.Length > VbaTestRuntimeSource.MaximumMessageLength
                    || (status != "Passed" && status != "Failed" && status != "Error" && status != "Inconclusive")
                    || ((status == "Passed" || status == "Inconclusive") && errorNumber != 0))
                { Reject("The callback verdict is invalid, repeated or belongs to another attempt."); return false; }
                active.Published = true;
                active.Verdict = new object[] { status, message, errorNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) };
                return true;
            }
        }

        /// <summary>Handles complete native for vba test result sink.</summary>
        /// <param name="completed">Indicates whether completed is enabled.</param>
        /// <param name="nativeError">Text that supplies the native error value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for complete native on vba test result sink.</returns>
        internal object CompleteNative(bool completed, string nativeError = null)
        {
            RequireOwner();
            lock (gate)
            {
                if (active == null) throw new InvalidOperationException("No callback attempt exists.");
                var attempt = active;
                active = null;
                if (!attempt.Started) throw new VbaTestInvocationException("The callback attempt was never dispatched.", false);
                if (disposed || !completed || attempt.Fault != null || !attempt.Claimed || !attempt.Published)
                {
                    uncertain = true;
                    throw new VbaTestInvocationException(attempt.Fault ?? nativeError ?? "The native call completed without a verified callback verdict.", true);
                }
                return attempt.Verdict;
            }
        }

        /// <summary>Determines whether it can cel undispatched for vba test result sink.</summary>
        internal void CancelUndispatched()
        {
            RequireOwner();
            lock (gate)
            {
                if (active?.Started == true) throw new InvalidOperationException("A dispatched attempt cannot be cancelled or retried.");
                active = null;
            }
        }

        /// <summary>Requires callback for vba test result sink.</summary>
        private void RequireCallback()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread || disposed || active == null || !active.Started || active.Fault != null)
                Reject("The callback channel is not available on this owning thread and attempt.");
        }

        /// <summary>Handles reject for vba test result sink.</summary>
        /// <param name="reason">Text that supplies the reason value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for reject on vba test result sink.</returns>
        private object Reject(string reason)
        {
            if (active?.Started == true && active.Fault == null) active.Fault = reason;
            throw new InvalidOperationException(reason);
        }

        /// <summary>Requires owner for vba test result sink.</summary>
        private void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("The callback channel requires its owning thread."); }

        /// <summary>Handles bound for vba test result sink.</summary>
        /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
        /// <param name="maximum">int that supplies the maximum for this operation.</param>
        /// <returns>Boolean indicating the result of the check for bound on vba test result sink.</returns>
        private static bool Bound(string value, int maximum) => !string.IsNullOrWhiteSpace(value) && value.Length <= maximum;

        /// <summary>Handles valid phase for vba test result sink.</summary>
        /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for valid phase on vba test result sink.</returns>
        private static bool ValidPhase(string value) => value == "Test" || value == "ModuleInitialize" || value == "ModuleCleanup"
            || value == "TestInitialize" || value == "TestCleanup";

        /// <summary>Revokes claims and publications; a dispatched call remains uncertain.</summary>
        public void Dispose()
        {
            RequireOwner();
            lock (gate) { disposed = true; if (active?.Started == true) active.Fault = "The callback channel disconnected during execution."; }
        }
    }
}
