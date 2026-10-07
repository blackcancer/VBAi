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
        /// <param name="project">Live project object whose COM identity must match the armed attempt.</param>
        /// <param name="hostPath">Host document path used to bind the callback to the selected document.</param>
        /// <param name="supportVersion">Version token of the project-local callback module.</param>
        /// <returns>Array containing nonce, module, procedure, phase, revision, run ID, and test ID; null is never a valid claim.</returns>
        [DispId(1)] object Request(object project, string hostPath, string supportVersion);

        /// <summary>Publishes one bounded verdict for the claimed call; native completion remains mandatory.</summary>
        /// <param name="nonce">One-time token returned by the successful request claim.</param>
        /// <param name="revision">Source revision returned by that claim.</param>
        /// <param name="status">One of Passed, Failed, Error, or Inconclusive.</param>
        /// <param name="message">Bounded result text; must not be null.</param>
        /// <param name="errorNumber">VBA error number; must be zero for Passed and Inconclusive.</param>
        /// <returns>True only when the first valid verdict is accepted for the active claim.</returns>
        [DispId(2)] bool Publish(string nonce, string revision, string status, string message, int errorNumber);
    }

    /// <summary>STA-owned bounded callback state, published through the existing VBIDE AddIn.Object property.</summary>
    [ComVisible(true), Guid("744E3C39-BF87-49EA-926A-8B6D07F1A0DF"), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestResultSink))]
    public sealed class VbaTestResultSink : IVbaTestResultSink, IDisposable
    {

        /// <summary>Managed thread identifier required for all callback state transitions.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Compares callback project identity with the armed live project.</summary>
        private readonly Func<object, object, bool> sameProject;

        /// <summary>Protects callback claims, verdicts, disposal, and uncertainty state.</summary>
        private readonly object gate = new object();

        /// <summary>Single attempt currently eligible to claim or publish a result.</summary>
        private Attempt active;

        /// <summary>Closes the channel after disposal or any dispatched attempt without verified completion.</summary>
        private bool disposed, uncertain;

        /// <summary>Immutable correlation identity plus one-shot claim/publication state for a single procedure call.</summary>
        private sealed class Attempt
        {

            /// <summary>COM identity of the project authorized for this callback.</summary>
            internal object Project;

            /// <summary>Bound path, support version/signature, nonce, revision, run, descriptor, phase, and first callback fault.</summary>
            internal string Path, Version, Signature, Nonce, Revision, RunId, Module, Procedure, TestId, Phase, Fault;

            /// <summary>Enforces the sequence arm, native dispatch, one claim, and one publication.</summary>
            internal bool Started, Claimed, Published;

            /// <summary>Accepted status, message, and invariant-culture error number awaiting native completion proof.</summary>
            internal object[] Verdict;
        }

        /// <summary>Creates a sink that validates callback verdicts against the pending test attempt.</summary>
        /// <param name="sameProject">Optional COM identity comparer; defaults to the VBE project's native identity comparison.</param>
        internal VbaTestResultSink(Func<object, object, bool> sameProject = null)
        { this.sameProject = sameProject ?? VbeDebug.NativeProcedureValuesHost.SameComIdentity; }

        /// <summary>Opens one bounded callback attempt and generates its unpredictable one-use claim nonce.</summary>
        /// <param name="project">Exact live project object authorized for the call.</param>
        /// <param name="path">Canonical host path associated with the project.</param>
        /// <param name="supportVersion">Expected project-local callback module version.</param>
        /// <param name="revision">Project source revision being executed.</param>
        /// <param name="runId">Current run correlation identifier.</param>
        /// <param name="test">Discovered descriptor expected in the callback.</param>
        /// <param name="phase">One of the supported test/setup/cleanup phases.</param>
        /// <param name="supportSignature">Optional fingerprint of the generated support module.</param>
        /// <exception cref="InvalidOperationException">The sink is closed or another attempt is active.</exception>
        /// <exception cref="ArgumentException">Any identity field is absent, malformed, or exceeds its bound.</exception>
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
                active = new Attempt
                {
                    Project = project,
                    Path = path,
                    Version = supportVersion,
                    Signature = supportSignature,
                    Revision = revision,
                    RunId = runId,
                    TestId = test.Id,
                    Module = test.Module,
                    Procedure = test.Procedure,
                    Phase = phase,
                    Nonce = BitConverter.ToString(bytes).Replace("-", "")
                };
            }
        }

        /// <summary>Returns the active nonce to the project-local runtime after verifying callback-thread availability.</summary>
        /// <returns>Nonce for the active attempt.</returns>
        internal string BindRuntime()
        {
            lock (gate) { RequireCallback(); return active.Nonce; }
        }

        /// <summary>Atomically claims an attempt for the in-process runtime after nonce, version, and support-signature checks.</summary>
        /// <param name="binding">Nonce bound into the project-local runtime instance.</param>
        /// <param name="supportVersion">Runtime version being claimed.</param>
        /// <param name="supportSignature">Current generated support-module fingerprint.</param>
        /// <returns>Correlation tuple for the claimed procedure; a second claim is rejected.</returns>
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

        /// <summary>Publishes from the bound runtime instance after matching its run and descriptor identifiers.</summary>
        /// <param name="binding">Runtime-instance nonce used to prevent another loaded support module from publishing.</param>
        /// <param name="nonce">Nonce returned by the one-time request claim.</param>
        /// <param name="revision">Revision returned for the active project snapshot.</param>
        /// <param name="runId">Run identifier returned by the active claim.</param>
        /// <param name="testId">Descriptor identifier returned by the active claim.</param>
        /// <param name="status">Allowed terminal test verdict.</param>
        /// <param name="message">Bounded diagnostic text.</param>
        /// <param name="errorNumber">VBA error number associated with the verdict.</param>
        /// <returns>Whether the bound runtime instance's verdict was accepted.</returns>
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

        /// <summary>Gets whether a valid callback verdict has been published for the active attempt.</summary>
        /// <value>False until the one permitted verdict is accepted.</value>
        internal bool HasVerdict { get { lock (gate) { return active?.Published == true; } } }

        /// <summary>Gets whether any rejected callback has latched a failure on the active dispatched attempt.</summary>
        /// <value>True means the native receipt cannot be accepted as successful.</value>
        internal bool HasFault { get { lock (gate) { return active?.Fault != null; } } }

        /// <summary>Marks the attempt as dispatched, after which cancellation and replay are forbidden.</summary>
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
        /// <summary>Claims the COM-visible attempt once after matching project COM identity, path, and support version.</summary>
        /// <param name="project">Callback caller's VBIDE project object.</param>
        /// <param name="hostPath">Callback caller's host document path.</param>
        /// <param name="supportVersion">Version of the calling support module.</param>
        /// <returns>Correlation tuple consumed by the runtime; mismatches latch a fault and throw.</returns>
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
        /// <summary>Accepts one bounded verdict only after the active attempt has been claimed.</summary>
        /// <param name="nonce">Attempt nonce supplied by the claim response.</param>
        /// <param name="revision">Revision token returned by that claim.</param>
        /// <param name="status">Allowed terminal status.</param>
        /// <param name="message">Non-null message within the runtime's maximum length.</param>
        /// <param name="errorNumber">VBA error number, constrained by verdict type.</param>
        /// <returns>True when accepted; invalid or repeated publications latch a fault and return false.</returns>
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

        /// <summary>Closes the active attempt only when native completion and the claimed callback verdict are both verified.</summary>
        /// <param name="completed">Whether the host returned to the expected completed mode.</param>
        /// <param name="nativeError">Native transport diagnostic used when no callback fault is more specific.</param>
        /// <returns>The accepted verdict tuple.</returns>
        /// <exception cref="VbaTestInvocationException">Dispatch did not occur or completion evidence is missing/uncertain.</exception>
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

        /// <summary>Clears an attempt only before it is marked dispatched; native calls cannot be cancelled or replayed.</summary>
        internal void CancelUndispatched()
        {
            RequireOwner();
            lock (gate)
            {
                if (active?.Started == true) throw new InvalidOperationException("A dispatched attempt cannot be cancelled or retried.");
                active = null;
            }
        }

        /// <summary>Rejects callback access unless it is on the owner thread with a started, open, fault-free attempt.</summary>
        private void RequireCallback()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread || disposed || active == null || !active.Started || active.Fault != null)
                Reject("The callback channel is not available on this owning thread and attempt.");
        }

        /// <summary>Latches the first post-dispatch callback fault and throws to the caller.</summary>
        /// <param name="reason">Diagnostic for the rejected callback condition.</param>
        /// <returns>This method always throws and has no normal return.</returns>
        private object Reject(string reason)
        {
            if (active?.Started == true && active.Fault == null) active.Fault = reason;
            throw new InvalidOperationException(reason);
        }

        /// <summary>Throws unless the sink is accessed on the thread that created it.</summary>
        private void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("The callback channel requires its owning thread."); }

        /// <summary>Checks that a string is nonblank and does not exceed a caller-supplied character limit.</summary>
        /// <param name="value">String to validate.</param>
        /// <param name="maximum">Maximum allowed UTF-16 character count.</param>
        /// <returns>True when the string is present and within the bound.</returns>
        private static bool Bound(string value, int maximum) => !string.IsNullOrWhiteSpace(value) && value.Length <= maximum;

        /// <summary>Checks whether a phase is one of the five supported test/setup/cleanup labels.</summary>
        /// <param name="value">Phase label from the runner.</param>
        /// <returns>True only for Test, ModuleInitialize, ModuleCleanup, TestInitialize, or TestCleanup.</returns>
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
