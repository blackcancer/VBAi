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
        [DispId(1)] object Request(object project, string hostPath, string supportVersion);
        /// <summary>Publishes one bounded verdict for the claimed call; native completion remains mandatory.</summary>
        [DispId(2)] bool Publish(string nonce, string revision, string status, string message, int errorNumber);
    }

    /// <summary>STA-owned bounded callback state, published through the existing VBIDE AddIn.Object property.</summary>
    [ComVisible(true), Guid("744E3C39-BF87-49EA-926A-8B6D07F1A0DF"), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IVbaTestResultSink))]
    public sealed class VbaTestResultSink : IVbaTestResultSink, IDisposable
    {
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private readonly Func<object, object, bool> sameProject;
        private readonly object gate = new object();
        private Attempt active;
        private bool disposed, uncertain;

        private sealed class Attempt
        {
            internal object Project;
            internal string Path, Version, Signature, Nonce, Revision, RunId, Module, Procedure, TestId, Phase, Fault;
            internal bool Started, Claimed, Published;
            internal object[] Verdict;
        }

        internal VbaTestResultSink(Func<object, object, bool> sameProject = null)
        { this.sameProject = sameProject ?? VbeDebug.NativeProcedureValuesHost.SameComIdentity; }

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

        internal string BindRuntime()
        {
            lock (gate) { RequireCallback(); return active.Nonce; }
        }

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

        internal bool HasVerdict { get { lock (gate) { return active?.Published == true; } } }
        internal bool HasFault { get { lock (gate) { return active?.Fault != null; } } }

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

        internal void CancelUndispatched()
        {
            RequireOwner();
            lock (gate)
            {
                if (active?.Started == true) throw new InvalidOperationException("A dispatched attempt cannot be cancelled or retried.");
                active = null;
            }
        }

        private void RequireCallback()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread || disposed || active == null || !active.Started || active.Fault != null)
                Reject("The callback channel is not available on this owning thread and attempt.");
        }

        private object Reject(string reason)
        {
            if (active?.Started == true && active.Fault == null) active.Fault = reason;
            throw new InvalidOperationException(reason);
        }

        private void RequireOwner()
        { if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("The callback channel requires its owning thread."); }

        private static bool Bound(string value, int maximum) => !string.IsNullOrWhiteSpace(value) && value.Length <= maximum;
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
