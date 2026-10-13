using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;

namespace VBAi
{
    /// <summary>Attempts each independent shutdown resource once and preserves all unexpected cleanup failures.</summary>
    internal sealed class AddInShutdownCleanup
    {
        /// <summary>Optional observer; disabled diagnostics do not change cleanup semantics.</summary>
        private readonly AddInShutdownDiagnostic diagnostic;
        /// <summary>Original exceptions, retained in attempt order without replacing them with log failures.</summary>
        private readonly List<Exception> errors = new List<Exception>();

        /// <summary>Creates a single shutdown sequence's failure collector.</summary>
        internal AddInShutdownCleanup(AddInShutdownDiagnostic diagnostic) { this.diagnostic = diagnostic; }

        /// <summary>Attempts a step once; the supplied policy may admit an already released native site.</summary>
        internal void Step(string stage, Action action, Func<Exception, bool> ignore = null)
        {
            diagnostic?.Enter(stage);
            try { action(); diagnostic?.Returned(stage); }
            catch (Exception error)
            {
                diagnostic?.CaughtFault(stage, error);
                try { if (ignore == null || !ignore(error)) errors.Add(error); }
                catch (Exception policyError) { errors.Add(error); errors.Add(policyError); }
            }
        }

        /// <summary>Clears a field after its single disposal attempt even when cleanup fails, avoiding a later implicit retry.</summary>
        internal void Resource<T>(ref T resource, string name, Action<T> dispose, string operation = "Dispose", Func<Exception, bool> ignore = null)
            where T : class
        {
            try { if (resource != null) { var owned = resource; Step(name + "." + operation, () => dispose(owned), ignore); } }
            finally { resource = null; diagnostic?.ManagedReferenceCleared(name); }
        }

        /// <summary>Re-throws a single original failure, or reports every distinct failure after all resources were attempted.</summary>
        internal void ThrowIfFailed()
        {
            if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
            if (errors.Count > 1) throw new AggregateException("Multiple VBAi shutdown resources failed to close.", errors);
        }
    }
}
