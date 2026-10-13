using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;

namespace VBAi
{
    /// <summary>Retains the original connection failure while attempting diagnostics and rollback independently.</summary>
    internal static class AddInStartupFailure
    {
        /// <summary>Attempts each failure handler once, then rethrows the original error or an ordered aggregate.</summary>
        /// <param name="original">The failure that interrupted initialization.</param>
        /// <param name="log">Writes the initialization failure to the local log.</param>
        /// <param name="report">Offers the failure to the reporter, if one was created.</param>
        /// <param name="cleanup">Releases resources allocated before initialization failed.</param>
        internal static void Rethrow(Exception original, Action log, Action report, Action cleanup)
        {
            if (original == null) throw new ArgumentNullException(nameof(original));
            if (log == null) throw new ArgumentNullException(nameof(log));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (cleanup == null) throw new ArgumentNullException(nameof(cleanup));
            var failures = new List<Exception> { original };
            foreach (var action in new[] { log, report, cleanup })
            {
                try { action(); }
                catch (Exception error)
                {
                    if (!failures.Exists(previous => ReferenceEquals(previous, error))) failures.Add(error);
                }
            }
            if (failures.Count == 1) ExceptionDispatchInfo.Capture(original).Throw();
            throw new AggregateException("VBAi initialization and failure recovery failed.", failures);
        }
    }
}
