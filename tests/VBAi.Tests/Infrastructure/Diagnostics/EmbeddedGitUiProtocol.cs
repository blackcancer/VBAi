using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    /// <summary>Tracks unique UI delivery separately from a verified terminal operation.</summary>
    internal sealed class EmbeddedGitUiProtocol
    {
        private readonly HashSet<string> emitted = new HashSet<string>(StringComparer.Ordinal);
        private readonly Action<object> evidence;
        private string pending;
        internal bool Uncertain { get; private set; }
        internal bool CanClose => pending == null && !Uncertain;
        internal EmbeddedGitUiProtocol(Action<object> evidence) { this.evidence = evidence ?? throw new ArgumentNullException(nameof(evidence)); }

        internal void EmitOnce(string action, Action emit)
        {
            if (Uncertain || pending != null || emitted.Contains(action))
                throw new InvalidOperationException("UI delivery is pending, uncertain or already emitted; do not replay.");
            evidence(new { Phase = "Intent", Action = action, Utc = DateTime.UtcNow.ToString("o") });
            emitted.Add(action); pending = action;
            try { emit(); evidence(new { Phase = "DeliveryReturned", Action = action }); }
            catch { Uncertain = true; throw; }
        }

        internal void Terminal(string action, bool proof, bool success)
        {
            if (pending != action || !proof || Uncertain)
                throw new InvalidOperationException("Missing exact terminal proof; idle or a returned Invoke alone is insufficient.");
            evidence(new { Phase = "Terminal", Action = action, Success = success, Utc = DateTime.UtcNow.ToString("o") });
            pending = null;
        }

        internal void MarkUncertain(string reason)
        {
            Uncertain = true;
            evidence(new { Phase = "Uncertain", Action = pending, Reason = reason });
        }

        internal static void RequireOwner(int expectedPid, uint expectedTid, long expectedHandle,
            int observedPid, uint observedTid, long observedHandle)
        {
            if (expectedPid <= 0 || expectedTid == 0 || expectedHandle == 0 || expectedPid != observedPid ||
                expectedTid != observedTid || expectedHandle != observedHandle)
                throw new InvalidOperationException("The exact owned UI identity is absent or changed; no action is allowed.");
        }

        internal static bool IsTerminal(bool idle, bool busyObserved, string before, string after)
            => idle && (busyObserved || !string.Equals(before, after, StringComparison.Ordinal));

        internal static string[] ReadSession(string description)
        {
            string[] fields = (description ?? "").Split('/'); Guid parsed;
            if (fields.Length != 6 || fields[0] != "VBAi.GitSession" || fields[1] != "1" ||
                !Guid.TryParseExact(fields[2], "N", out parsed) ||
                (fields[3] != "none" && !Guid.TryParseExact(fields[3], "N", out parsed)))
                throw new InvalidOperationException("Exact Git session correlation is absent.");
            if (fields[3] == "none" && (fields[4] != "none" || fields[5] != "Idle"))
                throw new InvalidOperationException("An idle Git session cannot claim an operation.");
            return fields;
        }

        internal static bool HandoffTerminal(string expectedSession, string previousOperation, string expectedAction,
            string description, ref string observedOperation, out bool success)
        {
            string[] state = ReadSession(description); success = false;
            if (state[2] != expectedSession) throw new InvalidOperationException("A different Git session appeared during handoff.");
            if (state[3] == previousOperation || state[3] == "none") return false;
            if (state[4] != expectedAction || observedOperation != null && observedOperation != state[3])
                throw new InvalidOperationException("The one delivered Git operation changed during handoff.");
            observedOperation = state[3];
            if (state[5] == "Succeeded") { success = true; return true; }
            if (state[5] == "Failed") return true;
            if (state[5] != "AwaitingModalReturn" && state[5] != "Executing" && state[5] != "Refused")
                throw new InvalidOperationException("Unknown Git handoff phase.");
            return false;
        }

        internal static bool MustRetainOwner(bool nativePending, bool menuEmitted, bool modalClosed)
            => nativePending || menuEmitted && !modalClosed;

        internal static Exception PreserveFailures(Exception primary, params Exception[] observed)
        {
            var failures = new List<Exception> { primary };
            foreach (var error in observed)
            {
                if (error == null) continue;
                var candidates = error is AggregateException aggregate ? aggregate.Flatten().InnerExceptions :
                    (IEnumerable<Exception>)new[] { error };
                foreach (var candidate in candidates)
                    if (!failures.Exists(known => ReferenceEquals(known, candidate) ||
                        known is AggregateException knownAggregate && knownAggregate.Flatten().InnerExceptions.Contains(candidate)))
                        failures.Add(candidate);
            }
            return failures.Count == 1 ? primary : new AggregateException("Original scenario and separately observed owner/UI/cleanup errors are preserved.", failures);
        }
    }
}
