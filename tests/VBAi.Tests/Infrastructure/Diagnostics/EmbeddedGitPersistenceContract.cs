using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    /// <summary>Requires observed normal lifecycle before advancing a synthetic imported workbook to another host.</summary>
    internal static class EmbeddedGitPersistenceContract
    {
        internal static void RequireNormalExit(IDictionary<string, object> shutdown)
        {
            if (shutdown == null || !shutdown.TryGetValue("Exited", out var exited) || !(exited is bool) || !(bool)exited ||
                !shutdown.TryGetValue("ForcedTermination", out var forced) || !(forced is bool) || (bool)forced ||
                !shutdown.TryGetValue("ExitCode", out var code) || !(code is int) || (int)code != 0)
                throw new InvalidOperationException("A terminal, observed, unforced zero exit is required before fresh-process acceptance.");
        }

        internal static void RequireFreshIdentity(string oldRoot, string oldStart, int oldPid, string freshRoot, string freshStart, int freshPid)
        {
            if (string.IsNullOrWhiteSpace(oldRoot) || string.IsNullOrWhiteSpace(freshRoot) ||
                string.Equals(oldRoot, freshRoot, StringComparison.OrdinalIgnoreCase) || oldPid <= 0 || freshPid <= 0 ||
                !DateTimeOffset.TryParse(oldStart, out var before) || !DateTimeOffset.TryParse(freshStart, out var after) || after <= before)
                throw new InvalidOperationException("The reopened workbook requires a new owned root and later verified process start; PID inequality alone is insufficient.");
        }

        internal static void RequireDiskHash(string expected, string actual)
        {
            if (string.IsNullOrEmpty(expected) || expected.Length != 64 || !string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The immutable post-import saved workbook changed during read-only reopen or normal shutdown.");
        }
    }
}
