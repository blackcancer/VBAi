using System;
using System.IO;

namespace VBAi
{

    /// <summary>Resolves an existing per-document Git binding while preserving exact and legacy cache keys.</summary>
    internal sealed partial class MacroGitRepository
    {

        /// <summary>Finds one existing document binding without moving or rewriting persisted cache keys.</summary>
        /// <param name="nativeScope">Exact host-provided document scope used as the preferred cache key.</param>
        /// <returns>Bound cache directory when one regular binding exists, otherwise the exact native-key directory.</returns>
        internal static string ResolveScopeDirectory(string nativeScope)
        {
            return ResolveScopeDirectory(nativeScope, ScopeDirectory, File.GetAttributes);
        }

        /// <summary>Checks the native and former uppercase keys through read-only, per-call boundaries.</summary>
        /// <param name="nativeScope">Exact document scope and source for the uppercase legacy key.</param>
        /// <param name="directoryForScope">Maps a scope key to its cache directory.</param>
        /// <param name="readAttributes">Reads binding attributes; only definite file/directory absence permits fallback.</param>
        /// <returns>The uniquely bound cache, or the exact-key cache if neither key has a binding.</returns>
        /// <exception cref="IOException">Both keys are bound, or a binding path is a directory or reparse point.</exception>
        internal static string ResolveScopeDirectory(string nativeScope, Func<string, string> directoryForScope,
            Func<string, FileAttributes> readAttributes)
        {
            string exact = directoryForScope(nativeScope);
            string legacy = directoryForScope(nativeScope.ToUpperInvariant());
            bool exactBound = HasRegularBinding(exact, readAttributes);
            if (string.Equals(exact, legacy, StringComparison.OrdinalIgnoreCase)) return exact;
            bool legacyBound = HasRegularBinding(legacy, readAttributes);
            if (exactBound && legacyBound)
                throw new IOException("This document has bindings in both its native and legacy Git caches. Resolve the ambiguity before opening Git; neither cache was changed.");
            return legacyBound ? legacy : exact;
        }

        /// <summary>Only definite absence permits another cache; unreadable, directory or link entries refuse lookup.</summary>
        /// <param name="directory">Candidate cache directory whose binding.json entry is inspected.</param>
        /// <param name="readAttributes">Filesystem attribute reader used to classify the entry.</param>
        /// <returns><see langword="true"/> for an existing regular binding file; <see langword="false"/> only for definite absence.</returns>
        /// <exception cref="IOException">The binding exists as a directory or reparse point.</exception>
        private static bool HasRegularBinding(string directory, Func<string, FileAttributes> readAttributes)
        {
            FileAttributes attributes;
            try { attributes = readAttributes(Path.Combine(directory, "binding.json")); }
            catch (FileNotFoundException) { return false; }
            catch (DirectoryNotFoundException) { return false; }
            if ((attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
                throw new IOException("The Git binding is not a regular file; the cache cannot be selected automatically.");
            return true;
        }
    }
}
