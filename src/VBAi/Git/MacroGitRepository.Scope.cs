using System;
using System.IO;

namespace VBAi
{
    internal sealed partial class MacroGitRepository
    {
        /// <summary>Finds one existing document binding without moving or rewriting persisted cache keys.</summary>
        internal static string ResolveScopeDirectory(string nativeScope)
        {
            return ResolveScopeDirectory(nativeScope, ScopeDirectory, File.GetAttributes);
        }

        /// <summary>Checks the native and former uppercase keys through read-only, per-call boundaries.</summary>
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
