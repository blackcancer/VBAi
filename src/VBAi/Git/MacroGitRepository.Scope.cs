using System;
using System.IO;

namespace VBAi
{

    /// <summary>Owns the macro git repository state and operations.</summary>
    internal sealed partial class MacroGitRepository
    {

        /// <summary>Finds one existing document binding without moving or rewriting persisted cache keys.</summary>
        /// <param name="nativeScope">Text that supplies the native scope value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for resolve scope directory on macro git repository.</returns>
        internal static string ResolveScopeDirectory(string nativeScope)
        {
            return ResolveScopeDirectory(nativeScope, ScopeDirectory, File.GetAttributes);
        }

        /// <summary>Checks the native and former uppercase keys through read-only, per-call boundaries.</summary>
        /// <param name="nativeScope">Text that supplies the native scope value. Use the format required by the calling operation.</param>
        /// <param name="directoryForScope">func&lt;string, string&gt; that supplies the directory for scope for this operation.</param>
        /// <param name="readAttributes">func&lt;string, file attributes&gt; that supplies the read attributes for this operation.</param>
        /// <returns>Text produced by the operation for resolve scope directory on macro git repository.</returns>
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
        /// <param name="directory">Text that supplies the directory value. Use the format required by the calling operation.</param>
        /// <param name="readAttributes">func&lt;string, file attributes&gt; that supplies the read attributes for this operation.</param>
        /// <returns>Boolean indicating the result of the check for has regular binding on macro git repository.</returns>
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
