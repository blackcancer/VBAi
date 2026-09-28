using System;
using System.IO;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Finds the Codex CLI independently of the host application's PATH.</summary>
    internal static class CodexCliLocator
    {
        internal static string Resolve(Func<string, bool> fileExists,
            Func<string, string[]> getDirectories, Func<string, DateTime> getLastWriteTimeUtc)
        {
            string configured = Environment.GetEnvironmentVariable("CODEXVBE_CODEX_CLI");
            if (!string.IsNullOrWhiteSpace(configured)) return configured;

            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string legacy = Path.Combine(local, "Programs", "OpenAI", "Codex", "bin", "codex.exe");
            if (fileExists(legacy)) return legacy;

            string desktopBin = Path.Combine(local, "OpenAI", "Codex", "bin");
            string[] directories;
            try { directories = getDirectories(desktopBin); }
            catch (IOException) { directories = new string[0]; }
            catch (UnauthorizedAccessException) { directories = new string[0]; }

            string versioned = directories.Select(directory => Path.Combine(directory, "codex.exe"))
                .Where(fileExists)
                .OrderByDescending(getLastWriteTimeUtc)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
            return versioned ?? "codex.exe";
        }
    }
}
