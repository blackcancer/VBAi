using System;
using System.IO;

/// <summary>Owns short synthetic Git paths independently of a deeply nested assembly output directory.</summary>
internal sealed class GitScratchDirectory
{
    // Git receive-pack appends object quarantine and SHA paths to the fixture root.
    internal const int MaximumRootCharacters = 140;
    internal string Root { get; private set; }
    internal string Boundary { get; private set; }
    private GitScratchDirectory() { }

    internal static GitScratchDirectory Create()
    {
        return Create(AppDomain.CurrentDomain.BaseDirectory, Path.GetTempPath());
    }

    /// <summary>Uses a short repository artifact directory on the assembly volume; standalone smoke uses owned temp.</summary>
    internal static GitScratchDirectory Create(string assemblyDirectory, string temporaryDirectory)
    {
        string boundary = null;
        for (var directory = new DirectoryInfo(Path.GetFullPath(assemblyDirectory)); directory != null; directory = directory.Parent)
        {
            string marker = Path.Combine(directory.FullName, ".git");
            if (!File.Exists(marker) && !Directory.Exists(marker)) continue;
            string candidate = Path.Combine(directory.FullName, "artifacts", "gs");
            if (candidate.Length + 33 <= MaximumRootCharacters) { boundary = candidate; break; }
        }
        if (boundary == null) boundary = Path.Combine(Path.GetFullPath(temporaryDirectory), "VBAi-gs");
        string root = Path.Combine(boundary, Guid.NewGuid().ToString("N"));
        if (root.Length > MaximumRootCharacters)
            throw new PathTooLongException("The owned Git scratch root leaves insufficient space for native Git object paths.");
        RejectLinks(boundary);
        if (Directory.Exists(root) || File.Exists(root)) throw new IOException("The new owned Git scratch path already exists.");
        Directory.CreateDirectory(root);
        return new GitScratchDirectory { Root = root, Boundary = Path.GetFullPath(boundary) };
    }

    /// <summary>Requires the exact original GUID child before a fixture considers deleting its contents.</summary>
    internal void ValidateCleanupRoot(string requestedRoot)
    {
        string full = Path.GetFullPath(requestedRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        Guid identity;
        if (!string.Equals(full, Root, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetDirectoryName(full), Boundary, StringComparison.OrdinalIgnoreCase) ||
            !Guid.TryParseExact(Path.GetFileName(full), "N", out identity))
            throw new InvalidOperationException("Git scratch cleanup requires its exact original owned GUID directory.");
        RejectLinks(full);
    }

    /// <summary>Refuses existing reparse ancestors before creation or deletion can traverse an unintended tree.</summary>
    private static void RejectLinks(string path)
    {
        for (var directory = new DirectoryInfo(Path.GetFullPath(path)); directory != null; directory = directory.Parent)
            if ((Directory.Exists(directory.FullName) || File.Exists(directory.FullName)) &&
                (File.GetAttributes(directory.FullName) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("Git scratch ownership refuses filesystem links.");
    }
}
