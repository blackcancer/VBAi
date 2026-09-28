using System;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Validates update cache paths, computes asset hashes, and performs atomic text writes.</summary>
internal static class UpdatePaths
    {
        /// <summary>Gets or sets the per-user root that contains update preferences, downloads, and host leases.</summary>
        /// <value>Update data directory below LocalApplicationData.</value>
internal static string Root { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "Updates");
        /// <summary>Computes a lowercase SHA-256 digest for a file.</summary><param name="path">File to hash.</param><returns>64-character hexadecimal digest.</returns>
internal static string Hash(string path) { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "").ToLowerInvariant(); }
        /// <summary>Checks whether a string is a lowercase 64-character hexadecimal hash.</summary><param name="value">Candidate hash.</param><returns>Whether it matches the accepted hash format.</returns>
internal static bool IsHash(string value) => Regex.IsMatch(value ?? "", "^[a-f0-9]{64}$");
        /// <summary>Builds the cache path for one supported installer asset after validating its hash and file name.</summary>
        /// <param name="root">Update cache root.</param><param name="hash">Expected lowercase SHA-256 digest.</param><param name="fileName">Supported MSI or setup EXE name.</param>
        /// <returns>Full asset path under the hash directory.</returns><exception cref="InvalidDataException">The digest or file name is not supported.</exception>
internal static string AssetPath(string root, string hash, string fileName)
        {
            if (!IsHash(hash) || (fileName != "VBAi-Setup-win-x64.exe" && fileName != "VBAi-Setup-win-x64.msi")) throw new InvalidDataException("Invalid update asset.");
            return Path.Combine(Path.GetFullPath(root), hash, fileName);
        }
        /// <summary>Compares two paths after resolving them to full paths and trimming trailing separators.</summary>
        /// <param name="left">First path.</param><param name="right">Second path.</param><returns>Whether they refer to the same path, ignoring case.</returns>
internal static bool SamePath(string left, string right) => string.Equals(Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar), Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
        /// <summary>Writes UTF-8 text through a temporary file and atomically replaces or creates the destination.</summary>
        /// <param name="path">Destination file path.</param><param name="content">Text to write without a UTF-8 BOM.</param>
internal static void WriteAtomic(string path, string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, content, new System.Text.UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
