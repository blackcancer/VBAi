using System;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal static class UpdatePaths
    {
        internal static string Root { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "Updates");
        internal static string Hash(string path) { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "").ToLowerInvariant(); }
        internal static bool IsHash(string value) => Regex.IsMatch(value ?? "", "^[a-f0-9]{64}$");
        internal static string AssetPath(string root, string hash, string fileName)
        {
            if (!IsHash(hash) || (fileName != "VBAi-Setup-win-x64.exe" && fileName != "VBAi-Setup-win-x64.msi")) throw new InvalidDataException("Invalid update asset.");
            return Path.Combine(Path.GetFullPath(root), hash, fileName);
        }
        internal static bool SamePath(string left, string right) => string.Equals(Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar), Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
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
