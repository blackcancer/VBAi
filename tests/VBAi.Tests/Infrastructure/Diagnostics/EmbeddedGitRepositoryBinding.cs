using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    /// <summary>Retains the selected bare repository's binding and immutable filesystem identity before a Save is considered.</summary>
    internal sealed class EmbeddedGitRepositoryBinding
    {
        internal string Cache, Remote, Branch, RepositoryPath;
        internal byte[] BindingBytes, ConfigBytes, HeadBytes;

        internal static string SelectedPath(string cache, string remote, string branch)
        {
            using (var hash = SHA256.Create())
                return Path.Combine(Path.GetFullPath(cache), BitConverter.ToString(hash.ComputeHash(
                    Encoding.UTF8.GetBytes(remote + "\n" + branch))).Replace("-", "") + ".git");
        }

        /// <summary>Reads only existing files; the caller must separately attest the selected commit through production Resolve.</summary>
        internal static EmbeddedGitRepositoryBinding Capture(string cache, string remote, string branch)
        {
            var proof = new EmbeddedGitRepositoryBinding { Cache = Path.GetFullPath(cache), Remote = remote, Branch = branch,
                RepositoryPath = SelectedPath(cache, remote, branch) };
            RequireDirectory(proof.Cache); RequireDirectory(proof.RepositoryPath);
            proof.BindingBytes = ReadRegular(Path.Combine(proof.Cache, "binding.json"));
            RequireBinding(proof.BindingBytes, remote, branch);
            proof.ConfigBytes = ReadRegular(Path.Combine(proof.RepositoryPath, "config"));
            proof.HeadBytes = ReadRegular(Path.Combine(proof.RepositoryPath, "HEAD"));
            return proof;
        }

        /// <summary>Rejects changed/missing repository identity and observes recovery exclusively in the selected child.</summary>
        internal MacroGitRepository RequireReadyForSave(string cache, string remote, string branch,
            Func<string, FileAttributes> recoveryAttributes = null)
        {
            string selected = SelectedPath(cache, remote, branch);
            if (!string.Equals(Cache, Path.GetFullPath(cache), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(RepositoryPath, selected, StringComparison.OrdinalIgnoreCase) ||
                Remote != remote || Branch != branch)
                throw new InvalidOperationException("The selected repository path/binding changed; Save is refused.");
            RequireDirectory(Cache); RequireDirectory(selected);
            byte[] binding = ReadRegular(Path.Combine(Cache, "binding.json"));
            RequireBinding(binding, remote, branch);
            if (!BindingBytes.SequenceEqual(binding) || !ConfigBytes.SequenceEqual(ReadRegular(Path.Combine(selected, "config"))) ||
                !HeadBytes.SequenceEqual(ReadRegular(Path.Combine(selected, "HEAD"))))
                throw new InvalidOperationException("The bound repository's binding/config/HEAD bytes changed; Save is refused.");
            var repository = new MacroGitRepository(selected, branch);
            if (recoveryAttributes != null) repository.RecoveryAttributes = recoveryAttributes;
            if (repository.RecoveryPending)
                throw new InvalidOperationException("Pending recovery in the selected repository prohibits post-import Save.");
            return repository;
        }

        private static void RequireDirectory(string path)
        {
            FileAttributes attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.Directory) == 0 || (attributes & FileAttributes.ReparsePoint) != 0)
                throw new IOException("The selected repository must exist as an ordinary directory.");
        }

        private static byte[] ReadRegular(string path)
        {
            if ((File.GetAttributes(path) & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
                throw new IOException("Repository identity evidence must be an existing ordinary file.");
            return File.ReadAllBytes(path);
        }

        private static void RequireBinding(byte[] bytes, string remote, string branch)
        {
            var binding = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(new UTF8Encoding(false, true).GetString(bytes));
            if (binding == null || !binding.TryGetValue("Remote", out var actualRemote) || actualRemote != remote ||
                !binding.TryGetValue("Branch", out var actualBranch) || actualBranch != branch)
                throw new InvalidOperationException("The document binding does not select the validated remote and branch.");
        }
    }
}
