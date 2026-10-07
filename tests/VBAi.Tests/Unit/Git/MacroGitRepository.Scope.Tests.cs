using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitRepositoryTests
    {
        [DataTestMethod, DataRow("none"), DataRow("exact"), DataRow("legacy"), DataRow("both")]
        public void ScopeLookupSelectsOnlyAnUnambiguousExistingBindingWithoutChangingStorage(string state)
        {
            InRecoveryMarker(repository =>
            {
                string root = Path.GetDirectoryName(repository.RecoveryFile);
                string scope = Path.Combine(root, "Développement", "MixedCaseÉté.xlsm");
                Func<string, string> cache = key => Path.Combine(root, Path.GetFileName(MacroGitRepository.ScopeDirectory(key)));
                string exact = cache(scope), legacy = cache(scope.ToUpperInvariant());
                Assert.AreNotEqual(exact, legacy);
                foreach (string directory in new[] { exact, legacy })
                {
                    Directory.CreateDirectory(directory);
                    File.WriteAllText(Path.Combine(directory, "sentinel.txt"), "retained refs/hash/transport " + directory);
                    if (state == "both" || state == (directory == exact ? "exact" : "legacy"))
                        File.WriteAllText(Path.Combine(directory, "binding.json"), "synthetic binding retained byte-exact");
                }
                var before = Directory.GetFiles(root, "*", SearchOption.AllDirectories).ToDictionary(path => path, File.ReadAllBytes);
                if (state == "both") StringAssert.Contains(Assert.ThrowsException<IOException>(() =>
                    MacroGitRepository.ResolveScopeDirectory(scope, cache, File.GetAttributes)).Message, "both");
                else Assert.AreEqual(state == "legacy" ? legacy : exact,
                    MacroGitRepository.ResolveScopeDirectory(scope, cache, File.GetAttributes));
                CollectionAssert.AreEquivalent(before.Keys.ToArray(), Directory.GetFiles(root, "*", SearchOption.AllDirectories));
                foreach (var file in before) CollectionAssert.AreEqual(file.Value, File.ReadAllBytes(file.Key));
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ScopeLookupObservesAnAlreadyUppercaseKeyOnlyOnceAndTreatsOnlyMissingAsAbsence(bool bound)
        {
            string scope = @"C:\DÉVELOPPEMENT\ÉTÉ.XLSM"; int observations = 0;
            string expected = MacroGitRepository.ScopeDirectory(scope);
            Assert.AreEqual(expected, MacroGitRepository.ResolveScopeDirectory(scope, MacroGitRepository.ScopeDirectory, path =>
            {
                observations++; Assert.AreEqual(Path.Combine(expected, "binding.json"), path);
                if (!bound) throw new DirectoryNotFoundException("known absent parent");
                return FileAttributes.Archive;
            }));
            Assert.AreEqual(1, observations);
        }

        [DataTestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void BothMissingExceptionTypesPermitNativeKeyWithoutWritingASecondCache(bool exactParent, bool legacyParent)
        {
            string scope = @"C:\MixedCase\ClasseurÉté.xlsm"; int observations = 0;
            Assert.AreEqual(MacroGitRepository.ScopeDirectory(scope), MacroGitRepository.ResolveScopeDirectory(scope,
                MacroGitRepository.ScopeDirectory, path =>
                {
                    bool parent = ++observations == 1 ? exactParent : legacyParent;
                    if (parent) throw new DirectoryNotFoundException("known absent parent");
                    throw new FileNotFoundException("known absent binding");
                }));
            Assert.AreEqual(2, observations);
        }

        [DataTestMethod, DataRow(false, "access"), DataRow(false, "io"), DataRow(true, "access"), DataRow(true, "io")]
        public void UnreadableBindingCannotFallbackToAnotherKeyEvenIfTheOtherBindingExists(bool legacyFails, string kind)
        {
            string scope = @"C:\MixedCase\ClasseurÉté.xlsm"; int observations = 0;
            Exception original = kind == "access" ? (Exception)new UnauthorizedAccessException("synthetic binding metadata failure") :
                new IOException("synthetic binding I/O failure");
            var error = RecoveryFailure(() => MacroGitRepository.ResolveScopeDirectory(scope, MacroGitRepository.ScopeDirectory, path =>
            {
                if (++observations == (legacyFails ? 2 : 1)) throw original;
                return FileAttributes.Normal;
            }));
            Assert.AreSame(original, error); Assert.AreEqual(legacyFails ? 2 : 1, observations);
        }

        [DataTestMethod, DataRow(false, FileAttributes.Directory), DataRow(true, FileAttributes.Directory)]
        [DataRow(false, FileAttributes.ReparsePoint), DataRow(true, FileAttributes.ReparsePoint)]
        [DataRow(false, FileAttributes.Directory | FileAttributes.ReparsePoint), DataRow(true, FileAttributes.Directory | FileAttributes.ReparsePoint)]
        public void InvalidBindingEntryRefusesLookupBeforeEitherCacheIsChosen(bool legacyInvalid, FileAttributes invalid)
        {
            int observations = 0;
            Assert.ThrowsException<IOException>(() => MacroGitRepository.ResolveScopeDirectory(@"C:\MixedCase\Été.xlsm",
                MacroGitRepository.ScopeDirectory, path => ++observations == (legacyInvalid ? 2 : 1) ? invalid : FileAttributes.Normal));
            Assert.AreEqual(legacyInvalid ? 2 : 1, observations);
        }

        [TestMethod]
        public void GitWindowAndLlmOperationsDefaultToTheSameCompatibleScopeResolver()
        {
            Assert.AreEqual(nameof(MacroGitRepository.ResolveScopeDirectory), GitWindow.CacheDirectory.Method.Name);
            Assert.AreEqual(GitWindow.CacheDirectory.Method, MacroGitOperations.CacheDirectory.Method);
            Assert.AreEqual(typeof(MacroGitRepository), GitWindow.CacheDirectory.Method.DeclaringType);
            string scope = @"C:\Développement\ClasseurÉté.xlsm";
            Assert.AreNotEqual(MacroGitRepository.ScopeDirectory(scope), MacroGitRepository.ScopeDirectory(scope.ToUpperInvariant()),
                "Persisted raw hashes must remain distinct; compatibility is a bounded lookup, not normalization.");
        }
    }
}
