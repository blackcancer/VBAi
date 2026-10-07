using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdatePathsTests
    {
        [TestMethod]
        public void CachePathsAreFixedAndAtomicPreferencesReplacementLeavesNoTemporaryFiles()
        {
            using (var scope = new UpdateScope())
            {
                string hash = new string('a', 64);
                Assert.AreEqual(Path.Combine(scope.Root, hash, "VBAi-Setup-win-x64.exe"), UpdatePaths.AssetPath(scope.Root, hash, "VBAi-Setup-win-x64.exe"));
                Assert.ThrowsException<InvalidDataException>(() => UpdatePaths.AssetPath(scope.Root, "../bad", "VBAi-Setup-win-x64.exe"));
                Assert.ThrowsException<InvalidDataException>(() => UpdatePaths.AssetPath(scope.Root, hash, "../installer.exe"));
                string path = Path.Combine(scope.Root, "state.json"); UpdatePaths.WriteAtomic(path, "first"); UpdatePaths.WriteAtomic(path, "second");
                Assert.AreEqual("second", File.ReadAllText(path)); Assert.AreEqual(0, Directory.GetFiles(scope.Root, "*.tmp").Length);
                Assert.IsTrue(UpdatePaths.SamePath(path, path.ToUpperInvariant()));
            }
        }
        [TestMethod]
        public void LockedOwnedReplacementPreservesDestinationAndCleansItsTemporaryPayload()
        {
            using (var scope = new UpdateScope())
            {
                Assert.IsFalse(UpdatePaths.IsHash(null)); Assert.IsFalse(UpdatePaths.IsHash(new string('A', 64)));
                Assert.AreEqual(Path.Combine(scope.Root, new string('b', 64), "VBAi-Setup-win-x64.msi"), UpdatePaths.AssetPath(scope.Root, new string('b', 64), "VBAi-Setup-win-x64.msi"));
                string path = Path.Combine(scope.Root, "locked.json"); File.WriteAllText(path, "original");
                using (var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                    Assert.ThrowsException<IOException>(() => UpdatePaths.WriteAtomic(path, "replacement"));
                Assert.AreEqual("original", File.ReadAllText(path)); Assert.AreEqual(0, Directory.GetFiles(scope.Root, "*.tmp").Length);
            }
        }
    }
}
