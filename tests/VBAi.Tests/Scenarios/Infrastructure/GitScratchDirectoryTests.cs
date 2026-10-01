using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests
{
    /// <summary>Checks owned scratch selection and deletion boundaries without Git, hosts or UI.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class GitScratchDirectoryTests
    {
        [TestMethod]
        public void DeepAssemblyOutputUsesShortOwnedRepositoryPathOnItsVolume()
        {
            using (var scope = new ScratchScope())
            {
                Directory.CreateDirectory(Path.Combine(scope.Container, ".git"));
                string output = Path.Combine(scope.Container, new string('d', 110), "Debug", "net48");
                var scratch = scope.Create(output, scope.Container);
                Assert.AreEqual(Path.Combine(scope.Container, "artifacts", "gs"), scratch.Boundary);
                Assert.AreEqual(Path.GetPathRoot(output), Path.GetPathRoot(scratch.Root));
                Assert.IsTrue(scratch.Root.Length <= GitScratchDirectory.MaximumRootCharacters);
                Guid identity; Assert.IsTrue(Guid.TryParseExact(Path.GetFileName(scratch.Root), "N", out identity));
                scratch.ValidateCleanupRoot(scratch.Root);
            }
        }

        [TestMethod]
        public void DeepWorktreeMarkerCannotOverrideAShortContainingRepository()
        {
            using (var scope = new ScratchScope())
            {
                Directory.CreateDirectory(Path.Combine(scope.Container, ".git"));
                string deep = Path.Combine(scope.Container, new string('w', 70));
                Directory.CreateDirectory(deep); File.WriteAllText(Path.Combine(deep, ".git"), "synthetic worktree marker");
                var scratch = scope.Create(Path.Combine(deep, "out"), scope.Container);
                Assert.AreEqual(Path.Combine(scope.Container, "artifacts", "gs"), scratch.Boundary);
            }
        }

        [TestMethod]
        public void StandaloneOutputUsesAnOwnedTemporaryGuidDirectory()
        {
            using (var scope = new ScratchScope())
            {
                var scratch = scope.Create(Path.Combine(scope.Container, "standalone", "out"), scope.Container);
                Assert.AreEqual(Path.Combine(scope.Container, "VBAi-gs"), scratch.Boundary);
                Assert.IsTrue(Directory.Exists(scratch.Root));
            }
        }

        [TestMethod]
        public void CleanupRejectsSiblingParentAndTraversalWithoutDeletingEvidence()
        {
            using (var scope = new ScratchScope())
            {
                var scratch = scope.Create(scope.Container, scope.Container);
                string marker = Path.Combine(scratch.Root, "retained.txt"); File.WriteAllText(marker, "owned evidence");
                Assert.ThrowsException<InvalidOperationException>(() => scratch.ValidateCleanupRoot(scratch.Boundary));
                Assert.ThrowsException<InvalidOperationException>(() => scratch.ValidateCleanupRoot(Path.Combine(scratch.Boundary, Guid.NewGuid().ToString("N"))));
                Assert.ThrowsException<InvalidOperationException>(() => scratch.ValidateCleanupRoot(Path.Combine(scratch.Root, "..", "another-root")));
                Assert.AreEqual("owned evidence", File.ReadAllText(marker));
            }
        }

        [TestMethod]
        public void ExcessiveScratchPathRefusesBeforeCreatingTheTarget()
        {
            using (var scope = new ScratchScope())
            {
                string excessive = Path.Combine(scope.Container, new string('x', 90));
                Assert.ThrowsException<PathTooLongException>(() => GitScratchDirectory.Create(scope.Container, excessive));
                Assert.IsFalse(Directory.Exists(excessive));
            }
        }

        [TestMethod]
        public void NewScratchRetainsEarlierOwnedEvidenceAndNeverReusesItsDirectory()
        {
            using (var scope = new ScratchScope())
            {
                var first = scope.Create(scope.Container, scope.Container);
                string marker = Path.Combine(first.Root, "failure.bin"); File.WriteAllBytes(marker, new byte[] { 0, 42, 255 });
                var second = scope.Create(scope.Container, scope.Container);
                Assert.AreNotEqual(first.Root, second.Root);
                CollectionAssert.AreEqual(new byte[] { 0, 42, 255 }, File.ReadAllBytes(marker));
            }
        }

        private sealed class ScratchScope : IDisposable
        {
            internal readonly string Container = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            private readonly List<GitScratchDirectory> scratches = new List<GitScratchDirectory>();
            internal ScratchScope() { Directory.CreateDirectory(Container); }
            internal GitScratchDirectory Create(string output, string temporary)
            {
                var scratch = GitScratchDirectory.Create(output, temporary); scratches.Add(scratch); return scratch;
            }
            public void Dispose()
            {
                foreach (var scratch in scratches)
                {
                    scratch.ValidateCleanupRoot(scratch.Root);
                    Directory.Delete(scratch.Root, true);
                }
                string parent = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
                Guid identity;
                if (!string.Equals(Path.GetDirectoryName(Container), parent, StringComparison.OrdinalIgnoreCase) ||
                    !Guid.TryParseExact(Path.GetFileName(Container), "N", out identity) ||
                    (File.GetAttributes(Container) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidOperationException("Scratch regression cleanup lost its original temporary GUID boundary.");
                Directory.Delete(Container, true);
            }
        }
    }
}
