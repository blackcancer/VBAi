namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class MacroGitRepositoryTests
    {
        [TestMethod]
        public void BranchTrackingCheckpointsAndMergeStateGuardsMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.CreateBranch("empty"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.BranchCommit("missing"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.BeginMerge("missing"));
                string root = f.Seed(); f.Repository.CreateBranch("feature"); f.Repository.Push(root);
                f.Repository.SelectBranch("feature"); Assert.IsNull(f.Repository.Resolve("refs/remotes/origin/selected"));
                f.Repository.Initialize(f.Remote); Assert.AreEqual("feature", f.Repository.Branch);
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.TrackBranch("feature"));
                f.Git("--git-dir=" + f.Remote, "update-ref", "refs/heads/tracked", root);
                f.Repository.TrackBranch("tracked"); Assert.AreEqual(root, f.Repository.BranchCommit("tracked"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.BeginMerge("missing"));
                foreach (string label in new[] { null, " ", new string('x', 161), "invalid\nname" })
                    Assert.ThrowsException<ArgumentException>(() => f.Repository.Checkpoint(f.Project.Capture(), label));
                var checkpoint = f.Repository.Checkpoint(f.Project.Capture(), "Disposable checkpoint");
                StringAssert.Contains(checkpoint.ToString(), "Disposable checkpoint");
                Assert.AreEqual(checkpoint.Commit, f.Repository.CheckpointCommit(checkpoint.Id));
                Assert.AreEqual(1, f.Repository.Checkpoints().Length);
                Assert.ThrowsException<ArgumentException>(() => f.Repository.CheckpointCommit(null));
                Assert.ThrowsException<ArgumentException>(() => f.Repository.CheckpointCommit("invalid"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.CheckpointCommit("20260928000000000-deadbeef"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.AssertMerge(null));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.AssertMerge(new GitMergePlan { Branch = "wrong", Ours = root }));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.AssertMerge(new GitMergePlan { Branch = "feature", Ours = new string('0', 40) }));
                f.Repository.AbortMerge();
                var plan = f.Repository.BeginMerge("main");
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.BeginMerge("main"));
                Assert.ThrowsException<ArgumentException>(() => f.Repository.ConflictContent("absent"));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.ResolveConflict("absent", "ours"));
                Assert.ThrowsException<ArgumentException>(() => f.Repository.MergeCommit(plan, " "));
                plan.Conflicts = new[] { "vba/Module1.bas" };
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.MergeCommit(plan, "Unresolved"));
                f.Repository.AbortMerge(); f.Repository.AbortMerge();
                foreach (bool invalidTree in new[] { false, true })
                {
                    f.Repository.CommandOverride = (args, input, use, allow) => args[0] == "merge-tree"
                        ? Reply("invalid-tree", invalidTree ? 0 : 2) : Native(f.Repository, args, input, use, allow);
                    Assert.ThrowsException<InvalidOperationException>(() => f.Repository.BeginMerge("main"));
                }
                f.Repository.CommandOverride = null;
            }
        }

        [TestMethod]
        public void ConflictPreviewHandlesMissingDirectoriesBinaryLargeAndInvalidUtf8NativeObjects()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                Func<byte[], string> blob = bytes => Encoding.UTF8.GetString(Run(f.Repository, new[] { "hash-object", "-w", "--stdin" }, bytes).Bytes).Trim();
                string empty = Encoding.UTF8.GetString(Run(f.Repository, new[] { "mktree" }, new byte[0]).Bytes).Trim();
                string plain = blob(Encoding.UTF8.GetBytes("Disposable source")), bad = blob(new byte[] { 255 }), large = blob(new byte[65537]);
                string tree = (string)Call(f.Repository, "MakeTree", (object)new[] {
                    "100644 blob " + plain + "\tsource.bas", "100644 blob " + plain + "\tpicture.frx",
                    "100644 blob " + bad + "\tinvalid.bas", "100644 blob " + large + "\tlarge.bas", "040000 tree " + empty + "\tfolder" });
                Assert.AreEqual("", Call(f.Repository, "ConflictText", tree, "missing.bas"));
                Assert.AreEqual("Disposable source", Call(f.Repository, "ConflictText", tree, "source.bas"));
                foreach (string path in new[] { "folder", "picture.frx" })
                    StringAssert.Contains((string)Call(f.Repository, "ConflictText", tree, path), UiText.Get("[Binary resource or directory: choose local or incoming]"));
                StringAssert.Contains((string)Call(f.Repository, "ConflictText", tree, "large.bas"), UiText.Get("[Preview unavailable above 64 KiB: choose local or incoming]"));
                StringAssert.Contains((string)Call(f.Repository, "ConflictText", tree, "invalid.bas"), UiText.Get("[Non-UTF-8 content: choose local or incoming]"));
                string removed = (string)Call(f.Repository, "ReplacePath", tree, new[] { "source.bas" }, 0, null);
                Assert.IsNull(Call(f.Repository, "TreeAtPath", removed, "source.bas"));
                Assert.ThrowsException<InvalidOperationException>(() => Call(f.Repository, "ReplacePath", tree, new[] { "source.bas", "nested" }, 0, null));
                string leaf = "100644 blob " + plain + "\tleaf.bas";
                string nested = (string)Call(f.Repository, "ReplacePath", null, new[] { "new", "leaf.bas" }, 0, leaf);
                Assert.IsNotNull(Call(f.Repository, "TreeAtPath", nested, "new/leaf.bas"));
                nested = (string)Call(f.Repository, "ReplacePath", nested, new[] { "new", "leaf.bas" }, 0, null);
                Assert.IsNull(Call(f.Repository, "TreeAtPath", nested, "new/leaf.bas"));
                Assert.AreEqual(empty, Call(f.Repository, "ReplacePath", null, new[] { "absent.bas" }, 0, null));
            }
        }

        [TestMethod]
        public void ConflictResolutionChoicesAndTextEligibilityMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                f.Seed(); f.Repository.CreateBranch("feature");
                var plan = f.Repository.BeginMerge("feature");
                foreach (var item in new[] {
                    new { Path = "vba/Module1.bas", Choice = "unknown", Text = "text" },
                    new { Path = "outside.bas", Choice = "text", Text = "text" },
                    new { Path = "vba/picture.frx", Choice = "text", Text = "text" },
                    new { Path = "vba/Module1.bas", Choice = "text", Text = (string)null },
                    new { Path = "vba/Module1.bas", Choice = "text", Text = new string('x', 1024 * 1024) } })
                {
                    plan.Conflicts = new[] { item.Path }; Call(f.Repository, "SaveMerge", plan);
                    Assert.ThrowsException<ArgumentException>(() => f.Repository.ResolveConflict(item.Path, item.Choice, item.Text));
                }
                foreach (string choice in new[] { "ours", "theirs", "text" })
                {
                    plan.Conflicts = new[] { "vba/Module1.bas" }; Call(f.Repository, "SaveMerge", plan);
                    var resolved = f.Repository.ResolveConflict("vba/Module1.bas", choice,
                        "Attribute VB_Name = \"Module1\"\nOption Explicit\n");
                    Assert.AreEqual(0, resolved.Conflicts.Length); plan = resolved;
                }
                plan.Conflicts = new[] { "vba/missing.bas" }; Call(f.Repository, "SaveMerge", plan);
                Assert.AreEqual(0, f.Repository.ResolveConflict("vba/missing.bas", "ours").Conflicts.Length);
                f.Repository.AbortMerge();
            }
        }

        [TestMethod]
        public void NativeConflictingBranchesResolveToAValidatedMergeCommit()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                string root = f.Seed(); f.Repository.CreateBranch("incoming");
                f.Repository.SetRef("refs/heads/incoming", f.Commit(f.Snapshot("2"), root));
                f.Repository.SetRef(f.Repository.Head, f.Commit(f.Snapshot("3"), root));
                var plan = f.Repository.BeginMerge("incoming"); Assert.AreEqual(1, plan.Conflicts.Length);
                var content = f.Repository.ConflictContent("vba/Module1.bas");
                Assert.AreEqual("vba/Module1.bas", content.Path);
                StringAssert.Contains(content.Base, "Value = 1"); StringAssert.Contains(content.Ours, "Value = 3"); StringAssert.Contains(content.Theirs, "Value = 2");
                plan = f.Repository.ResolveConflict("vba/Module1.bas", "ours");
                string merged = f.Repository.MergeCommit(plan, "Validated native merge");
                Assert.IsTrue(f.Snapshot("3").SameAs(f.Repository.Read(merged)));
                Assert.IsTrue(f.Repository.Branches().Length >= 2);
                f.Repository.AbortMerge();
            }
        }
    }
}
