namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class MacroGitRepositoryTests
    {
        [TestMethod]
        public void CommitReviewAndDraftValidationMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                Assert.AreEqual(0, f.Repository.Commits().Length); Assert.IsNull(f.Repository.PullDraft());
                foreach (string id in new[] { null, "", "invalid", "fffffff" }) Assert.ThrowsException<ArgumentException>(() => f.Repository.VerifiedCommit(id));
                string root = f.Seed(); var commit = f.Repository.Commits()[0];
                Assert.AreEqual(root, commit.Id); Assert.IsFalse(string.IsNullOrWhiteSpace(commit.Author)); Assert.IsFalse(string.IsNullOrWhiteSpace(commit.Subject));
                StringAssert.Contains(commit.ToString(), commit.Date); Assert.AreEqual(root, f.Repository.VerifiedCommit(root.Substring(0, 7)));
                StringAssert.Contains(f.Repository.CommitDetails(root), "Coverage Fixture"); Assert.IsNull(f.Repository.ParentCommit(root));
                string next = f.Commit(f.Snapshot("2"), root); Assert.AreEqual(root, f.Repository.ParentCommit(next));
                foreach (string title in new[] { null, " ", new string('x', 257) }) Assert.ThrowsException<ArgumentException>(() => f.Repository.SavePullDraft("main", title, null));
                Assert.ThrowsException<ArgumentException>(() => f.Repository.SavePullDraft("main", "Valid title", new string('x', 60001)));
                f.Repository.SavePullDraft("main", "First draft", null); Assert.IsNull(f.Repository.PullDraft().Body);
                f.Repository.SavePullDraft("main", "Second draft", "Disposable body");
                var draft = f.Repository.PullDraft(); Assert.AreEqual("main", draft.Branch); Assert.AreEqual("main", draft.Target);
                Assert.AreEqual("Second draft", draft.Title); Assert.AreEqual("Disposable body", draft.Body);
                f.Repository.CreateBranch("feature"); f.Repository.SelectBranch("feature"); Assert.IsNull(f.Repository.PullDraft());
                f.Git("--git-dir=" + f.Cache, "remote", "set-url", "origin", "https://github.com/example/disposable.git");
                Assert.AreEqual("https://github.com/example/disposable.git", f.Repository.RemoteUrl);
            }
        }
    }
}
