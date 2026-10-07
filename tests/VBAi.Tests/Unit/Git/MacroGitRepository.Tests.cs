namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    /// <summary>Vérifie les règles de dépôt, les arbres de fichiers et le lancement des commandes Git natives.</summary>
    [DoNotParallelize]
    public sealed partial class MacroGitRepositoryTests
    {
        /// <summary>Vérifie les noms de branche, comptes, URL distantes et chemins de scope acceptés ou refusés.</summary>
        [TestMethod]
        public void BranchAccountRemoteAndScopeValidationMatrix()
        {
            foreach (string branch in new[] { null, "", "with space", "-invalid", "a..b", "a//b", "a/", "a/.hidden", "a/trailing.", "a/file.lock", new string('x', 129) })
                Assert.ThrowsException<ArgumentException>(() => new MacroGitRepository(Path.GetTempPath(), branch));
            foreach (string account in new[] { null, "", "coverage-fixture" })
                Assert.AreEqual("a/path-with.dots", new MacroGitRepository(Path.GetTempPath(), "a/path-with.dots", account).Branch);
            Assert.ThrowsException<ArgumentException>(() => new MacroGitRepository(Path.GetTempPath(), "main", "invalid account"));
            foreach (string remote in new[] { null, "", "http://github.com/o/r", "https://other.invalid/o/r", "https://user:token@github.com/o/r" })
                Assert.ThrowsException<ArgumentException>(() => MacroGitRepository.ValidateRemote(remote));
            Assert.AreEqual("https://github.com/example/fixture.git", MacroGitRepository.ValidateRemote(" https://github.com/example/fixture.git/ "));
            string scope = MacroGitRepository.ScopeDirectory("Disposable été scope");
            Assert.AreEqual(scope, MacroGitRepository.ScopeDirectory("Disposable été scope"));
            Assert.AreNotEqual(scope, MacroGitRepository.ScopeDirectory("Another disposable scope"));
        }

        [TestMethod]
        public void NativeLongCachePathCanInitializeCommitAndReadBackWithoutGlobalConfiguration()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                string prefix = Path.Combine(f.Root, "long-cache-");
                string cache = prefix + new string('x', Math.Max(1, 221 - Encoding.UTF8.GetByteCount(prefix)));
                Assert.IsTrue(Encoding.UTF8.GetByteCount(cache) > 220);
                var repository = new MacroGitRepository(cache, "main");
                repository.Initialize(f.Remote);
                Run(repository, new[] { "config", "user.name", "Qualification Fixture" });
                Run(repository, new[] { "config", "user.email", "qualification@example.invalid" });
                // Object paths extend beyond MAX_PATH even though the cache path itself is valid.
                var snapshot = f.Project.Capture();
                string commit = repository.Commit(snapshot, null, "Long cache qualification");
                repository.SetRef(repository.Head, commit);
                Assert.AreEqual(commit, repository.Resolve(repository.Head));
                Assert.IsTrue(snapshot.SameAs(repository.Read(commit)));
                repository.Initialize(f.Remote);
                Assert.ThrowsException<InvalidOperationException>(() => repository.Initialize("other-origin"));
                Assert.AreEqual(1, Run(repository, new[] { "config", "--local", "--get", "core.longpaths" }, allowFailure: true).ExitCode);
            }
        }

        /// <summary>Initialise un dépôt Git et vérifie la récupération, la synchronisation et l’ascendance des commits.</summary>
        [TestMethod]
        public void NativeInitializeFetchSynchronizationAndAncestryMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                var progress = new System.Collections.Generic.List<string>();
                f.Repository.Progress = progress.Add;
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.Initialize("another-local-origin"));
                Assert.AreEqual(0, f.Repository.History().Length);
                StringAssert.Contains(f.Repository.SynchronizationStatus(), UiText.Get("No local commit · use Fetch to inspect the remote repository"));
                string root = f.Commit(f.Project.Capture());
                f.Repository.SetRef("refs/remotes/origin/selected", root);
                StringAssert.Contains(f.Repository.SynchronizationStatus(), "premier pull requis");
                Assert.IsNull(f.Repository.Fetch()); Assert.IsNull(f.Repository.Resolve("refs/remotes/origin/selected"));
                f.Repository.SetRef(f.Repository.Head, root);
                StringAssert.Contains(f.Repository.SynchronizationStatus(), UiText.Get("Local branch · remote state unknown (Fetch)"));
                f.Repository.Push(root); Assert.AreEqual(root, f.Repository.Fetch()); Assert.AreEqual(1, f.Repository.History().Length);
                StringAssert.Contains(f.Repository.SynchronizationStatus(), "dernier Fetch");
                f.Git("--git-dir=" + f.Remote, "update-ref", "-d", "refs/heads/main");
                Assert.IsNull(f.Repository.Fetch()); Assert.IsNull(f.Repository.Resolve("refs/remotes/origin/selected"));
                string next = f.Commit(f.Snapshot("2"), root), sibling = f.Commit(f.Snapshot("3"), root);
                f.Repository.RequireFastForward(null, null); f.Repository.RequireFastForward(root, root); f.Repository.RequireFastForward(root, next);
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.RequireFastForward(root, null));
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.RequireFastForward(next, sibling));
                f.Git("--git-dir=" + f.Cache, "config", "codex.activeBranch", "other");
                f.Repository.Initialize(f.Remote); Assert.AreEqual("other", f.Repository.Branch);
                f.Git("--git-dir=" + f.Cache, "config", "codex.activeBranch", "invalid/.");
                Assert.ThrowsException<ArgumentException>(() => f.Repository.Initialize(f.Remote));
                Assert.ThrowsException<ArgumentException>(() => f.Repository.Commit(f.Project.Capture(), null, " "));
                Assert.IsNull(f.Repository.Read(null));
                var empty = Encoding.UTF8.GetString(Run(f.Repository, new[] { "mktree" }, new byte[0]).Bytes).Trim();
                Assert.IsNull(f.Repository.Read(empty));
                f.Repository.CompleteRecovery();
                Assert.IsTrue(progress.Count > 0);
            }
        }

        /// <summary>Vérifie les noms de fichiers du dépôt, sa structure et les limites de ressources d’export.</summary>
        [TestMethod]
        public void RepositoryTreeShapeNamesAndPackageResourceLimitsMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                f.Repository.CommandOverride = (args, input, use, allow) => Reply("100644 blob blob-id\tvba\0");
                Assert.ThrowsException<InvalidOperationException>(() => f.Repository.Read("root"));
                foreach (string entry in new[] { "120000 blob x\tlink.bas", "100644 tree x\tnested", "100644 blob x\tbad:name.bas" })
                {
                    f.Repository.CommandOverride = (args, input, use, allow) => Reply(args[2] == "root" ? "040000 tree subtree\tvba\0" : entry + "\0");
                    Assert.ThrowsException<InvalidOperationException>(() => f.Repository.Read("root"));
                }
                foreach (bool countLimit in new[] { false, true })
                {
                    string entries = countLimit ? string.Join("\0", Enumerable.Range(0, 2050).Select(i => "100644 blob blob-id\tM" + i + ".bas")) + "\0" : "100644 blob blob-id\tlarge.bas\0";
                    f.Repository.CommandOverride = (args, input, use, allow) =>
                    {
                        if (args[0] == "ls-tree") return Reply(args[2] == "root" ? "040000 tree subtree\tvba\0" : entries);
                        if (args[1] == "-s") return Reply(countLimit ? "0" : (VbaGitSnapshot.MaxBytes + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                        return Reply("");
                    };
                    Assert.ThrowsException<InvalidOperationException>(() => f.Repository.Read("root"));
                }
                f.Repository.CommandOverride = null;
                string commit = f.Seed(); Assert.IsTrue(f.Project.Capture().SameAs(f.Repository.Read(commit)));
            }
        }

        /// <summary>Vérifie l’environnement du processus Git, les identifiants et les erreurs ou délais dépassés.</summary>
        [TestMethod]
        public void NativeProcessEnvironmentCredentialsFailureAndTimeoutMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                var withAccount = new MacroGitRepository(f.Cache, "main", "coverage-fixture");
                var previous = Environment.GetEnvironmentVariable("GIT_COVERAGE_DISPOSABLE");
                try
                {
                    Environment.SetEnvironmentVariable("GIT_COVERAGE_DISPOSABLE", "remove from child");
                    var start = withAccount.StartProcess;
                    withAccount.StartProcess = process =>
                    {
                        Assert.IsNull(process.StartInfo.EnvironmentVariables["GIT_COVERAGE_DISPOSABLE"]);
                        Assert.AreEqual("0", process.StartInfo.EnvironmentVariables["GIT_TERMINAL_PROMPT"]);
                        StringAssert.Contains(process.StartInfo.Arguments, "credential.https://github.com.username=coverage-fixture");
                        return start(process);
                    };
                    Assert.AreEqual(0, Run(withAccount, new[] { "--version" }, useRepository: false).ExitCode);
                }
                finally { Environment.SetEnvironmentVariable("GIT_COVERAGE_DISPOSABLE", previous); }
                f.Repository.StartProcess = _ => false;
                Assert.ThrowsException<InvalidOperationException>(() => Run(f.Repository, new[] { "--version" }));
                f.Repository.StartProcess = ProcessInput.StartWithoutPreamble;
                f.Repository.WaitForExit = (process, timeout) => { Assert.AreEqual(120000, timeout); process.WaitForExit(); return false; };
                Assert.ThrowsException<TimeoutException>(() => Run(f.Repository, new[] { "--version" }));
                f.Repository.WaitForExit = (process, timeout) => process.WaitForExit(timeout);
                Assert.ThrowsException<InvalidOperationException>(() => Run(f.Repository, new[] { "not-a-git-command" }));
                Assert.AreNotEqual(0, Run(f.Repository, new[] { "not-a-git-command" }, allowFailure: true).ExitCode);
                f.Repository.StartProcess = process => { bool started = ProcessInput.StartWithoutPreamble(process); process.StandardInput.Close(); return started; };
                Assert.ThrowsException<AggregateException>(() => Run(f.Repository, new[] { "hash-object", "--stdin" }, new byte[1]));
            }
        }

        /// <summary>Vérifie l’annulation du processus Git avant, pendant et après sa terminaison.</summary>
        [TestMethod]
        public void NativeProcessCancellationBeforeDuringAndAfterExitMatrix()
        {
            using (var f = new MacroGitOperationsTests.Fixture())
            {
                using (var pending = new CancellationTokenSource())
                {
                    pending.Cancel(); f.Repository.Cancellation = pending.Token;
                    Assert.ThrowsException<OperationCanceledException>(() => Run(f.Repository, new[] { "--version" }));
                }
                foreach (int mode in new[] { 0, 1, 2 })
                    using (var pending = new CancellationTokenSource())
                    {
                        f.Repository.Cancellation = pending.Token;
                        f.Repository.WaitForExit = (process, timeout) =>
                        {
                            if (mode == 2) process.WaitForExit();
                            pending.Cancel(); return process.WaitForExit(timeout);
                        };
                        f.Repository.StopProcess = process =>
                        {
                            if (mode == 1) throw new InvalidOperationException("Disposable cancellation race");
                            process.Kill();
                        };
                        Assert.ThrowsException<OperationCanceledException>(() => Run(f.Repository, new[] { "hash-object", "--stdin" }, new byte[64 * 1024 * 1024]));
                    }
                f.Repository.Cancellation = CancellationToken.None;
            }
        }
    }
}
