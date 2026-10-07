namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using VBAi;
    using VBAi.Tests.Infrastructure;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class GitHubAccountServiceTests
    {
        private static Task<string> Execute(GitHubAccountService service, string arguments, CancellationToken token)
        { return (Task<string>)typeof(GitHubAccountService).GetMethod("Execute", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(service, new object[] { arguments, token }); }

        [TestMethod]
        public async Task AccountListAndLoginUseDisposableCliAndValidateUntrustedAccountNames()
        {
            using (var scope = new AccountProcessScope())
            {
                var service = new GitHubAccountService(); service.StartProcess = p => scope.StartGit(p, "accounts");
                CollectionAssert.AreEqual(new[] { "alice", "zeta" }, await service.ListAsync(CancellationToken.None));
                StringAssert.Contains(scope.Commands[0], "github list --url https://github.com --no-ui");
                service.StartProcess = p => scope.StartGit(p, "empty"); await service.LoginAsync(CancellationToken.None);
                StringAssert.Contains(scope.Commands[1], "github login --url https://github.com --browser");
                CollectionAssert.AreEqual(new string[0], GitHubAccountService.ParseAccounts(null));
                CollectionAssert.AreEqual(new string[0], GitHubAccountService.ParseAccounts(""));
                foreach (string name in new[] { null, "", "-name", "name-", "space name", new string('x', 40) }) Assert.IsFalse(GitHubAccountService.ValidAccount(name));
                foreach (string name in new[] { "a", "A-9", new string('x', 39) }) Assert.IsTrue(GitHubAccountService.ValidAccount(name));
                Assert.ThrowsException<InvalidOperationException>(() => GitHubAccountService.ParseAccounts("alice\ninvalid token"));
                var alternate = new GitHubAccountService((arguments, token) => Task.FromResult("fixture"));
                CollectionAssert.AreEqual(new[] { "fixture" }, await alternate.ListAsync(CancellationToken.None));
            }
        }

        [TestMethod]
        public async Task NativeGitVersionMissingExecutableFailureAndTimeoutExposeOnlyPublicMessages()
        {
            using (var scope = new AccountProcessScope())
            {
                var service = new GitHubAccountService();
                StringAssert.StartsWith(await Execute(service, "--version", CancellationToken.None), "git version ");
                service.StartProcess = p => { p.StartInfo.FileName = scope.Executable + ".missing"; p.Start(); };
                var missing = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => service.ListAsync(CancellationToken.None));
                StringAssert.Contains(missing.Message, "Git for Windows");
                service.StartProcess = p => scope.StartGit(p, "error");
                var failure = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => service.ListAsync(CancellationToken.None));
                Assert.IsFalse(failure.Message.Contains("fixture diagnostic body"));
                service.StartProcess = p => scope.StartGit(p, "hang");
                service.WaitForExit = (p, milliseconds) => { Assert.AreEqual(300000, milliseconds); return p.WaitForExit(20); };
                await Assert.ThrowsExceptionAsync<TimeoutException>(() => service.ListAsync(CancellationToken.None));
                service.KillProcess = p => { p.Kill(); p.WaitForExit(5000); throw new InvalidOperationException("kill raced exit"); };
                await Assert.ThrowsExceptionAsync<TimeoutException>(() => service.ListAsync(CancellationToken.None));
            }
        }

        [TestMethod]
        public async Task CancellationMatrixHandlesPreCanceledActiveExitedAndTerminationRaces()
        {
            using (var scope = new AccountProcessScope())
            {
                var service = new GitHubAccountService(); int starts = 0;
                service.StartProcess = p => { starts++; scope.StartGit(p, "hang"); };
                using (var canceled = new CancellationTokenSource())
                {
                    canceled.Cancel(); await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => service.ListAsync(canceled.Token)); Assert.AreEqual(0, starts);
                }
                using (var canceled = new CancellationTokenSource())
                {
                    service.WaitForExit = (p, m) => { canceled.Cancel(); return p.WaitForExit(5000); };
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => service.ListAsync(canceled.Token));
                }
                using (var canceled = new CancellationTokenSource())
                {
                    service.StartProcess = p => scope.StartGit(p, "empty");
                    service.WaitForExit = (p, m) => { Assert.IsTrue(p.WaitForExit(5000)); canceled.Cancel(); return true; };
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => service.ListAsync(canceled.Token));
                }
                foreach (Exception error in new Exception[] { new InvalidOperationException("exit race"), new Win32Exception(5) })
                    using (var canceled = new CancellationTokenSource())
                    {
                        service.StartProcess = p => scope.StartGit(p, "hang");
                        service.KillProcess = p => { throw error; };
                        service.WaitForExit = (p, m) => { canceled.Cancel(); p.Kill(); return p.WaitForExit(5000); };
                        await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => service.ListAsync(canceled.Token));
                    }
            }
        }
    }
}
