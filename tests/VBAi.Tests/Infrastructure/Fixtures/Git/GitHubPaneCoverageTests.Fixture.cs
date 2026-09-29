using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class GitHubPaneCoverageTests
    {
        private static T Field<T>(GitHubPane pane, string name) { return LlmBoundaryScope.Get<T>(pane, name); }
        private static void Idle(GitHubPane pane)
        {
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (pane.Busy && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsFalse(pane.Busy, "The isolated pane operation did not finish.");
        }
        private static void Invoke(GitHubPane pane, string name, object sender = null) { LlmBoundaryScope.Call(pane, name, sender ?? pane, EventArgs.Empty); Idle(pane); }
        private static void Api(GitHubPane pane, params string[] replies) { pane.ApiFactory = account => new GitHubApi(account, new LlmHttpFixture(replies), ct => Task.FromResult("fixture-token")); }
    }
}
