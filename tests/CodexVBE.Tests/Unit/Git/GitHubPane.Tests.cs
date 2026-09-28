using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie les actions GitHubPane par les événements réels de ses contrôles.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed partial class GitHubPaneCoverageTests
    {
                /// <summary>Vérifie les dépôts, le filtrage, la sélection et la création depuis l’interface.</summary>
[STATestMethod]
        public void DesignerRepositoriesFilteringSelectionAndCreationUseActualControlEvents()
        {
            using (var scope = new HostUiScope()) using (var pane = new GitHubPane())
            {
                pane.Configure("fixture", "https://github.com/owner/repo.git", "feature"); Assert.IsFalse(pane.Busy); Assert.IsNotNull(Field<TabControl>(pane, "pages")); Assert.AreEqual(2, Field<TabControl>(pane, "pages").TabCount); Assert.IsTrue(Field<Label>(pane, "sourceLabel").Text.EndsWith("feature")); Assert.IsFalse(Field<Button>(pane, "cancel").Enabled);
                Invoke(pane, "RepositoryChanged"); Invoke(pane, "UseRepository_Click");
                Api(pane, "[{\"full_name\":\"Owner/Alpha\",\"clone_url\":\"https://github.com/owner/alpha.git\",\"default_branch\":null},{\"full_name\":null,\"clone_url\":\"https://github.com/owner/other.git\",\"default_branch\":\"dev\"}]", "[{\"login\":\"team\"}]"); Invoke(pane, "LoadRepositories_Click"); Assert.AreEqual(2, Field<ListBox>(pane, "repositoryList").Items.Count); Assert.AreEqual(2, Field<ComboBox>(pane, "organization").Items.Count);
                var search = Field<TextBox>(pane, "repositorySearch"); search.Text = "ALPHA"; Assert.AreEqual(1, Field<ListBox>(pane, "repositoryList").Items.Count); search.Clear();
                Api(pane, "[{\"name\":\"main\"},{\"name\":\"dev\"}]"); Field<ListBox>(pane, "repositoryList").SelectedIndex = 0; Idle(pane); Assert.AreEqual("main", Field<ComboBox>(pane, "repositoryBranch").Text); Assert.AreEqual(2, Field<ComboBox>(pane, "repositoryBranch").Items.Count);
                Invoke(pane, "UseRepository_Click"); string chosen = null; pane.RepositorySelected = (url, branch) => chosen = url + ":" + branch; Invoke(pane, "UseRepository_Click"); StringAssert.Contains(chosen, "owner/alpha.git:main"); Field<ComboBox>(pane, "repositoryBranch").Text = "bad branch"; Invoke(pane, "UseRepository_Click"); Assert.IsFalse(string.IsNullOrWhiteSpace(Field<Label>(pane, "status").Text));
                foreach (bool organization in new[] { false, true }) foreach (bool defaultBranch in new[] { false, true }) foreach (bool callback in new[] { false, true }) { Field<ComboBox>(pane, "organization").SelectedIndex = organization ? 1 : 0; Field<TextBox>(pane, "repositoryName").Text = "new"; pane.RepositorySelected = callback ? (Action<string, string>)((u, b) => chosen = b) : null; Api(pane, "{\"full_name\":\"owner/new\",\"clone_url\":\"https://github.com/owner/new.git\",\"default_branch\":" + (defaultBranch ? "\"dev\"" : "null") + "}"); Invoke(pane, "CreateRepository_Click"); Assert.AreEqual(defaultBranch ? "dev" : "main", Field<ComboBox>(pane, "repositoryBranch").Text); if (callback) Assert.AreEqual(defaultBranch ? "dev" : "main", chosen); }
                Api(pane, "[{\"name\":\"dev\"}]"); Field<ListBox>(pane, "repositoryList").SelectedIndex = 1; Idle(pane); Assert.AreEqual("dev", Field<ComboBox>(pane, "repositoryBranch").Text);
            }
        }
                /// <summary>Vérifie les détails et brouillons de demandes de fusion ainsi que la navigation aux modules.</summary>
[STATestMethod]
        public void PullDetailsChecksDraftsAndModuleNavigationKeepNativeValidationAndNullCallbacks()
        {
            using (var scope = new HostUiScope()) using (var pane = new GitHubPane())
            {
                pane.Configure("fixture", "https://github.com/owner/repo.git", "feature"); Invoke(pane, "PullChanged"); Invoke(pane, "OpenPull_Click"); Invoke(pane, "CommentChanged"); Assert.AreEqual("", Field<TextBox>(pane, "commentBody").Text);
                foreach (bool main in new[] { false, true }) { Api(pane, "[{\"number\":12,\"title\":\"Title\"}]", main ? "[{\"name\":\"main\"}]" : "[{\"name\":\"dev\"}]"); Invoke(pane, "LoadPulls_Click"); if (main) Assert.AreEqual("main", Field<ComboBox>(pane, "targetBranch").Text); }
                foreach (bool merged in new[] { false, true }) foreach (bool checksFail in new[] { false, true })
                {
                    var handler = new LlmHttpFixture("{\"title\":\"Title\",\"state\":\"open\",\"body\":\"Body\",\"merged\":" + merged.ToString().ToLowerInvariant() + ",\"head\":{\"sha\":\"" + new string('a', 40) + "\"}}", "[{\"filename\":\"vba/M.bas\",\"status\":\"modified\"}]", "[{\"body\":\"general\"}]", "[{\"body\":\"inline\",\"path\":\"vba/M.bas\",\"line\":3}]"); handler.Replies.Enqueue(new LlmHttpFixture.Reply(checksFail ? "private-error" : "{\"check_runs\":[]}") { Status = checksFail ? HttpStatusCode.Forbidden : HttpStatusCode.OK }); if (!checksFail) handler.Replies.Enqueue(new LlmHttpFixture.Reply("{\"state\":\"success\"}")); pane.ApiFactory = a => new GitHubApi(a, handler, ct => Task.FromResult("fixture")); if (Field<ListBox>(pane, "pulls").SelectedIndex != 0) Field<ListBox>(pane, "pulls").SelectedIndex = 0; else Invoke(pane, "PullChanged"); Idle(pane); Assert.AreEqual(2, Field<ListBox>(pane, "comments").Items.Count); StringAssert.Contains(Field<TextBox>(pane, "pullDetails").Text, "Body"); Assert.AreEqual(merged, Field<TextBox>(pane, "pullDetails").Text.Contains(UiText.Get("Merged"))); StringAssert.Contains(Field<TextBox>(pane, "checks").Text, checksFail ? "403" : "success");
                }
                pane.OpenExternalLink = url => Assert.AreEqual("https://github.com/owner/repo/pull/13", url); Field<TextBox>(pane, "pullTitle").Text = "New"; Field<ComboBox>(pane, "targetBranch").Text = "main"; Api(pane, "{\"number\":13,\"html_url\":\"https://github.com/owner/repo/pull/13\"}"); Invoke(pane, "CreatePull_Click"); Invoke(pane, "OpenPull_Click"); pane.OpenExternalLink = url => throw new IOException("link unavailable"); Invoke(pane, "OpenPull_Click"); Assert.AreEqual("link unavailable", Field<Label>(pane, "status").Text);
                var files = Field<ListBox>(pane, "files"); var comments = Field<ListBox>(pane, "comments"); files.SelectedIndex = -1; Invoke(pane, "OpenFile_Click", files); comments.SelectedIndex = -1; Invoke(pane, "OpenFile_Click", comments); Assert.AreEqual("", Field<TextBox>(pane, "commentBody").Text);
                int line = 0; string module = null; foreach (bool callback in new[] { false, true }) { pane.OpenModule = callback ? (Action<string, int>)((n, l) => { module = n; line = l; }) : null; foreach (var path in new[] { null, "other/M.bas", "vba/folder/M.bas", "vba/M.bas" }) { files.Items.Clear(); files.Items.Add(new GitHubFile { filename = path }); files.SelectedIndex = 0; Invoke(pane, "OpenFile_Click", files); } if (callback) { Assert.AreEqual("M", module); Assert.AreEqual(1, line); } }
                comments.Items.Clear(); comments.Items.Add(new GitHubComment { path = "vba/M.bas", line = null, body = "text" }); comments.SelectedIndex = 0; Assert.AreEqual("text", Field<TextBox>(pane, "commentBody").Text); Invoke(pane, "OpenFile_Click", comments); Assert.AreEqual(1, line); comments.Items.Add(new GitHubComment { path = "vba/M.bas", line = 9 }); comments.SelectedIndex = 1; Invoke(pane, "OpenFile_Click", comments); Assert.AreEqual(9, line); pane.OpenModule = (n, l) => throw new IOException("module unavailable"); Invoke(pane, "OpenFile_Click", comments); Assert.AreEqual("module unavailable", Field<Label>(pane, "status").Text);
                Invoke(pane, "LoadDraft_Click"); pane.LoadDraft = () => null; Invoke(pane, "LoadDraft_Click"); pane.LoadDraft = () => new GitPullDraft { Target = "dev", Title = "Draft", Body = "Prepared" }; Invoke(pane, "LoadDraft_Click"); Assert.AreEqual("Prepared", Field<TextBox>(pane, "pullBody").Text); Assert.AreSame(Field<TabPage>(pane, "composeTab"), Field<TabControl>(pane, "pullTabs").SelectedTab); pane.LoadDraft = () => throw new IOException("draft unavailable"); Invoke(pane, "LoadDraft_Click"); Assert.AreEqual("draft unavailable", Field<Label>(pane, "status").Text);
            }
        }
                /// <summary>Vérifie l’annulation, la destruction et les erreurs des opérations en cours.</summary>
[STATestMethod]
        public void PendingOperationsCancellationDisposalAndErrorKindsRestoreEnabledState()
        {
            using (var scope = new HostUiScope())
            {
                using (var pane = new GitHubPane())
                {
                    LlmBoundaryScope.Pump((Task)LlmBoundaryScope.Call(pane, "Run", (Func<GitHubApi, CancellationToken, Task>)((a, ct) => Task.CompletedTask), true)); Assert.IsFalse(pane.Busy);
                    foreach (var error in new Exception[] { new OperationCanceledException(), new HttpRequestException("private endpoint"), new IOException("fixture failure") }) { LlmBoundaryScope.Pump((Task)LlmBoundaryScope.Call(pane, "Run", (Func<GitHubApi, CancellationToken, Task>)((a, ct) => Task.FromException(error)), false)); Assert.IsTrue(Field<TabControl>(pane, "pages").Enabled); Assert.IsFalse(Field<Button>(pane, "cancel").Enabled); Assert.IsFalse(Field<Label>(pane, "status").Text.Contains("private endpoint")); }
                    var pending = new TaskCompletionSource<bool>(); var task = (Task)LlmBoundaryScope.Call(pane, "Run", (Func<GitHubApi, CancellationToken, Task>)((a, ct) => { ct.Register(() => pending.TrySetCanceled()); return pending.Task; }), true); Assert.IsTrue(pane.Busy); int starts = 0; LlmBoundaryScope.Pump((Task)LlmBoundaryScope.Call(pane, "Run", (Func<GitHubApi, CancellationToken, Task>)((a, ct) => { starts++; return Task.CompletedTask; }), true)); Assert.AreEqual(0, starts); Invoke(pane, "Cancel_Click"); LlmBoundaryScope.Pump(task); Invoke(pane, "Cancel_Click"); Assert.IsFalse(pane.Busy);
                }
                using (var pane = new GitHubPane()) { var pending = new TaskCompletionSource<bool>(); var task = (Task)LlmBoundaryScope.Call(pane, "Run", (Func<GitHubApi, CancellationToken, Task>)((a, ct) => { ct.Register(() => pending.TrySetCanceled()); return pending.Task; }), true); pane.Dispose(); LlmBoundaryScope.Pump(task); Assert.IsTrue(pane.IsDisposed); }
                using (var pane = new GitHubPane()) LlmBoundaryScope.Call(pane, "Dispose", false);
                using (var pane = new GitHubPane()) { var components = Field<System.ComponentModel.IContainer>(pane, "components"); try { LlmBoundaryScope.Set(pane, "components", null); pane.Dispose(); } finally { LlmBoundaryScope.Set(pane, "components", components); components.Dispose(); } }
            }
        }
    }
}
