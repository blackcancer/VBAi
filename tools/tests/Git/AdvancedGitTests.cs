using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;

internal static partial class GitTests
{
    private static T Pump<T>(Task<T> task)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        while (!task.IsCompleted)
        {
            System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(5);
            if (watch.ElapsedMilliseconds > 60000) throw new Exception("Asynchronous test timed out.");
        }
        return task.GetAwaiter().GetResult();
    }
    private static void Advanced()
    {
        string remote = Path.Combine(root, "advanced-origin.git"); Git(root, "init --bare \"" + remote + "\"");
        var repo = Repo("advanced", remote);
        var host = new FakeProject { FileName = Path.Combine(root, "advanced.xlsm") };
        const string source = "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n";
        host.VBComponents.Add(new FakeComponent("Module1", 1, source));
        var project = new VbaGitProject(() => host, host.FileName);
        var operations = new MacroGitOperations(project, repo);
        var json = new JavaScriptSerializer();
        Func<string> revision = () => operations.Revision(project.Capture());
        Func<string, string, string, object> action = (command, name, text) => Pump(operations.ExecuteAsync(command, revision(), name, text));
        action("commit", null, "Base");
        string initial = repo.Resolve(repo.Head);
        var checkpoint = (GitCheckpoint)action("checkpoint_create", "Version stable", null);
        Assert(repo.Checkpoints().Any(x => x.Id == checkpoint.Id && x.Label == "Version stable"), "Named checkpoint discoverable");
        host.VBComponents.Item("Module1").CodeModule.Text = source.Replace("Value = 1", "Value = 9");
        action("checkpoint_restore", checkpoint.Id, null);
        Assert(host.VBComponents.Item("Module1").CodeModule.Text == source, "Checkpoint restores uncommitted source");
        Assert(repo.Resolve(repo.Head) == initial, "Checkpoint does not reset branch history");
        Assert(repo.Checkpoints().Length >= 2, "Restore preserves displaced work in automatic checkpoint");
        action("branch_create", "feature/test", null);
        Reject(() => action("branch_create", "feature/test", null), "Existing branch not overwritten");
        action("branch_switch", "feature/test", null);
        Assert(repo.Branch == "feature/test", "Switch changes active branch");
        host.VBComponents.Item("Module1").CodeModule.Text = source.Replace("Value = 1", "Value = 2");
        Reject(() => action("branch_switch", "main", null), "Dirty branch switch refused");
        action("commit", null, "Feature");
        action("branch_switch", "main", null);
        Assert(project.Capture().Files["Module1.bas"].SequenceEqual(System.Text.Encoding.UTF8.GetBytes(source)), "Switch imports target branch");
        host.VBComponents.Item("Module1").CodeModule.Text = source.Replace("Value = 1", "Value = 3");
        action("commit", null, "Main edit");
        string beforeMerge = repo.Resolve(repo.Head);
        var plan = (GitMergePlan)action("merge_begin", "feature/test", null);
        Assert(plan.Conflicts.Contains("vba/Module1.bas"), "Three-way merge reports exact path");
        Assert(repo.Resolve(repo.Head) == beforeMerge && host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 3"), "Conflicted merge does not mutate VBA or HEAD");
        Reject(() => action("merge_complete", null, "Merge"), "Unresolved merge cannot be committed");
        var content = Pump(operations.ConflictAsync("vba/Module1.bas"));
        Assert(content.Ours.Contains("Value = 3") && content.Theirs.Contains("Value = 2"), "Both conflict sides available");
        Pump(operations.ExecuteAsync("merge_resolve", revision(), text: "", choice: "theirs", path: "vba/Module1.bas"));
        action("merge_complete", null, "Merge resolved feature");
        Assert(host.VBComponents.Item("Module1").CodeModule.Text.Contains("Value = 2"), "Resolved merge imports chosen content");
        Assert(Git(root, "--git-dir=\"" + Path.Combine(root, "advanced") + "\" rev-list --parents -n 1 " + repo.Resolve(repo.Head)).Trim().Split(' ').Length == 3, "Two merge parents preserved");
        Assert(repo.PendingMerge == null, "Merge plan cleared after successful apply");
        action("merge_begin", "feature/test", null); action("merge_abort", null, null);
        Assert(repo.PendingMerge == null && repo.Resolve(repo.Head) != beforeMerge, "Abort prepared merge leaves committed history intact");
        action("push", null, null);
        Assert(!Git(root, "--git-dir=\"" + remote + "\" for-each-ref").Contains("refs/codex"), "Named checkpoints and merge drafts stay private after push");
        var other = Repo("advanced-other", remote);
        string downloaded = other.Fetch(); other.SetRef(other.Head, downloaded); other.CreateBranch("remote-feature");
        Git(root, "--git-dir=\"" + Path.Combine(root, "advanced-other") + "\" push origin refs/heads/remote-feature");
        Assert(repo.RemoteBranches().Contains("remote-feature"), "Remote branch discovery");
        action("branch_track", "remote-feature", null);
        Assert(repo.Branches().Contains("remote-feature") && repo.Branch == "main", "Tracking remote branch preserves active VBA branch");
        var reopened = new MacroGitRepository(Path.Combine(root, "advanced"), "unused"); reopened.Initialize(remote);
        Assert(reopened.Branch == "main", "Active branch restored when cache reopens");
        string stale = revision(); host.VBComponents.Item("Module1").CodeModule.Text += "' local change\n";
        Reject(() => Pump(operations.ExecuteAsync("checkpoint_create", stale, "Stale")), "Stale Git state rejected");

        var settings = new LlmSettings { VbeEditApproval = "Automatic" };
        var tools = new LlmVbeTools(null, null, settings) { BoundProject = host.FileName, GitOperationsFactory = ignored => new MacroGitOperations(project, repo) };
        Func<string, Dictionary<string, object>, Response> invoke = (name, input) => json.Deserialize<Response>(Pump(tools.InvokeAsync(name, json.Serialize(input))));
        Func<Dictionary<string, object>> argsForEdit = () => new Dictionary<string, object> { { "Project", host.FileName }, { "ExpectedState", revision() } };
        Assert(LlmVbeTools.Definitions.Cast<dynamic>().Any(x => (string)x.function.name == "git_branch_switch"), "Git definitions reach provider catalog");
        Assert(LlmVbeTools.Definitions.Length <= 128 && LlmVbeTools.Definitions.Cast<dynamic>().Select(x => (string)x.function.name).Distinct().Count() == LlmVbeTools.Definitions.Length, "Provider tool catalog remains unique and within 128 tools");
        var status = invoke("git_status", new Dictionary<string, object> { { "Project", host.FileName } });
        Assert(status.Ok && ((Dictionary<string, object>)status.Data).ContainsKey("State"), "Agent status returns revision");
        var args = argsForEdit(); args["Name"] = "Agent checkpoint";
        Assert(invoke("git_checkpoint_create", args).Ok, "Agent can create actual checkpoint");
        args = argsForEdit(); args["Name"] = "agent-branch";
        Assert(invoke("git_branch_create", args).Ok && repo.Branches().Contains("agent-branch"), "Agent can create actual Git branch");
        tools.Mode = ChatMode.Plan;
        args = argsForEdit(); args["Name"] = "forbidden";
        Assert(!invoke("git_branch_create", args).Ok, "Plan blocks Git mutation");
        Assert(invoke("git_status", new Dictionary<string, object> { { "Project", host.FileName } }).Ok, "Plan permits Git status");
        tools.Mode = ChatMode.Agent; settings.VbeEditApproval = "ReadOnly";
        Assert(!invoke("git_checkpoint_create", args).Ok, "ReadOnly policy blocks Git mutation");
        settings.VbeEditApproval = "Automatic";
        args = argsForEdit(); args["Project"] = "other.xlsm";
        Assert(!invoke("git_status", new Dictionary<string, object> { { "Project", "other.xlsm" } }).Ok, "Even Git reads reject another document");
        args = argsForEdit(); args["Remote"] = "https://github.com/other/repo";
        Assert(!invoke("git_push", args).Ok, "Agent cannot redirect a push");
        args = argsForEdit(); args["ExpectedState"] = stale; args["Name"] = "stale-agent";
        Assert(!invoke("git_checkpoint_create", args).Ok, "Agent stale revision rejected");
        Console.WriteLine("PASS named checkpoints, branch switches, persisted branch, real merges/conflicts and agent tool policies");
    }
}
