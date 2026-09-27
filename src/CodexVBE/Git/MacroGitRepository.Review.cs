using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace CodexVBE
{
    internal sealed class GitPullDraft { public string Branch { get; set; } public string Target { get; set; } public string Title { get; set; } public string Body { get; set; } }
    internal sealed class GitCommitInfo
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public string Date { get; set; }
        public string Subject { get; set; }
        public override string ToString() { return Id.Substring(0, 8) + " · " + Date + " · " + Subject; }
    }
    internal sealed partial class MacroGitRepository
    {
        internal CancellationToken Cancellation { get; set; }
        internal Action<string> Progress { get; set; }
        internal string RemoteUrl { get { return ValidateRemote(Text("remote", "get-url", "origin")); } }
        internal string VerifiedCommit(string id)
        {
            if (!Regex.IsMatch(id ?? "", "^[a-fA-F0-9]{7,40}$")) throw new ArgumentException(UiText.Get("Select a commit from history."));
            return Resolve(id + "^{commit}") ?? throw new ArgumentException(UiText.Get("Commit not found."));
        }
        internal GitCommitInfo[] Commits()
        {
            string head = Resolve(Head);
            if (head == null) return new GitCommitInfo[0];
            return Text("log", "-200", "--date=iso-strict", "--format=%H%x09%an%x09%ad%x09%s", head)
                .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(line => {
                    var fields = line.TrimEnd('\r').Split(new[] { '\t' }, 4);
                    return new GitCommitInfo { Id = fields[0], Author = fields[1], Date = fields[2], Subject = fields[3] };
                }).ToArray();
        }
        internal string CommitDetails(string id) { return Text("show", "--no-patch", "--format=fuller", VerifiedCommit(id)); }
        internal string ParentCommit(string id) { return Resolve(VerifiedCommit(id) + "^1"); }
        internal void SavePullDraft(string target, string title, string body)
        {
            ValidateBranch(target);
            if (string.IsNullOrWhiteSpace(title) || title.Length > 256 || (body ?? "").Length > 60000) throw new ArgumentException(UiText.Get("Invalid pull request draft."));
            string file = System.IO.Path.Combine(directory, "codex-pr-draft.json");
            string pending = file + ".pending";
            System.IO.File.WriteAllText(pending, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(new GitPullDraft { Branch = Branch, Target = target, Title = title, Body = body }));
            if (System.IO.File.Exists(file)) System.IO.File.Replace(pending, file, null);
            else System.IO.File.Move(pending, file);
        }
        internal GitPullDraft PullDraft()
        {
            string file = System.IO.Path.Combine(directory, "codex-pr-draft.json");
            if (!System.IO.File.Exists(file)) return null;
            var draft = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<GitPullDraft>(System.IO.File.ReadAllText(file));
            return draft.Branch == Branch ? draft : null;
        }
    }
}
