using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class GitHubRepositoryInfo
    {
        public string full_name { get; set; }
        public string clone_url { get; set; }
        public string default_branch { get; set; }
        public override string ToString() { return full_name; }
    }
    internal sealed class GitHubPull
    {
        public int number { get; set; }
        public string title { get; set; }
        public string body { get; set; }
        public string state { get; set; }
        public bool draft { get; set; }
        public bool merged { get; set; }
        public string html_url { get; set; }
        public GitHubHead head { get; set; }
        public override string ToString() { return "#" + number + " · " + state + (draft ? " · " + UiText.Get("Draft") : "") + " · " + title; }
    }
    internal sealed class GitHubHead { public string sha { get; set; } }
    internal sealed class GitHubFile { public string filename { get; set; } public string status { get; set; } public override string ToString() { return status + " · " + filename; } }
    internal sealed class GitHubComment { public string body { get; set; } public string path { get; set; } public int? line { get; set; } public override string ToString() { return (path == null ? "" : path + ":" + line + " · ") + body; } }
    internal sealed class GitHubCheck { public string name { get; set; } public string status { get; set; } public string conclusion { get; set; } }
    internal sealed class GitHubChecks { public GitHubCheck[] check_runs { get; set; } }
    internal sealed class GitHubStatus { public string state { get; set; } }
    internal sealed class GitHubOrganization { public string login { get; set; } }
    internal sealed class GitHubBranch { public string name { get; set; } }

    // Credentials are obtained from the user's existing GCM account, kept in memory and
    // sent only to api.github.com. Response bodies and credentials never enter diagnostics.
    internal sealed class GitHubApi : IDisposable
    {
        private readonly HttpClient client;
        private readonly Func<CancellationToken, Task<string>> credential;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        internal GitHubApi(string account, HttpMessageHandler handler = null, Func<CancellationToken, Task<string>> credential = null)
        {
            client = new HttpClient(handler ?? new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(45) };
            this.credential = credential ?? (ct => ReadCredential(account, ct));
        }
        public void Dispose() { client.Dispose(); }
        internal static string RepositoryPath(string url)
        {
            var path = new Uri(MacroGitRepository.ValidateRemote(url)).AbsolutePath.Trim('/');
            if (path.EndsWith(".git", StringComparison.OrdinalIgnoreCase)) path = path.Substring(0, path.Length - 4);
            return "/repos/" + path;
        }
        internal async Task<T> Request<T>(HttpMethod method, string path, object data, CancellationToken ct)
        {
            if (!path.StartsWith("/", StringComparison.Ordinal) || path.StartsWith("//", StringComparison.Ordinal) || path.Contains("\\") || path.Contains("..")) throw new ArgumentException("Invalid GitHub API path.");
            using (var request = new HttpRequestMessage(method, "https://api.github.com" + path))
            {
                request.Headers.UserAgent.ParseAdd("VBAi/1.0");
                request.Headers.Accept.ParseAdd("application/vnd.github+json");
                request.Headers.Add("X-GitHub-Api-Version", "2026-03-10");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await credential(ct));
                if (data != null) request.Content = new StringContent(json.Serialize(data), Encoding.UTF8, "application/json");
                using (var response = await client.SendAsync(request, ct))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        int code = (int)response.StatusCode;
                        string help = code == 401 ? "Sign in to GitHub again in Settings." : code == 403 || code == 429 ? "Check repository permissions or wait for the GitHub rate limit to reset." :
                            code == 404 ? "Check the repository name and account access." : code == 422 ? "Check the branch, title and repository name; a pull request may already exist." : "GitHub is unavailable. Try again later.";
                        throw new InvalidOperationException("GitHub " + code + " · " + UiText.Get(help));
                    }
                    string body = await response.Content.ReadAsStringAsync();
                    ct.ThrowIfCancellationRequested();
                    return json.Deserialize<T>(body);
                }
            }
        }
        internal async Task<T[]> List<T>(string path, CancellationToken ct)
        {
            var result = new List<T>();
            for (int page = 1; ; page++)
            {
                ct.ThrowIfCancellationRequested();
                var items = await Request<T[]>(HttpMethod.Get, path + (path.Contains("?") ? "&" : "?") + "per_page=100&page=" + page, null, ct);
                result.AddRange(items);
                if (items.Length < 100) return result.ToArray();
            }
        }
        internal Task<GitHubRepositoryInfo[]> Repositories(CancellationToken ct) { return List<GitHubRepositoryInfo>("/user/repos?sort=updated&affiliation=owner,collaborator,organization_member", ct); }
        internal Task<GitHubOrganization[]> Organizations(CancellationToken ct) { return List<GitHubOrganization>("/user/orgs", ct); }
        internal Task<GitHubBranch[]> Branches(string url, CancellationToken ct) { return List<GitHubBranch>(RepositoryPath(url) + "/branches", ct); }
        internal Task<GitHubRepositoryInfo> CreateRepository(string name, string organization, bool isPrivate, CancellationToken ct)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(name ?? "", "^[A-Za-z0-9_.-]{1,100}$")) throw new ArgumentException(UiText.Get("Invalid repository name."));
            if (!string.IsNullOrEmpty(organization) && !GitHubAccountService.ValidAccount(organization)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            return Request<GitHubRepositoryInfo>(HttpMethod.Post, string.IsNullOrEmpty(organization) ? "/user/repos" : "/orgs/" + organization + "/repos", new { name, @private = isPrivate, auto_init = false }, ct);
        }
        internal Task<GitHubPull[]> Pulls(string url, CancellationToken ct) { return List<GitHubPull>(RepositoryPath(url) + "/pulls?state=all&sort=updated", ct); }
        internal Task<GitHubPull> CreatePull(string url, string head, string target, string title, string body, bool draft, CancellationToken ct)
        {
            MacroGitRepository.ValidateBranch(head); MacroGitRepository.ValidateBranch(target);
            if (string.IsNullOrWhiteSpace(title) || head == target) throw new ArgumentException(UiText.Get("Enter a title and a different target branch."));
            return Request<GitHubPull>(HttpMethod.Post, RepositoryPath(url) + "/pulls", new { head, @base = target, title, body, draft }, ct);
        }
        internal async Task<string> Checks(string url, string sha, CancellationToken ct)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(sha ?? "", "^[a-f0-9]{40}$")) throw new ArgumentException("Invalid commit.");
            var lines = new List<string>();
            for (int page = 1; ; page++)
            {
                var checks = await Request<GitHubChecks>(HttpMethod.Get, RepositoryPath(url) + "/commits/" + sha + "/check-runs?per_page=100&page=" + page, null, ct);
                lines.AddRange(checks.check_runs.Select(x => x.name + ": " + (x.conclusion ?? x.status)));
                if (checks.check_runs.Length < 100) break;
            }
            var status = await Request<GitHubStatus>(HttpMethod.Get, RepositoryPath(url) + "/commits/" + sha + "/status", null, ct);
            return UiText.Get("Commit status") + ": " + status.state + Environment.NewLine + string.Join(Environment.NewLine, lines);
        }
        private static async Task<string> ReadCredential(string account, CancellationToken ct)
        {
            if (!string.IsNullOrEmpty(account) && !GitHubAccountService.ValidAccount(account)) throw new ArgumentException(UiText.Get("Invalid GitHub account."));
            var start = new ProcessStartInfo("git.exe", "credential-manager get") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true };
            start.EnvironmentVariables["GCM_INTERACTIVE"] = "never";
            start.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            using (var process = new Process { StartInfo = start })
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(30));
                ProcessInput.StartWithoutPreamble(process);
                using (timeout.Token.Register(() => { try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { } }))
                {
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    await process.StandardInput.WriteAsync("protocol=https\nhost=github.com\n" + (string.IsNullOrEmpty(account) ? "" : "username=" + account + "\n") + "\n");
                    process.StandardInput.Close();
                    await Task.Run(() => process.WaitForExit());
                    await error; string result = await output;
                    timeout.Token.ThrowIfCancellationRequested();
                    string token = result.Split('\n').FirstOrDefault(x => x.StartsWith("password=", StringComparison.Ordinal));
                    if (process.ExitCode != 0 || token == null) throw new InvalidOperationException(UiText.Get("Sign in to GitHub again in Settings."));
                    return token.Substring(9).TrimEnd('\r');
                }
            }
        }
    }
}
