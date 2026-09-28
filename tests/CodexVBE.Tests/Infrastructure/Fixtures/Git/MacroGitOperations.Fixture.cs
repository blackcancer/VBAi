namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using CodexVBE;

    public sealed partial class MacroGitOperationsTests
    {
        internal sealed class Fixture : IDisposable
        {
            internal readonly string Root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "git-branches", Guid.NewGuid().ToString("N"));
            internal readonly string Cache, Remote;
            internal readonly global::FakeProject Host;
            internal readonly VbaGitProject Project;
            internal readonly MacroGitRepository Repository;
            internal readonly MacroGitOperations Operations;
            internal Fixture()
            {
                Directory.CreateDirectory(Root);
                Cache = Path.Combine(Root, "cache"); Remote = Path.Combine(Root, "origin.git");
                Git("init", "--bare", Remote);
                Repository = new MacroGitRepository(Cache, "main"); Repository.Initialize(Remote);
                Git("--git-dir=" + Cache, "config", "user.name", "Coverage Fixture");
                Git("--git-dir=" + Cache, "config", "user.email", "coverage@example.invalid");
                Host = new global::FakeProject { FileName = Path.Combine(Root, "fixture.xlsm") };
                Host.VBComponents.Add(new global::FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
                Project = new VbaGitProject(() => Host, Host.FileName);
                Operations = new MacroGitOperations(Project, Repository);
            }
            internal void Git(params string[] arguments)
            {
                var info = new ProcessStartInfo("git.exe", string.Join(" ", Array.ConvertAll(arguments, a => "\"" + a.Replace("\"", "\\\"") + "\"")))
                { WorkingDirectory = Root, UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
                using (var process = Process.Start(info))
                {
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(30000)) { process.Kill(); throw new TimeoutException("Fixture Git timed out"); }
                    if (process.ExitCode != 0) throw new InvalidOperationException(error.GetAwaiter().GetResult());
                    output.GetAwaiter().GetResult();
                }
            }
            internal VbaGitSnapshot Snapshot(string value)
            {
                var files = Project.Capture().Serialize();
                files["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = " + value + "\n");
                return VbaGitSnapshot.Read(files);
            }
            internal string Commit(VbaGitSnapshot snapshot, string parent = null)
            { return Repository.Commit(snapshot, parent, "Fixture commit"); }
            internal string Seed()
            {
                string commit = Commit(Project.Capture());
                Repository.SetRef(Repository.Head, commit); Repository.SetRef(MacroGitRepository.Baseline, commit);
                return commit;
            }
            internal object Import(VbaGitSnapshot target, VbaGitSnapshot expected, bool rollback = false)
            {
                try
                {
                    var task = (System.Threading.Tasks.Task)typeof(MacroGitOperations).GetMethod("ImportAsync", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(Operations, new object[] { target, expected, rollback });
                    task.GetAwaiter().GetResult(); return null;
                }
                catch (System.Reflection.TargetInvocationException ex)
                { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
            }
            public void Dispose()
            {
                Operations.Dispose();
                string boundary = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "git-branches")) + Path.DirectorySeparatorChar;
                if (!Path.GetFullPath(Root).StartsWith(boundary, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Fixture cleanup escaped its own directory");
                if (!Directory.Exists(Root)) return;
                foreach (string file in Directory.GetFiles(Root, "*", SearchOption.AllDirectories)) File.SetAttributes(file, FileAttributes.Normal);
                Directory.Delete(Root, true);
            }
        }
    }
}
