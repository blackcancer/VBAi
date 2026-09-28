namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Diagnostics;
    using System.Web.Script.Serialization;
    using CodexVBE;

    public sealed class ToolGitVbe { public System.Collections.Generic.List<ToolGitProject> VBProjects { get; } = new System.Collections.Generic.List<ToolGitProject>(); }
    public sealed class ToolGitProject
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public int Mode { get; set; } = 2;
        public int Protection { get; set; }
        public global::FakeComponents VBComponents { get; } = new global::FakeComponents();
        public object[] References { get; } = new object[0];
    }

    public sealed partial class LlmVbeToolsGitTests
    {
        private sealed class Fixture : IDisposable
        {
            internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexToolGit-" + Guid.NewGuid().ToString("N"));
            internal readonly MacroGitRepository Repository;
            internal readonly VbaGitProject Project;
            internal readonly LlmSettings Settings = new LlmSettings { VbeEditApproval = "Automatic" };
            internal readonly LlmVbeTools Tools;
            internal readonly global::FakeProject Host;
            internal readonly string Initial;
            internal Fixture()
            {
                Directory.CreateDirectory(Root);
                string remote = Path.Combine(Root,"origin.git");
                Git("init","--bare",remote);
                Repository = new MacroGitRepository(Path.Combine(Root,"cache.git"),"main");
                Repository.Initialize(remote);
                Git("--git-dir="+Path.Combine(Root,"cache.git"),"config","user.name","Coverage Fixture");
                Git("--git-dir="+Path.Combine(Root,"cache.git"),"config","user.email","coverage@example.invalid");
                Host = new global::FakeProject { FileName = Path.Combine(Root,"fixture.xlsm") };
                Host.VBComponents.Add(new global::FakeComponent("Module1",1,"Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
                Project = new VbaGitProject(() => Host,Host.FileName);
                Initial = Repository.Commit(Project.Capture(),null,"Initial fixture");
                Repository.SetRef(Repository.Head,Initial); Repository.SetRef(MacroGitRepository.Baseline,Initial);
                Tools = new LlmVbeTools(null,null,Settings) { BoundProject = "P", GitOperationsFactory = p => new MacroGitOperations(Project,Repository) };
            }
            internal string Git(params string[] arguments)
            {
                var start = new ProcessStartInfo("git.exe",string.Join(" ",Array.ConvertAll(arguments,a=>"\""+a.Replace("\"","\\\"")+"\"")))
                { WorkingDirectory=Root,UseShellExecute=false,CreateNoWindow=true,RedirectStandardError=true,RedirectStandardOutput=true,RedirectStandardInput=true };
                using(var process=new Process { StartInfo=start })
                {
                    ProcessInput.StartWithoutPreamble(process);
                    process.StandardInput.Close();
                    var output=process.StandardOutput.ReadToEndAsync(); var error=process.StandardError.ReadToEndAsync();
                    if(!process.WaitForExit(30000)) { process.Kill(); throw new TimeoutException("Fixture git timeout"); }
                    if(process.ExitCode!=0) throw new InvalidOperationException(error.GetAwaiter().GetResult());
                    return output.GetAwaiter().GetResult().Trim();
                }
            }
            public void Dispose()
            {
                if (!Path.GetFullPath(Root).StartsWith(Path.GetFullPath(Path.GetTempPath())+"CodexToolGit-",StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Cleanup escaped fixture root");
                foreach(string file in Directory.GetFiles(Root,"*",SearchOption.AllDirectories)) File.SetAttributes(file,FileAttributes.Normal);
                Directory.Delete(Root,true);
            }
        }
    }
}
