namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolsGitTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static Response Response(string value) { return Json.Deserialize<Response>(value); }
        private static async Task<IDictionary<string,object>> Status(LlmVbeTools tools)
        {
            var response=Response(await tools.InvokeAsync("git_status","{\"Project\":\"P\"}"));
            Assert.IsTrue(response.Ok,response.Error);return (IDictionary<string,object>)response.Data;
        }
        private static async Task Rejected(LlmVbeTools tools,string name,object arguments,string fragment)
        {
            var response=Response(await tools.InvokeAsync(name,arguments is string ? (string)arguments : Json.Serialize(arguments)));
            Assert.IsFalse(response.Ok,name);StringAssert.Contains(response.Error,fragment);
        }

        [TestMethod]
        public async Task GitValidationMatrixRejectsMissingWrongTypedWhitespaceExtraFieldsAndCrossDocumentScope()
        {
            var tools=new LlmVbeTools(null,null,new LlmSettings { VbeEditApproval="Automatic" }) { BoundProject="P" };
            await Rejected(tools,"git_unknown","{}","Unknown Git tool");
            await Rejected(tools,"git_status","[]","object");
            await Rejected(tools,"git_status",new {Project="P", Extra=1},"Unexpected Git argument");
            foreach(object arguments in new object[] { "{}", new { Project=2 },new { Project=" " } })
                await Rejected(tools,"git_status",arguments,"Project is required");
            tools.BoundProject=null;
            await Rejected(tools,"git_status",new { Project="P" },"conversation");
            tools.BoundProject="P";
            await Rejected(tools,"git_status",new { Project="other" },"conversation");
            await Rejected(tools,"git_commit",new { Project="P",ExpectedState="state" },"Text is required");
            var settings=new LlmSettings { VbeEditApproval="ReadOnly" };
            tools=new LlmVbeTools(null,null,settings) { BoundProject="P" };
            await Rejected(tools,"git_merge_abort",new {Project="P",ExpectedState="state"},UiText.Get("The VBE policy does not allow this Git operation."));
            settings.VbeEditApproval="invalid";
            await Rejected(tools,"git_merge_abort",new {Project="P",ExpectedState="state"},UiText.Get("The VBE policy does not allow this Git operation."));
            settings.VbeEditApproval="AskEachTime";
            tools.ConfirmGit=(owner,text,title)=>DialogResult.No;
            await Rejected(tools,"git_merge_abort",new {Project="P",ExpectedState="state"},UiText.Get("Git operation declined by the user."));
        }

        [TestMethod]
        public async Task GitStatusHistoryCommitReadMutationAndSelectiveCommitUseDisposableBareRepository()
        {
            using(var fixture=new Fixture())
            {
                var tools=fixture.Tools;
                foreach(string name in new[] {"git_status","git_history","git_branches","git_checkpoints","git_conflicts"})
                {
                    var response=Response(await tools.InvokeAsync(name,"{\"Project\":\"P\"}"));
                    Assert.IsTrue(response.Ok,name+": "+response.Error);
                    Assert.AreEqual(fixture.Initial,((IDictionary<string,object>)response.Data)["Head"]);
                }
                var read=Response(await tools.InvokeAsync("git_commit_read",Json.Serialize(new {Project="P",Name=fixture.Initial})));
                Assert.IsTrue(read.Ok,read.Error);
                var data=(IDictionary<string,object>)read.Data;
                Assert.IsNotNull(data["Details"]); Assert.AreEqual(1,((object[])data["Modules"]).Length);
                string tree=fixture.Git("--git-dir="+Path.Combine(fixture.Root,"cache.git"),"hash-object","-t","tree","-w","--stdin");
                string empty=fixture.Git("--git-dir="+Path.Combine(fixture.Root,"cache.git"),"commit-tree",tree,"-m","Empty fixture commit");
                read=Response(await tools.InvokeAsync("git_commit_read",Json.Serialize(new {Project="P",Name=empty})));
                Assert.IsTrue(read.Ok,read.Error); Assert.IsNull(((IDictionary<string,object>)read.Data)["Modules"]);
                await Rejected(tools,"git_conflict_read",new {Project="P",Path="Module1.bas"},UiText.Get("The branch changed since the merge started."));
                string state=(string)(await Status(tools))["State"];
                var created=Response(await tools.InvokeAsync("git_branch_create",Json.Serialize(new {Project="P",ExpectedState=state,Name="fixture-branch"})));
                Assert.IsTrue(created.Ok,created.Error);
                Assert.AreEqual(fixture.Initial,fixture.Repository.Resolve("refs/heads/fixture-branch"));
                fixture.Settings.VbeEditApproval="AskEachTime";
                tools.ConfirmGit=(owner,text,title)=> { StringAssert.Contains(text,"git_merge_abort"); return DialogResult.Yes; };
                state=(string)(await Status(tools))["State"];
                var abort=Response(await tools.InvokeAsync("git_merge_abort",Json.Serialize(new {Project="P",ExpectedState=state})));
                Assert.IsTrue(abort.Ok,abort.Error);
                fixture.Settings.VbeEditApproval="Automatic";
                await Rejected(tools,"git_commit_selected",new {Project="P",ExpectedState=state,Name="Module1",Text="fixture",Choice="invalid"},"Choice");
                foreach(string choice in new[] {"modules","references"})
                {
                    state=(string)(await Status(tools))["State"];
                    var selected=Response(await tools.InvokeAsync("git_commit_selected",Json.Serialize(new {Project="P",ExpectedState=state,Name="Module1, ,",Text="fixture",Choice=choice})));
                    Assert.IsTrue(selected.Ok,selected.Error);
                    var selectedData=(IDictionary<string,object>)selected.Data;
                    Assert.AreEqual(fixture.Initial,selectedData["Commit"]);
                    Assert.AreEqual(false,selectedData["Published"]);
                }
                fixture.Git("--git-dir="+Path.Combine(fixture.Root,"cache.git"),"config","remote.origin.url","https://github.com/fixture/repository.git");
                var http=new LlmHttpFixture();http.Replies.Enqueue(new LlmHttpFixture.Reply("[{\"number\":7,\"title\":\"Fixture PR\"}]") );
                tools.GitHubApiFactory=account=>new GitHubApi(account,http,token=>Task.FromResult("fixture-token"));
                var pulls=Response(await tools.InvokeAsync("git_pull_requests","{\"Project\":\"P\"}"));
                Assert.IsTrue(pulls.Ok,pulls.Error);
                Assert.AreEqual(7,((IDictionary<string,object>)((object[])pulls.Data)[0])["number"]);
                Assert.AreEqual(1,http.Uris.Count);
            }
        }

        [TestMethod]
        public async Task DefaultGitOpeningUsesSessionDocumentAndRequiresExistingBinding()
        {
            using(var scope=new LlmBoundaryScope())
            {
                var prior=MacroGitOperations.CacheDirectory;
                try
                {
                    MacroGitOperations.CacheDirectory=s=>Path.Combine(scope.Root,"cache");
                    var vbe=new ToolGitVbe();
                    vbe.VBProjects.Add(new ToolGitProject { Name="P",FileName=Path.Combine(scope.Root,"host.xlsm") });
                    var tools=new LlmVbeTools(new VbeSession(vbe),null,new LlmSettings()) { BoundProject="P" };
                    await Rejected(tools,"git_status",new {Project="P"},"GitHub");
                    string cache=Path.Combine(scope.Root,"cache");Directory.CreateDirectory(cache);
                    File.WriteAllText(Path.Combine(cache,"binding.json"),Json.Serialize(new {Remote="https://github.com/fixture/repository.git",Branch="main"}));
                    var response=Response(await tools.InvokeAsync("git_status",Json.Serialize(new {Project="P"})));
                    Assert.IsTrue(response.Ok,response.Error);
                    Assert.IsTrue(Directory.GetFiles(cache,"HEAD",SearchOption.AllDirectories).Length==1);
                }
                finally { MacroGitOperations.CacheDirectory=prior; }
            }
        }
    }
}
