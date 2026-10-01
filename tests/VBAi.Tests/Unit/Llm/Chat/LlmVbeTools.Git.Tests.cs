namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie la validation et les opérations Git exposées par les outils LLM.</summary>
[TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolsGitTests
    {
        /// <summary>Sérialiseur utilisé pour décoder les réponses d’outil.</summary>
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        /// <summary>Désérialise une réponse d’outil.</summary>
        /// <param name="value">Réponse JSON.</param><returns>Objet réponse décodé.</returns>
        private static Response Response(string value) { return Json.Deserialize<Response>(value); }
        /// <summary>Demande le statut Git du projet puis renvoie ses données.</summary>
        /// <param name="tools">Orchestrateur à invoquer.</param><returns>Données du statut Git.</returns>
        private static async Task<IDictionary<string,object>> Status(LlmVbeTools tools)
        {
            var response=Response(await tools.InvokeAsync("git_status","{\"Project\":\"P\"}"));
            Assert.IsTrue(response.Ok,response.Error);return (IDictionary<string,object>)response.Data;
        }
        /// <summary>Vérifie qu’un appel Git est refusé avec le fragment d’erreur attendu.</summary>
        /// <param name="tools">Orchestrateur.</param><param name="name">Nom de l’outil.</param><param name="arguments">Arguments à transmettre.</param><param name="fragment">Fragment attendu dans l’erreur.</param>
        /// <returns>Tâche terminée lorsque l’appel est refusé comme attendu.</returns>
        private static async Task Rejected(LlmVbeTools tools,string name,object arguments,string fragment)
        {
            var response=Response(await tools.InvokeAsync(name,arguments is string ? (string)arguments : Json.Serialize(arguments)));
            Assert.IsFalse(response.Ok,name);StringAssert.Contains(response.Error,fragment);
        }

        /// <summary>Refuse les arguments absents, mal typés, supplémentaires ou liés à un autre scope.</summary>
        /// <returns>Tâche terminée lorsque tous les arguments invalides sont refusés.</returns>
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

        /// <summary>Exerce lecture, branches, commits, mutations et pull requests sur un dépôt Git temporaire.</summary>
        /// <returns>Tâche terminée lorsque les opérations du dépôt satisfont les assertions.</returns>
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

        /// <summary>Ouvre le dépôt par défaut depuis le document de session uniquement lorsque sa liaison existe.</summary>
        /// <returns>Tâche terminée lorsque l’ouverture respecte la liaison existante.</returns>
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

        [DataTestMethod, DataRow("MixedCase\\Classeur.xlsm", false), DataRow("MixedCase\\Classeur.xlsm", true)]
        [DataRow("Développement\\ClasseurÉté.xlsm", false), DataRow("Développement\\ClasseurÉté.xlsm", true)]
        public async Task DefaultOpenGitReadsAndMutatesTheExistingUiCacheWithoutFactoryOrCaseRekeying(string suffix, bool legacy)
        {
            using (var f = new Fixture())
            {
                var priorOperations = MacroGitOperations.CacheDirectory; var priorUi = GitWindow.CacheDirectory;
                try
                {
                    string nativeScope = Path.GetFullPath(Path.Combine(f.Root, suffix));
                    Func<string, string> rawCache = key => Path.Combine(f.Root, "scope-cache", Path.GetFileName(MacroGitRepository.ScopeDirectory(key)));
                    var lookups = new List<string>();
                    GitWindow.CacheDirectory = key => MacroGitRepository.ResolveScopeDirectory(key, rawCache, File.GetAttributes);
                    MacroGitOperations.CacheDirectory = key => { lookups.Add(key); return GitWindow.CacheDirectory(key); };
                    string boundCache = rawCache(legacy ? nativeScope.ToUpperInvariant() : nativeScope);
                    string alternate = rawCache(legacy ? nativeScope : nativeScope.ToUpperInvariant());
                    Directory.CreateDirectory(boundCache);
                    byte[] binding = System.Text.Encoding.UTF8.GetBytes(Json.Serialize(new { Remote = "https://github.com/fixture/repository.git", Branch = "main" }));
                    File.WriteAllBytes(Path.Combine(boundCache, "binding.json"), binding);
                    Assert.AreEqual(boundCache, GitWindow.CacheDirectory(nativeScope), "The UI and tools must find the same persisted binding.");
                    var vbe = new ToolGitVbe(); var host = new ToolGitProject { Name = "P", FileName = nativeScope };
                    host.VBComponents.Add(new global::FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\n")); vbe.VBProjects.Add(host);
                    var tools = new LlmVbeTools(new VbeSession(vbe), null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "P" };
                    Assert.IsNull(tools.GitOperationsFactory);
                    string state = (string)(await Status(tools))["State"];
                    var result = Response(await tools.InvokeAsync("git_pr_prepare", Json.Serialize(new { Project = "P", ExpectedState = state,
                        Name = "main", Text = "Local disposable draft", Choice = "Synthetic body; never publish" })));
                    Assert.IsTrue(result.Ok, result.Error);
                    CollectionAssert.AreEqual(new[] { nativeScope, nativeScope }, lookups.ToArray());
                    Assert.IsFalse(Directory.Exists(alternate), "Lookup must not create or migrate an alternate cache.");
                    CollectionAssert.AreEqual(binding, File.ReadAllBytes(Path.Combine(boundCache, "binding.json")));
                    string draftPath = Directory.GetFiles(boundCache, "codex-pr-draft.json", SearchOption.AllDirectories).Single();
                    Assert.AreEqual("Local disposable draft", Json.Deserialize<GitPullDraft>(File.ReadAllText(draftPath)).Title);
                    Assert.AreEqual(0, host.VBComponents.ImportAttempts); Assert.AreEqual(1, host.VBComponents.Cast<object>().Count());
                }
                finally { MacroGitOperations.CacheDirectory = priorOperations; GitWindow.CacheDirectory = priorUi; }
            }
        }

        [DataTestMethod, DataRow("missing"), DataRow("ambiguous"), DataRow("directory")]
        public async Task DefaultOpenGitRefusesMissingAmbiguousOrInvalidBindingWithoutCacheOrSourceMutation(string state)
        {
            using (var f = new Fixture())
            {
                var prior = MacroGitOperations.CacheDirectory;
                try
                {
                    string nativeScope = Path.GetFullPath(Path.Combine(f.Root, "MixedCase", "ClasseurÉté.xlsm"));
                    Func<string, string> rawCache = key => Path.Combine(f.Root, "scope-cache", Path.GetFileName(MacroGitRepository.ScopeDirectory(key)));
                    MacroGitOperations.CacheDirectory = key => MacroGitRepository.ResolveScopeDirectory(key, rawCache, File.GetAttributes);
                    string exact = rawCache(nativeScope), legacy = rawCache(nativeScope.ToUpperInvariant());
                    foreach (string cache in state == "missing" ? new string[0] : new[] { exact, legacy })
                    {
                        Directory.CreateDirectory(cache);
                        if (state == "directory" && cache == exact) Directory.CreateDirectory(Path.Combine(cache, "binding.json"));
                        else File.WriteAllText(Path.Combine(cache, "binding.json"), Json.Serialize(new { Remote = "https://github.com/fixture/repository.git", Branch = "main" }));
                    }
                    var beforeFiles = Directory.GetFiles(f.Root, "*", SearchOption.AllDirectories).ToDictionary(path => path, File.ReadAllBytes);
                    var beforeDirectories = Directory.GetDirectories(f.Root, "*", SearchOption.AllDirectories);
                    var vbe = new ToolGitVbe(); var host = new ToolGitProject { Name = "P", FileName = nativeScope }; vbe.VBProjects.Add(host);
                    var tools = new LlmVbeTools(new VbeSession(vbe), null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "P" };
                    string fragment = state == "missing" ? "GitHub" : state == "ambiguous" ? "both" : "not a regular file";
                    await Rejected(tools, "git_status", new { Project = "P" }, fragment);
                    await Rejected(tools, "git_pr_prepare", new { Project = "P", ExpectedState = "unopened", Name = "main", Text = "Refused draft", Choice = "Synthetic" }, fragment);
                    CollectionAssert.AreEquivalent(beforeDirectories, Directory.GetDirectories(f.Root, "*", SearchOption.AllDirectories));
                    CollectionAssert.AreEquivalent(beforeFiles.Keys.ToArray(), Directory.GetFiles(f.Root, "*", SearchOption.AllDirectories));
                    foreach (var file in beforeFiles) CollectionAssert.AreEqual(file.Value, File.ReadAllBytes(file.Key));
                    Assert.AreEqual(0, host.VBComponents.ImportAttempts);
                }
                finally { MacroGitOperations.CacheDirectory = prior; }
            }
        }

        [TestMethod]
        public async Task DefaultOpenGitPreservesConversationGuardAndLiveDocumentIdentityAcrossCacheLookup()
        {
            using (var f = new Fixture())
            {
                var prior = MacroGitOperations.CacheDirectory;
                try
                {
                    string nativeScope = Path.Combine(f.Root, "MixedCaseÉté.xlsm"), cache = Path.Combine(f.Root, "owned-bound-cache");
                    Directory.CreateDirectory(cache); File.WriteAllText(Path.Combine(cache, "binding.json"), Json.Serialize(new { Remote = "https://github.com/fixture/repository.git", Branch = "main" }));
                    var vbe = new ToolGitVbe(); var host = new ToolGitProject { Name = "P", FileName = nativeScope }; vbe.VBProjects.Add(host); int lookups = 0;
                    MacroGitOperations.CacheDirectory = key => { lookups++; Assert.AreEqual(nativeScope, key); host.FileName = Path.Combine(f.Root, "Other.xlsm"); return cache; };
                    var tools = new LlmVbeTools(new VbeSession(vbe), null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "P" };
                    await Rejected(tools, "git_status", new { Project = "Other" }, "conversation"); Assert.AreEqual(0, lookups);
                    await Rejected(tools, "git_status", new { Project = "P" }, UiText.Get("The linked document changed. Reopen GitHub integration."));
                    Assert.AreEqual(1, lookups); Assert.AreEqual(0, host.VBComponents.ImportAttempts);
                    Assert.AreEqual(0, Directory.GetFiles(cache, "codex-pr-draft.json", SearchOption.AllDirectories).Length);
                }
                finally { MacroGitOperations.CacheDirectory = prior; }
            }
        }
    }
}
