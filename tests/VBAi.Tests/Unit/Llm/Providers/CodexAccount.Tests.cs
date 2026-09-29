namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using System.Diagnostics;
    using System.Threading.Tasks;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class CodexAccountTests
    {
        [TestMethod]
        public async Task NativeStatusMatrixReadsChatGptOtherAccountErrorsAndEmptyDiagnostics()
        {
            using(var scope=new AccountProcessScope())
            {
                foreach(string mode in new[] {"chatgpt","api","error","error-empty"})
                {
                    scope.Mode(mode); CodexAccount.StartProcess=info=> {Assert.AreEqual(ProviderSessionStorage.CodexHome,info.EnvironmentVariables["CODEX_HOME"]);return Process.Start(info);};var status=await CodexAccount.ReadStatusAsync();
                    Assert.AreEqual(mode=="chatgpt",status.ChatGptConnected);
                    if(mode=="chatgpt")Assert.AreEqual("Connected to ChatGPT",status.Text);
                    if(mode=="api")StringAssert.Contains(status.Text,"API account fixture");
                    if(mode=="error")StringAssert.Contains(status.Text," : fixture diagnostic body");
                    if(mode=="error-empty")Assert.AreEqual("Not signed in to ChatGPT.",status.Text);
                }
                scope.Mode("volume");
                var large=await CodexAccount.ReadStatusAsync();
                Assert.IsFalse(large.ChatGptConnected);Assert.IsTrue(large.Text.Length>262144);
                StringAssert.Contains(large.Text,new string('x',1000));StringAssert.Contains(large.Text,new string('y',1000));
                scope.Mode("hang");Process observer=null;
                CodexAccount.StartProcess=info=> {var process=Process.Start(info);observer=Process.GetProcessById(process.Id);return process;};
                CodexAccount.WaitForExit=(p,m)=> {Assert.AreEqual(10000,m);return p.WaitForExit(20);};
                try
                {
                    await Assert.ThrowsExceptionAsync<TimeoutException>(()=>CodexAccount.ReadStatusAsync());
                    Assert.IsNotNull(observer);Assert.IsTrue(observer.WaitForExit(5000));Assert.IsTrue(observer.HasExited);
                }
                finally {if(observer!=null) {if(!observer.HasExited) {observer.Kill();observer.WaitForExit(5000);}observer.Dispose();}}
            }
        }

        [TestMethod]
        public void ExecutableResolutionAndLoginUseOnlyTheConfiguredDisposableCli()
        {
            using(var scope=new AccountProcessScope())
            {
                Assert.AreEqual(scope.Executable,CodexAccount.Executable);
                Environment.SetEnvironmentVariable("VBAi_CODEX_CLI",null);
                CodexAccount.GetDirectories=path=>new string[0];
                CodexAccount.FileExists=path=>false;Assert.AreEqual("codex.exe",CodexAccount.Executable);
                string installed=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Programs","OpenAI","Codex","bin","codex.exe");
                CodexAccount.FileExists=path=> {Assert.AreEqual(installed,path);return true;};Assert.AreEqual(installed,CodexAccount.Executable);
                string desktop=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"OpenAI","Codex","bin");
                string older=Path.Combine(desktop,"older");string newer=Path.Combine(desktop,"newer");
                CodexAccount.GetDirectories=path=> {Assert.AreEqual(desktop,path);return new[] {older,newer};};
                CodexAccount.FileExists=path=>path==Path.Combine(older,"codex.exe")||path==Path.Combine(newer,"codex.exe");
                CodexAccount.GetLastWriteTimeUtc=path=>path==Path.Combine(newer,"codex.exe")?new DateTime(2026,9,28):new DateTime(2026,9,27);
                Assert.AreEqual(Path.Combine(newer,"codex.exe"),CodexAccount.Executable);
                Environment.SetEnvironmentVariable("VBAi_CODEX_CLI",scope.Executable);
                CodexAccount.StartProcess=scope.StartLogin;CodexAccount.StartLogin();scope.AssertLoginFinished();
            }
        }
    }
}
