namespace VBAi.Tests.Infrastructure
{
    using Microsoft.CSharp;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.CodeDom.Compiler;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using VBAi;

    internal sealed class AccountProcessScope : IDisposable
    {
        private readonly LlmBoundaryScope scope = new LlmBoundaryScope();
        private readonly Func<ProcessStartInfo, Process> start = CodexAccount.StartProcess;
        private readonly Func<Process, int, bool> wait = CodexAccount.WaitForExit;
        private readonly Func<string, bool> exists = CodexAccount.FileExists;
        private readonly Func<string, string[]> directories = CodexAccount.GetDirectories;
        private readonly Func<string, DateTime> lastWrite = CodexAccount.GetLastWriteTimeUtc;
        private readonly string priorMode = Environment.GetEnvironmentVariable("VBAi_TEST_ACCOUNT_MODE");
        private readonly string priorMarker = Environment.GetEnvironmentVariable("VBAi_TEST_ACCOUNT_MARKER");
        private readonly List<Process> logins = new List<Process>();
        internal readonly string Executable;
        internal readonly string Marker;
        internal readonly List<string> Commands = new List<string>();
        internal AccountProcessScope()
        {
            Executable = Path.Combine(scope.Root, "account-fixture.exe"); Marker = Path.Combine(scope.Root, "login.marker");
            using (var compiler = new CSharpCodeProvider())
            {
                var options = new CompilerParameters(new[] { "System.dll", "System.Core.dll" }, Executable)
                { GenerateExecutable = true, GenerateInMemory = false, CompilerOptions = "/target:winexe /optimize+" };
                var result = compiler.CompileAssemblyFromSource(options, Program);
                Assert.IsFalse(result.Errors.HasErrors, string.Join("\n", result.Errors.Cast<CompilerError>().Select(e => e.ToString())));
            }
            Environment.SetEnvironmentVariable("VBAi_CODEX_CLI", Executable);
            Environment.SetEnvironmentVariable("VBAi_TEST_ACCOUNT_MARKER", Marker);
        }
        internal void Mode(string mode) { Environment.SetEnvironmentVariable("VBAi_TEST_ACCOUNT_MODE", mode); }
        internal void StartGit(Process process, string mode)
        {
            Commands.Add(process.StartInfo.Arguments);
            Assert.AreEqual("git.exe", process.StartInfo.FileName);
            Assert.IsFalse(process.StartInfo.UseShellExecute); Assert.IsTrue(process.StartInfo.CreateNoWindow);
            Assert.IsTrue(process.StartInfo.RedirectStandardOutput); Assert.IsTrue(process.StartInfo.RedirectStandardError);
            process.StartInfo.FileName = Executable; process.StartInfo.Arguments = mode; process.Start();
        }
        internal Process StartLogin(ProcessStartInfo info)
        {
            Assert.AreEqual(Executable, info.FileName); Assert.AreEqual("login", info.Arguments); Assert.IsFalse(info.UseShellExecute); Assert.IsFalse(info.CreateNoWindow); Assert.AreEqual(ProviderSessionStorage.CodexHome, info.EnvironmentVariables["CODEX_HOME"]);
            var process = Process.Start(info); logins.Add(process); return process;
        }
        internal void AssertLoginFinished()
        {
            foreach (var process in logins) Assert.IsTrue(process.WaitForExit(5000));
            Assert.IsTrue(File.Exists(Marker)); Assert.AreEqual("fixture login", File.ReadAllText(Marker));
        }
        public void Dispose()
        {
            foreach (var process in logins) { if (!process.HasExited) { process.Kill(); process.WaitForExit(5000); } process.Dispose(); }
            CodexAccount.StartProcess = start; CodexAccount.WaitForExit = wait; CodexAccount.FileExists = exists;
            CodexAccount.GetDirectories = directories; CodexAccount.GetLastWriteTimeUtc = lastWrite;
            Environment.SetEnvironmentVariable("VBAi_TEST_ACCOUNT_MODE", priorMode);
            Environment.SetEnvironmentVariable("VBAi_TEST_ACCOUNT_MARKER", priorMarker);
            scope.Dispose();
        }
        private const string Program = @"
using System;
using System.IO;
using System.Threading;
static class AccountFixtureProgram
{
    static int Main(string[] arguments)
    {
        if(arguments.Length==1 && arguments[0]==""login"")
        {File.WriteAllText(Environment.GetEnvironmentVariable(""VBAi_TEST_ACCOUNT_MARKER""),""fixture login"");return 0;}
        string mode=arguments.Length==1 ? arguments[0] : Environment.GetEnvironmentVariable(""VBAi_TEST_ACCOUNT_MODE"");
        using(var output=new StreamWriter(Console.OpenStandardOutput()) {AutoFlush=true})
        using(var error=new StreamWriter(Console.OpenStandardError()) {AutoFlush=true})
        {
            if(mode==""accounts""){output.Write(""zeta\r\nalice\r\nALICE\r\n"");return 0;}
            if(mode==""error""){error.Write(""fixture diagnostic body"");return 7;}
            if(mode==""error-empty"")return 7;
            if(mode==""chatgpt""){error.Write(""Logged in using cHaTgPt"");return 0;}
            if(mode==""api""){output.Write(""API account fixture"");return 0;}
            if(mode==""volume""){output.Write(new string('x',131072));error.Write(new string('y',131072));return 0;}
            if(mode==""hang""){Thread.Sleep(30000);return 0;}
            return 0;
        }
    }
}";
    }
}
