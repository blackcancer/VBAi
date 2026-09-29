using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Infrastructure
{
    internal sealed class NativeProtocolFixture : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexProtocol-" + Guid.NewGuid().ToString("N"));
        internal readonly string Executable;
        internal NativeProtocolFixture()
        {
            Directory.CreateDirectory(Root);
            Executable = Path.Combine(Root, "protocol.exe");
            using (var compiler = new CSharpCodeProvider())
            {
                var options = new CompilerParameters(new[] { "System.dll" }, Executable) { GenerateExecutable = true, CompilerOptions = "/target:winexe" };
                var result = compiler.CompileAssemblyFromSource(options, @"
using System;
using System.Threading;
using System.IO;
class Program {
 static void Main(string[] args) {
  if(args[0]==""echo"") { Console.Error.WriteLine(""diagnostic only""); string line; while((line=Console.ReadLine())!=null){if(line==""exit"")return;Console.WriteLine(line);}return; }
  if(args[0]==""immediate""){Console.WriteLine(""password=fixture-token"");return;}
  string input=Console.In.ReadToEnd();string capture=Environment.GetEnvironmentVariable(""VBAi_FIXTURE_INPUT"");if(capture!=null)File.WriteAllText(capture,input);
  if(args[0]==""wait""){Thread.Sleep(60000);return;}
  if(args[0]==""error""){Environment.Exit(7);return;}
  if(args[0]==""missing""){Console.WriteLine(""username=fixture"");return;}
  Console.WriteLine(""username=fixture"");Console.WriteLine(""password=fixture-token\r"");
 }
}");
                Assert.IsFalse(result.Errors.HasErrors, string.Join("\n", result.Errors.Cast<CompilerError>().Select(x => x.ToString())));
            }
        }
        internal bool Start(Process process, string mode)
        {
            process.StartInfo.FileName = Executable;
            process.StartInfo.Arguments = mode;
            process.StartInfo.EnvironmentVariables["VBAi_FIXTURE_INPUT"] = Path.Combine(Root, "input.txt");
            return VBAi.ProcessInput.StartWithoutPreamble(process);
        }
        public void Dispose()
        {
            for (int attempt = 0; Directory.Exists(Root); attempt++)
            {
                try { Directory.Delete(Root, true); }
                catch (IOException) { if (attempt >= 100) throw; System.Threading.Thread.Sleep(10); }
                catch (UnauthorizedAccessException) { if (attempt >= 100) throw; System.Threading.Thread.Sleep(10); }
            }
        }
    }
}
