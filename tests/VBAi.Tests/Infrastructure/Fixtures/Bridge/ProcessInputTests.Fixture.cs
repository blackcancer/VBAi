namespace VBAi.Tests.Unit
{
    using Microsoft.CSharp;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.CodeDom.Compiler;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using VBAi.Tests.Infrastructure;

    public sealed partial class ProcessInputTests
    {
        private static string Compile(LlmBoundaryScope scope)
        {
            string executable = Path.Combine(scope.Root, "input-bytes.exe");
            using (var compiler = new CSharpCodeProvider())
            {
                var options = new CompilerParameters(new[] { "System.dll", "System.Core.dll" }, executable)
                { GenerateExecutable = true, GenerateInMemory = false, CompilerOptions = "/target:winexe /optimize+" };
                var result = compiler.CompileAssemblyFromSource(options, @"using System;using System.IO;using System.Text;
static class Bytes {static void Main() {using(var memory=new MemoryStream()) {Console.OpenStandardInput().CopyTo(memory);var bytes=Encoding.ASCII.GetBytes(BitConverter.ToString(memory.ToArray()));Console.OpenStandardOutput().Write(bytes,0,bytes.Length);}}}");
                Assert.IsFalse(result.Errors.HasErrors, string.Join("\n", result.Errors.Cast<CompilerError>().Select(e => e.ToString())));
            }
            return executable;
        }
        private static Process Child(string executable)
        { return new Process { StartInfo = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true } }; }
    }
}
