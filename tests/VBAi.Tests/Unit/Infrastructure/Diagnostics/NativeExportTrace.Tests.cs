using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;

namespace VBAi.Tests
{
    /// <summary>Exercises the actual debugger script with synthetic register/memory fixtures; no debugger or Office process.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeExportTraceTests
    {
        [TestMethod]
        public void TraceScriptPreservesPidPathPrivacyAndPairedNativeStatus()
        {
            string source = FindSource();
            var info = new ProcessStartInfo("node.exe", "\"" + source + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using (var process = Process.Start(info))
            {
                var output = process.StandardOutput.ReadToEndAsync();
                var error = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(10000)) { process.Kill(); Assert.Fail("Owned pure Node script exceeded its bound; no native host was used."); }
                Assert.AreEqual(0, process.ExitCode, output.Result + error.Result);
                StringAssert.Contains(output.Result, "NativeExportTrace pure contracts PASS");
            }
        }

        private static string FindSource()
        {
            string candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tests", "VBAi.Tests", "Infrastructure", "Diagnostics", "NativeExportTrace.Tests.mjs");
            if (File.Exists(candidate)) return candidate;
            throw new FileNotFoundException("The matching debugger script regression fixture must be packaged beside this test assembly.", candidate);
        }
    }
}
