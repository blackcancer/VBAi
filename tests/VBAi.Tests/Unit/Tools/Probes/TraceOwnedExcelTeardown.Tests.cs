using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class TraceOwnedExcelTeardownTests
    {
        [DataTestMethod]
        [DataRow("Valid")]
        [DataRow("WrongNonce")]
        [DataRow("WrongScenario")]
        [DataRow("WrongCandidate")]
        [DataRow("ExecuteWithoutPreflight")]
        public void PrepareAndRefusalPathsNeverLookupHostOrStartDebugger(string variation)
        {
            string root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string assembly = Path.Combine(root, "VBAi.dll"); File.WriteAllText(assembly, "Synthetic identity; not an assembly", new UTF8Encoding(false));
                string hash; using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(assembly))).Replace("-", "");
                string mvid = Guid.NewGuid().ToString("D");
                var pending = new Dictionary<string, object> { ["Root"] = root, ["ProcessId"] = 424242,
                    ["ProcessStartedUtc"] = DateTime.UtcNow.ToString("o"), ["Nonce"] = variation == "WrongNonce" ? "stale" : Guid.NewGuid().ToString("N"),
                    ["Executable"] = @"C:\Owned\EXCEL.EXE", ["AssemblyPath"] = assembly,
                    ["AssemblyMvid"] = variation == "WrongCandidate" ? Guid.NewGuid().ToString("D") : mvid,
                    ["Scenario"] = variation == "WrongScenario" ? "UserMacro" : "NativeVariantArraysRoundTripWithBoundsAndOneInvocation" };
                string path = Path.Combine(root, "teardown.pending.json");
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(pending), new UTF8Encoding(false));
                string controller = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "probes", "Trace-OwnedExcelTeardown.ps1");
                Assert.IsTrue(File.Exists(controller), "Exact controller fixture must be packaged with the tests.");
                string command = "$ErrorActionPreference='Stop';[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false); & " + Quote(controller) +
                    " -PendingReport " + Quote(path) + " -ExpectedMvid " + Quote(mvid) + " -ExpectedAssemblySha256 " + Quote(hash) +
                    " -CdbPath " + Quote(assembly) + (variation == "ExecuteWithoutPreflight" ? " -Execute" : "");
                string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
                var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command))) {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
                info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
                using (var child = Process.Start(info))
                {
                    var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                    Assert.IsTrue(child.WaitForExit(10000), "Prepare-only child timeout; preserve evidence, never kill host/debugger.");
                    Assert.AreEqual(variation == "Valid", child.ExitCode == 0, output.Result + error.Result);
                    if (variation == "Valid")
                    {
                        var plan = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(output.Result);
                        Assert.AreEqual("PREPARE_ONLY", plan["Mode"]); Assert.AreEqual(0, plan["CleanupCallsIssuedByController"]);
                        Assert.AreEqual("LiveEventThread", plan["RegisterCaptureMode"]);
                        Assert.AreEqual(0, plan["HostTerminationCalls"]); Assert.AreEqual(0, plan["MemoryDumps"]);
                    }
                }
                Assert.AreEqual(2, Directory.GetFiles(root).Length, "No permission, debugger output or extra file may be written.");
            }
            finally { Directory.Delete(root, true); }
        }
        private static string Quote(string text) => "'" + text.Replace("'", "''") + "'";

        [TestMethod]
        public void FirstAndSecondChanceUseTheSameBoundedExceptionCollector()
        {
            string common = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            // Execute only the pure command generator, never start a debugger or target.
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";$commands=Get-TeardownCommands 'C:\\Owned\\trace.log' 424242 '0123456789abcdef0123456789abcdef';" +
                "$handler=[regex]::Match($commands,'(?m)^sxe -c \"([^\"]+)\" -c2 \"([^\"]+)\" 0xc0000409$');" +
                "if(-not $handler.Success){throw 'Both exception chances must be armed'};" +
                "if($handler.Groups[1].Value -cne $handler.Groups[2].Value){throw 'Chance collectors differ'};" +
                "if($handler.Groups[1].Value -cne '.echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .lastevent; .exr -1; .echo VBAI_TEARDOWN_REGISTER_MODE LiveEventThread; r; kv; .echo VBAI_TEARDOWN_EXCEPTION_END; gn'){throw 'Collector must preserve live-event register and unhandled fatal semantics'};" +
                "if($commands -notmatch '(?m)^sxn -c \"qd\" epr$'){throw 'Debugger may quit automatically only after process exit'};" +
                "if($commands -notmatch '(?m)^\\.echo VBAI_TEARDOWN_READY 424242 0123456789abcdef0123456789abcdef$'){throw 'Owned readiness identity missing'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command))) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }

        [TestMethod]
        public void OnlyExecutedExactReadinessCanArmTheOwnedDebugger()
        {
            string common = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            string nonce = "0123456789abcdef0123456789abcdef";
            string line = "VBAI_TEARDOWN_READY 424242 " + nonce;
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";$nonce=" + Quote(nonce) + ";" +
                "if(-not (Test-OwnedTeardownReady " + Quote(line + "\r\n") + " 424242 $nonce)){throw 'Executed exact readiness refused'};" +
                "if(Test-OwnedTeardownReady " + Quote("0:006> .echo " + line + "\r\n") + " 424242 $nonce){throw 'Readiness command echo accepted'};" +
                "if(Test-OwnedTeardownReady " + Quote(line) + " 424243 $nonce){throw 'Different PID accepted'};" +
                "if(Test-OwnedTeardownReady " + Quote(line) + " 424242 'fedcba9876543210fedcba9876543210'){throw 'Different nonce accepted'};" +
                "if(Test-OwnedTeardownReady " + Quote(line + " trailing") + " 424242 $nonce){throw 'Readiness suffix accepted'};" +
                "if(Test-OwnedTeardownReady " + Quote(line) + " 424242 '.*'){throw 'Regex nonce accepted'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command))) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }

        [TestMethod]
        public void ExecutedExceptionRecordAndStackAreRequiredInsteadOfEchoedDebuggerCommands()
        {
            string common = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            string genuine = "VBAI_TEARDOWN_EXCEPTION_BEGIN\r\nExceptionCode: c0000409\r\nrip=0000000000000010 rsp=0000000000000020\r\nChild-SP RetAddr Call Site\r\nVBAI_TEARDOWN_EXCEPTION_END\r\n";
            string echo = "0:000> sxe -c \".echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .exr -1; kv; .echo VBAI_TEARDOWN_EXCEPTION_END\" 0xc0000409";
            string missingContext = genuine.Replace("rip=0000000000000010 rsp=0000000000000020", "Unable to get exception context, HRESULT 0x8000FFFF");
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";if(-not (Test-TeardownExceptionCapture " + Quote(genuine) + ")){throw 'Executed exception missing'};" +
                "if(Test-TeardownExceptionCapture " + Quote(echo) + "){throw 'Command echo accepted'};" +
                "if(Test-TeardownExceptionCapture " + Quote(missingContext) + "){throw 'Incomplete exception context accepted'};" +
                "$valid=" + Quote(genuine) + ";$incomplete=" + Quote(missingContext) + ";" +
                "if(-not (Test-OwnedTeardownFatalExit -1073740791 $true -1073740791 $valid)){throw 'Exact terminal fatal propagation refused'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $false -1073740791 $valid){throw 'Live target accepted'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true $null $valid){throw 'Unknown target exit accepted'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true 0 $valid){throw 'Different target exit accepted'};" +
                "if(Test-OwnedTeardownFatalExit 5 $true 5 $valid){throw 'Generic matching nonzero accepted'};" +
                "if(Test-OwnedTeardownFatalExit 0 $true -1073740791 $valid){throw 'Zero debugger exit called fatal propagation'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true -1073740791 $incomplete){throw 'Missing registers accepted for fatal propagation'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command))) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }
    }
}
