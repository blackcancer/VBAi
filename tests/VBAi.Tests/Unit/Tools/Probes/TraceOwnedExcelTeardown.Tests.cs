using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class TraceOwnedExcelTeardownTests
    {
        [DataTestMethod]
        [DataRow("Valid")]
        [DataRow("ValidParamArray")]
        [DataRow("ValidPersistence")]
        [DataRow("ValidUnsigned")]
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
                var pending = new Dictionary<string, object>
                {
                    ["Root"] = root,
                    ["ProcessId"] = 424242,
                    ["ProcessStartedUtc"] = DateTime.UtcNow.ToString("o"),
                    ["Nonce"] = variation == "WrongNonce" ? "stale" : Guid.NewGuid().ToString("N"),
                    ["Executable"] = @"C:\Owned\EXCEL.EXE",
                    ["AssemblyPath"] = assembly,
                    ["AssemblyMvid"] = variation == "WrongCandidate" ? Guid.NewGuid().ToString("D") : mvid,
                    ["Scenario"] = variation == "WrongScenario" ? "UserMacro" : variation == "ValidParamArray"
                        ? "NativeParamArrayCallsPreserveArityValuesAndSingleInvocation" : variation == "ValidPersistence"
                        ? "ProductSavePreservesModuleClassAndFormInFreshOwnedProcessWithoutHelperSaving" : variation == "ValidUnsigned"
                        ? "SavedUnsignedWorkbookHasNoSignatureAndPreservesSource" : "NativeVariantArraysRoundTripWithBoundsAndOneInvocation"
                };
                string path = Path.Combine(root, "teardown.pending.json");
                File.WriteAllText(path, new JavaScriptSerializer().Serialize(pending), new UTF8Encoding(false));
                string controller = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "Trace-OwnedExcelTeardown.ps1");
                Assert.IsTrue(File.Exists(controller), "Exact controller fixture must be packaged with the tests.");
                string command = "$ErrorActionPreference='Stop';[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false); & " + Quote(controller) +
                    " -PendingReport " + Quote(path) + " -ExpectedMvid " + Quote(mvid) + " -ExpectedAssemblySha256 " + Quote(hash) +
                    " -CdbPath " + Quote(assembly) + (variation == "ExecuteWithoutPreflight" ? " -Execute" : "");
                string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
                var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
                using (var child = Process.Start(info))
                {
                    var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                    Assert.IsTrue(child.WaitForExit(10000), "Prepare-only child timeout; preserve evidence, never kill host/debugger.");
                    bool valid = variation.StartsWith("Valid", StringComparison.Ordinal);
                    Assert.AreEqual(valid, child.ExitCode == 0, output.Result + error.Result);
                    if (valid)
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
            string common = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            // Execute only the pure command generator, never start a debugger or target.
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";$commands=Get-TeardownCommands 'C:\\Owned\\trace.log' 424242 '0123456789abcdef0123456789abcdef';" +
                "$handler=[regex]::Match($commands,'(?m)^sxe -c \"([^\"]+)\" -c2 \"([^\"]+)\" 0xc0000409$');" +
                "if(-not $handler.Success){throw 'Both exception chances must be armed'};" +
                "if($handler.Groups[1].Value -cne $handler.Groups[2].Value){throw 'Chance collectors differ'};" +
                "if($handler.Groups[1].Value -cne '.echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .lastevent; .exr -1; .echo VBAI_TEARDOWN_REGISTER_MODE LiveEventThread; r; kv; .echo VBAI_TEARDOWN_EXCEPTION_END; gn'){throw 'Collector must preserve live-event register and unhandled fatal semantics'};" +
                "$av=[regex]::Match($commands,'(?m)^sxd -c \"([^\"]+)\" -c2 \"([^\"]+)\" 0xc0000005$');" +
                "if(-not $av.Success -or $av.Groups[2].Value -cne $handler.Groups[1].Value){throw 'Second-chance AV collector differs'};" +
                "if($av.Groups[1].Value -cne '.if (@$t19 < 4) { .echo VBAI_FIRSTCHANCE_AV_BEGIN 424242 0123456789abcdef0123456789abcdef; .lastevent; .exr -1; .echo VBAI_FIRSTCHANCE_REGISTER_MODE LiveEventThread; r; kv; .echo VBAI_FIRSTCHANCE_AV_END; }; r @$t19=@$t19+1; gn'){throw 'First-chance AV must be forwarded unhandled'};" +
                "if($commands -notmatch '(?m)^sxn -c \"qd\" epr$'){throw 'Debugger may quit automatically only after process exit'};" +
                "if($commands -notmatch '(?m)^\\.echo VBAI_TEARDOWN_MODULES_BEGIN 424242 0123456789abcdef0123456789abcdef\\r?\\nlm\\r?\\n\\.echo VBAI_TEARDOWN_MODULES_END\\r?$'){throw 'Exact owned module inventory must precede teardown'};" +
                "if($commands -notmatch '(?m)^\\.echo VBAI_TEARDOWN_READY 424242 0123456789abcdef0123456789abcdef$'){throw 'Owned readiness identity missing'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }

        [DataTestMethod]
        [DataRow("Complete", 1)]
        [DataRow("MissingMode", 0)]
        [DataRow("MissingRegisters", 0)]
        [DataRow("MissingStack", 0)]
        [DataRow("WrongPid", 0)]
        [DataRow("WrongNonce", 0)]
        [DataRow("SecondChance", 0)]
        [DataRow("Echo", 0)]
        [DataRow("TwoComplete", 2)]
        [DataRow("Mixed", 1)]
        [DataRow("Empty", 0)]
        [DataRow("Null", 0)]
        [DataRow("InvalidTarget", 0)]
        [DataRow("InvalidNonce", 0)]
        public void FirstChanceDetailsRequireExecutedOwnedLiveContext(string variation, int expected)
        {
            const string nonce = "0123456789abcdef0123456789abcdef";
            string complete = "VBAI_FIRSTCHANCE_AV_BEGIN 424242 " + nonce + "\n" +
                "Last event: 67932.abc: Access violation - code c0000005 (first chance)\n" +
                "ExceptionCode: c0000005\nVBAI_FIRSTCHANCE_REGISTER_MODE LiveEventThread\n" +
                "rip=0000000000123456\nChild-SP RetAddr\n0000000000001111 0000000000002222 native!Frame\n" +
                "VBAI_FIRSTCHANCE_AV_END\n";
            string text = complete;
            switch (variation)
            {
                case "MissingMode": text = text.Replace("VBAI_FIRSTCHANCE_REGISTER_MODE LiveEventThread\n", ""); break;
                case "MissingRegisters": text = text.Replace("rip=0000000000123456", "rip=0000000000000000"); break;
                case "MissingStack": text = text.Replace("0000000000001111 0000000000002222 native!Frame\n", ""); break;
                case "WrongPid": text = text.Replace("67932.abc", "67933.abc"); break;
                case "WrongNonce": text = text.Replace(nonce, "ffffffffffffffffffffffffffffffff"); break;
                case "SecondChance": text = text.Replace("first chance", "second chance"); break;
                case "Echo": text = "0:000> .echo " + text; break;
                case "TwoComplete": text += complete; break;
                case "Mixed": text += complete.Replace("rip=0000000000123456", "rip=0000000000000000"); break;
                case "Empty": text = ""; break;
                case "Null": text = null; break;
            }
            string common = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                "; $actual=Get-OwnedFirstChanceDetailCount " + (text == null ? "$null" : Quote(text)) +
                " " + (variation == "InvalidTarget" ? "0" : "424242") + " " + Quote(variation == "InvalidNonce" ? "bad" : nonce) +
                "; if($actual -ne " + expected + "){throw ('Unexpected qualified detail count: '+$actual)}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
            {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
            };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000), "Pure parser timeout; no host or debugger is involved.");
                Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }

        [TestMethod]
        public void OnlyExecutedExactReadinessCanArmTheOwnedDebugger()
        {
            string common = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
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
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }

        [TestMethod]
        public void PendingFaultAndEchoedStopEventsCannotAuthorizeDetachment()
        {
            string common = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            string nonce = "0123456789abcdef0123456789abcdef";
            string stop = "VBAI_STOP_PROBE_BEGIN " + nonce + "\r\nLast event: 67932.1: breakpoint - code 80000003 (first chance)\r\nVBAI_STOP_PROBE_END\r\n";
            string firstAv = "VBAI_FIRSTCHANCE_AV_BEGIN 424242 " + nonce + "\r\nLast event: 67932.1: AV - code c0000005 (first chance)\r\nVBAI_FIRSTCHANCE_AV_END\r\n";
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";$nonce=" + Quote(nonce) + ";$stop=" + Quote(stop) + ";$av=" + Quote(firstAv) + ";" +
                "if(-not (Test-OwnedStopBreakpoint $stop 424242 $nonce)){throw 'Exact stop breakpoint refused'};" +
                "if(Test-OwnedStopBreakpoint ($stop.Replace('80000003','c0000005')) 424242 $nonce){throw 'Pending AV accepted for qd'};" +
                "if(Test-OwnedStopBreakpoint ($stop.Replace('80000003','c0000409')) 424242 $nonce){throw 'Pending failfast accepted for qd'};" +
                "if(Test-OwnedStopBreakpoint $stop 424243 $nonce){throw 'Different stop PID accepted'};" +
                "if(Test-OwnedStopBreakpoint $stop 424242 'fedcba9876543210fedcba9876543210'){throw 'Different stop nonce accepted'};" +
                "if(Test-OwnedStopBreakpoint ('0:000> .echo '+$stop) 424242 $nonce){throw 'Echoed stop accepted'};" +
                "if(-not (Test-OwnedFirstChanceAv $av 424242 $nonce)){throw 'Executed first-chance AV missing'};" +
                "if(Test-OwnedFirstChanceAv $av 424243 $nonce){throw 'Different forwarding PID accepted'};" +
                "if(Test-OwnedFirstChanceAv ($av.Replace('first chance','second chance')) 424242 $nonce){throw 'Second-chance AV called forwarding'};" +
                "if(Test-OwnedFirstChanceAv ('0:000> .echo '+$av) 424242 $nonce){throw 'Echoed forwarding accepted'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
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
            string common = Path.Combine(VBAi.Tests.Infrastructure.TestAssemblyPaths.OutputDirectory, "tools", "probes", "OwnedTeardownTrace.Common.ps1");
            string genuine = "VBAI_TEARDOWN_EXCEPTION_BEGIN\r\nLast event: 67932.1: fatal - code c0000409 (!!! second chance !!!)\r\nExceptionCode: c0000409\r\nVBAI_TEARDOWN_REGISTER_MODE LiveEventThread\r\nrip=0000000000000010 rsp=0000000000000020\r\nChild-SP RetAddr Call Site\r\n0000000000000020 0000000000000030 module!frame\r\nVBAI_TEARDOWN_EXCEPTION_END\r\n";
            string echo = "0:000> sxe -c \".echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .exr -1; kv; .echo VBAI_TEARDOWN_EXCEPTION_END\" 0xc0000409";
            string missingContext = genuine.Replace("rip=0000000000000010 rsp=0000000000000020", "Unable to get exception context, HRESULT 0x8000FFFF");
            string command = "$ErrorActionPreference='Stop'; . " + Quote(common) +
                ";if(-not (Test-TeardownExceptionCapture " + Quote(genuine) + " 424242)){throw 'Executed exception missing'};" +
                "if(Test-TeardownExceptionCapture " + Quote(echo) + " 424242){throw 'Command echo accepted'};" +
                "if(Test-TeardownExceptionCapture " + Quote(missingContext) + " 424242){throw 'Incomplete exception context accepted'};" +
                "$valid=" + Quote(genuine) + ";$incomplete=" + Quote(missingContext) + ";" +
                "if(-not (Test-OwnedTeardownFatalExit -1073740791 $true -1073740791 $valid 424242)){throw 'Exact terminal fatal propagation refused'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $false -1073740791 $valid 424242){throw 'Live target accepted'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true $null $valid 424242){throw 'Unknown target exit accepted'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true 0 $valid 424242){throw 'Different target exit accepted'};" +
                "if(Test-OwnedTeardownFatalExit 5 $true 5 $valid 424242){throw 'Generic matching nonzero accepted'};" +
                "if(Test-OwnedTeardownFatalExit 0 $true -1073740791 $valid 424242){throw 'Zero debugger exit called fatal propagation'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true -1073740791 $incomplete 424242){throw 'Missing registers accepted for fatal propagation'};" +
                "if(Test-TeardownExceptionCapture $valid 424243){throw 'Different event PID accepted'};" +
                "if(Test-TeardownExceptionCapture ($valid.Replace('0000000000000020 0000000000000030 module!frame','')) 424242){throw 'Stack header without frame accepted'};" +
                "if(Test-TeardownExceptionCapture ($valid.Replace('LiveEventThread','DumpContext')) 424242){throw 'Wrong register provenance accepted'};" +
                "$av=$valid.Replace('c0000409','c0000005');" +
                "if(-not (Test-OwnedTeardownFatalExit -1073741819 $true -1073741819 $av 424242)){throw 'Exact terminal AV propagation refused'};" +
                "if(Test-OwnedTeardownFatalExit -1073740791 $true -1073740791 $av 424242){throw 'Captured and exit codes differ'};" +
                "if(Test-TeardownExceptionCapture ($av.Replace('second chance','first chance')) 424242){throw 'First-chance AV promoted to fatal'}";
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            var info = new ProcessStartInfo(ps, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(command)))
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(ps), "Modules");
            using (var child = Process.Start(info))
            {
                var output = child.StandardOutput.ReadToEndAsync(); var error = child.StandardError.ReadToEndAsync();
                Assert.IsTrue(child.WaitForExit(10000)); Assert.AreEqual(0, child.ExitCode, output.Result + error.Result);
            }
        }
    }
}
