using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests
{
    /// <summary>Runs the real prepare-only controller in Windows PowerShell 5 with BOM-less UTF-8 identity paths.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class TraceNativeUserFormExportTests
    {
        [TestMethod]
        public void PrepareOnlyPreservesNonAsciiUtf8PendingIdentityWithoutHostOrDebugger()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi_trace_utf8_" + Guid.NewGuid().ToString("N"), "D\u00e9veloppement_\u03a9");
            string child = Path.Combine(root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(child);
            bool passed = false;
            try
            {
                string assembly = Path.Combine(root, "VBAi-\u00e9.dll");
                File.WriteAllText(assembly, "Owned identity fixture, never loaded or executed.", new UTF8Encoding(false));
                string hash;
                using (var sha = SHA256.Create())
                    hash = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(assembly))).Replace("-", "");
                Guid mvid = Guid.NewGuid();
                string destination = Path.Combine(child, "Formulaire\u00e9.frm");
                string pendingPath = Path.Combine(root, "pending-\u00e9.json");
                var serializer = new JavaScriptSerializer();
                var pending = new Dictionary<string, object> {
                    { "ProcessId", 0 }, { "ProcessStartedUtc", DateTime.UtcNow.ToString("o") },
                    { "Destination", destination }, { "AssemblyMvid", mvid.ToString() }, { "AssemblyPath", assembly }
                };
                File.WriteAllText(pendingPath, serializer.Serialize(pending), new UTF8Encoding(false));
                byte[] raw = File.ReadAllBytes(pendingPath);
                Assert.AreEqual((byte)'{', raw[0], "Fixture must have no BOM, matching native pending reports.");
                string script = FindController();
                string command = "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; " +
                    "if($PSVersionTable.PSVersion.Major -ne 5){throw 'Windows PowerShell 5 required'}; " +
                    "[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false); " +
                    "& " + Quote(script) + " -PendingReport " + Quote(pendingPath) +
                    " -ExpectedMvid " + Quote(mvid.ToString()) + " -ExpectedAssemblySha256 " + Quote(hash) +
                    " -CdbPath " + Quote(assembly);
                string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
                var info = new ProcessStartInfo(powershell, "-NoLogo -NoProfile -NonInteractive -EncodedCommand " +
                    Convert.ToBase64String(Encoding.Unicode.GetBytes(command))) {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
                };
                // A PS7-launched testhost can inherit only PS7 module roots.
                // Scope this override to the new owned PS5 child; never alter
                // the user's environment, profile, trust or execution policy.
                info.EnvironmentVariables["PSModulePath"] = Path.Combine(Path.GetDirectoryName(powershell), "Modules");
                using (var process = Process.Start(info))
                {
                    var stdout = process.StandardOutput.ReadToEndAsync();
                    var stderr = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(10000)) { process.Kill(); Assert.Fail("Owned prepare-only PowerShell timed out. Fixture retained: " + root); }
                    Assert.AreEqual(0, process.ExitCode, "Fixture retained: " + root + "\n" + stdout.Result + stderr.Result);
                    var plan = serializer.Deserialize<Dictionary<string, object>>(stdout.Result);
                    Assert.AreEqual("PREPARE_ONLY", plan["Mode"]);
                    Assert.AreEqual(destination, plan["Destination"], "Unicode native identity must survive file decoding and plan output.");
                    Assert.AreEqual(hash, plan["ExpectedAssemblySha256"]);
                    Assert.AreEqual(mvid.ToString(), plan["ExpectedMvid"]);
                    Assert.AreEqual(0, Convert.ToInt32(plan["ProcessId"]), "PID 0 cannot be attached; prepare-only must return before process lookup.");
                }
                Assert.AreEqual(0, Directory.GetFiles(child).Length, "No export, command file or permission marker may be written in prepare-only.");
                Assert.AreEqual(2, Directory.GetFiles(root).Length, "Only the synthetic assembly and pending report should exist.");
                passed = true;
            }
            finally { if (passed) Directory.Delete(Path.GetDirectoryName(root), true); }
        }

        private static string Quote(string value) { return "'" + value.Replace("'", "''") + "'"; }

        private static string FindController()
        {
            foreach (string start in new[] { AppDomain.CurrentDomain.BaseDirectory, Environment.CurrentDirectory })
                for (var directory = new DirectoryInfo(start); directory != null; directory = directory.Parent)
                {
                    string candidate = Path.Combine(directory.FullName, "tools", "probes", "Trace-NativeUserFormExport.ps1");
                    if (File.Exists(candidate)) return candidate;
                }
            throw new FileNotFoundException("Matching prepare-only controller source is unavailable.");
        }
    }
}
