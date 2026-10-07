using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Checks Outlook's installed VBAi bridge without modifying its user VBA project or mailbox.</summary>
    [TestClass, TestCategory("Office")]
    public sealed class OutlookHostTests
    {
        /// <summary>Opens an unsaved inspector, activates VBE semantically, and reads host metadata only.</summary>
        [STATestMethod]
        public void OutlookReadOnlyVbeStartup()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OUTLOOK_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OUTLOOK_TESTS=1 for read-only Outlook qualification.");
            var existing = Process.GetProcessesByName("OUTLOOK");
            bool occupied = existing.Length != 0;
            foreach (var process in existing) process.Dispose();
            if (occupied) Assert.Inconclusive("Close Outlook before isolated read-only qualification.");
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) Assert.Inconclusive("Classic Outlook is not installed.");
            using (var profiles = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Profiles"))
                if (profiles == null || profiles.GetSubKeyNames().Length == 0)
                    Assert.Inconclusive("BLOCKED: classic Outlook 16 has no configured profile; qualification must not configure an account.");
            object application = null, item = null, inspector = null, commandBars = null;
            Process ownedProcess = null;
            bool scenarioPassed = false;
            int pid = 0;
            try
            {
                application = Activator.CreateInstance(type);
                var launched = Process.GetProcessesByName("OUTLOOK");
                try { Assert.AreEqual(1, launched.Length); pid = launched[0].Id; }
                finally { foreach (var process in launched) process.Dispose(); }
                ownedProcess = Process.GetProcessById(pid);
                _ = ownedProcess.Handle;
                item = ((dynamic)application).CreateItem(0);
                ((dynamic)item).Display(false);
                inspector = ((dynamic)item).GetInspector;
                commandBars = ((dynamic)inspector).CommandBars;
                ((dynamic)commandBars).ExecuteMso("VisualBasic");
                var status = VbeBridgeClient.Read(pid, "status");
                Assert.IsNotNull(status, "Outlook opened VBE but did not load the VBAi bridge.");
                Assert.AreEqual(true, status["Ok"], Convert.ToString(status["Error"]));
                var data = VbeBridgeClient.Object(status["Data"]);
                Assert.AreEqual(pid, Convert.ToInt32(data["HostProcessId"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), data["AssemblyModuleVersionId"]);
                Console.WriteLine("OUTLOOK BRIDGE: PID=" + pid + "; MVID=" + data["AssemblyModuleVersionId"]);
                var listed = VbeBridgeClient.Read(pid, "list_projects");
                Assert.IsNotNull(listed);
                Assert.AreEqual(true, listed["Ok"], Convert.ToString(listed["Error"]));
                var projects = ((object[])listed["Data"]).Select(VbeBridgeClient.Object).ToArray();
                Assert.AreEqual(1, projects.Length, "Read-only Outlook qualification requires one unambiguous project.");
                string selector = Convert.ToString(projects[0]["FileName"]);
                if (string.IsNullOrWhiteSpace(selector)) selector = Convert.ToString(projects[0]["Name"]);
                Assert.IsFalse(string.IsNullOrWhiteSpace(selector));
                foreach (var command in new[] { "debug_state", "vbe_environment" })
                {
                    var response = VbeBridgeClient.Read(pid, new { Command = command, Project = selector });
                    Assert.IsNotNull(response, command);
                    Assert.AreEqual(true, response["Ok"], command + ": " + Convert.ToString(response["Error"]));
                }
                scenarioPassed = true;
            }
            finally
            {
                if (pid != 0)
                {
                    if (inspector != null) try { ((dynamic)inspector).Close(1); } catch { } // olDiscard
                    if (application != null) try { ((dynamic)application).Quit(); } catch { }
                }
                foreach (var value in new[] { commandBars, inspector, item, application })
                    if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
                commandBars = inspector = item = application = null;
                GC.Collect(); GC.WaitForPendingFinalizers();
                GC.Collect(); GC.WaitForPendingFinalizers();
                if (ownedProcess != null) using (ownedProcess)
                {
                    bool exited = ownedProcess.WaitForExit(10000);
                    Console.WriteLine("OUTLOOK SHUTDOWN: PID=" + pid + "; exited=" + exited +
                        (exited ? "; exitCode=" + ownedProcess.ExitCode : "; left running for diagnosis"));
                    if (scenarioPassed)
                    {
                        Assert.IsTrue(exited, "Owned Outlook did not exit normally within 10 seconds; no process was killed.");
                        Assert.AreEqual(0, ownedProcess.ExitCode, "Owned Outlook exited abnormally.");
                    }
                }
            }
        }
    }
}
