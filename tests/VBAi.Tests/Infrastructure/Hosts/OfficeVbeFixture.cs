using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Threading;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Owns a visible Office instance and a disposable document; routes VBA operations through its installed add-in.</summary>
    internal sealed class OfficeVbeFixture : IDisposable
    {
        private object application, document;
        private bool owned;
        private volatile bool stopDialogs;
        private Thread dialogThread;
        private readonly List<object> steps = new List<object>();
        internal string Kind { get; private set; }
        internal int ProcessId { get; private set; }
        internal string Root { get; private set; }
        internal string DocumentPath { get; private set; }
        internal string Project { get; private set; }
        internal readonly List<string> Failures = new List<string>();
        private OfficeVbeFixture() { }

        /// <summary>Creates only a new process; existing host sessions are preserved.</summary>
        internal static OfficeVbeFixture Start(string kind)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_TESTS=1 to qualify installed Office hosts.");
            string executable = kind == "Word" ? "WINWORD" : kind == "PowerPoint" ? "POWERPNT" : kind == "Access" ? "MSACCESS" : "MSPUB";
            var processes = Process.GetProcessesByName(executable);
            int[] existing = processes.Select(p => p.Id).ToArray();
            foreach (var process in processes) process.Dispose();
            if (existing.Length != 0 && kind != "Word") Assert.Inconclusive("Close existing " + kind + " instances before isolated qualification.");
            var type = Type.GetTypeFromProgID(kind + ".Application");
            if (type == null) Assert.Inconclusive(kind + " is not installed.");
            var result = new OfficeVbeFixture { Kind = kind };
            string output = Environment.GetEnvironmentVariable("VBAi_OFFICE_RESULTS") ?? Path.Combine(Path.GetTempPath(), "VBAi-Office-tests");
            result.Root = Path.Combine(Path.GetFullPath(output), kind, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(result.Root);
            result.DocumentPath = Path.Combine(result.Root, "Disposable" + (kind == "Word" ? ".docm" : kind == "PowerPoint" ? ".pptm" : kind == "Access" ? ".accdb" : ".pub"));
            try
            {
                result.application = Activator.CreateInstance(type);
                var launched = Process.GetProcessesByName(executable);
                try
                {
                    var candidates = launched.Where(p => !existing.Contains(p.Id)).ToArray();
                    Assert.AreEqual(1, candidates.Length, "No unique new Office process; no document mutation is permitted.");
                    result.ProcessId = candidates[0].Id; result.owned = true;
                }
                finally { foreach (var process in launched) process.Dispose(); }
                if (kind == "Access" || kind == "Publisher") result.StartOwnedDialogHandler();
                dynamic app = result.application;
                if (kind == "Word")
                {
                    app.Visible = true; app.DisplayAlerts = 0; app.AutomationSecurity = 3;
                    result.document = app.Documents.Add();
                }
                else if (kind == "PowerPoint")
                {
                    app.Visible = -1; app.AutomationSecurity = 3;
                    result.document = app.Presentations.Add(-1);
                }
                else if (kind == "Access")
                {
                    app.Visible = true; app.NewCurrentDatabase(result.DocumentPath);
                }
                else
                {
                    result.document = app.NewDocument(); app.ActiveWindow.Visible = true;
                }
                result.SaveNative();
                if (kind == "Access") app.VBE.MainWindow.Visible = true;
                else app.CommandBars.ExecuteMso("VisualBasic");
                var status = result.Data("status");
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"], "Another add-in build is installed.");
                Assert.AreEqual(result.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                var projects = result.Items("list_projects");
                var matched = projects.Where(p => string.Equals(p["FileName"] as string, result.DocumentPath, StringComparison.OrdinalIgnoreCase)).ToArray();
                if (matched.Length == 1) result.Project = result.DocumentPath;
                else
                {
                    // A fresh Word document / Publisher publication has one non-template project.
                    var candidates = projects.Where(p => !string.Equals(p["Name"] as string, "Normal", StringComparison.OrdinalIgnoreCase)).ToArray();
                    Assert.AreEqual(1, candidates.Length, "No unique disposable document project; no mutation is permitted.");
                    result.Project = (string)candidates[0]["Name"];
                }
                result.Items("list_modules");
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        /// <summary>Reads a command response without retries after emission.</summary>
        internal IDictionary<string, object> Response(string name, params object[] pairs)
        {
            var request = new Dictionary<string, object> { ["Command"] = name };
            if (Project != null) request["Project"] = Project;
            for (int i = 0; i < pairs.Length; i += 2) request[(string)pairs[i]] = pairs[i + 1];
            var reply = VbeBridgeClient.Read(ProcessId, request);
            Assert.IsNotNull(reply, name + " did not answer.");
            steps.Add(new { Command = name, Response = reply });
            return reply;
        }
        /// <summary>Requires a successful command and returns its object result.</summary>
        internal IDictionary<string, object> Data(string name, params object[] pairs)
        {
            var reply = Response(name, pairs);
            Assert.AreEqual(true, reply["Ok"], name + ": " + (reply["Error"] ?? ""));
            return VbeBridgeClient.Object(reply["Data"]);
        }
        /// <summary>Requires a successful command and returns its array result.</summary>
        internal IDictionary<string, object>[] Items(string name, params object[] pairs)
        {
            var reply = Response(name, pairs);
            Assert.AreEqual(true, reply["Ok"], name + ": " + (reply["Error"] ?? ""));
            return ((object[])reply["Data"]).Select(VbeBridgeClient.Object).ToArray();
        }
        /// <summary>Records every scenario so a failed operation does not conceal independent missing functionality.</summary>
        internal void Scenario(string name, Action action)
        {
            try { action(); steps.Add(new { Scenario = name, Result = "PASS" }); }
            catch (Exception error) { Failures.Add(name + ": " + error.Message); steps.Add(new { Scenario = name, Result = "FAIL", Error = error.Message }); }
        }
        /// <summary>Saves only the owned disposable document, independently of VBAi's save adapter.</summary>
        internal void SaveNative()
        {
            if (document != null && File.Exists(DocumentPath)) ((dynamic)document).Save();
            else if (Kind == "Word") ((dynamic)document).SaveAs2(DocumentPath, 13);
            else if (Kind == "PowerPoint") ((dynamic)document).SaveAs(DocumentPath, 25);
            else if (Kind == "Publisher") ((dynamic)document).SaveAs(DocumentPath);
            // Access commits its module objects with its own close/reopen below.
        }
        /// <summary>Reopens only the owned file and never executes VBA automatically.</summary>
        internal void Reopen()
        {
            SaveNative(); dynamic app = application;
            if (Kind == "Word") { ((dynamic)document).Close(0); Release(document); document = app.Documents.Open(DocumentPath); }
            else if (Kind == "PowerPoint") { ((dynamic)document).Close(); Release(document); document = app.Presentations.Open(DocumentPath, 0, 0, -1); }
            else
            {
                // Publisher requires a new application for Open; Access can retain a VBIDE database reference.
                CloseOwnedHost();
                application = Activator.CreateInstance(Type.GetTypeFromProgID(Kind + ".Application"));
                var processes = Process.GetProcessesByName(Kind == "Access" ? "MSACCESS" : "MSPUB");
                try { Assert.AreEqual(1, processes.Length); ProcessId = processes[0].Id; owned = true; }
                finally { foreach (var process in processes) process.Dispose(); }
                app = application;
                StartOwnedDialogHandler();
                if (Kind == "Publisher") { document = app.Open(DocumentPath, false, false); app.ActiveWindow.Visible = true; app.CommandBars.ExecuteMso("VisualBasic"); }
                else { app.Visible = true; app.OpenCurrentDatabase(DocumentPath); app.VBE.MainWindow.Visible = true; }
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), Data("status")["AssemblyModuleVersionId"]);
            }
        }
        /// <summary>Handles only owned test dialogs: Access component saves and Publisher macro disabling.</summary>
        private void StartOwnedDialogHandler()
        {
            stopDialogs = false;
            dialogThread = new Thread(() => {
                while (!stopDialogs)
                {
                    try
                    {
                        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ProcessIdProperty, ProcessId));
                        foreach (AutomationElement window in windows)
                        {
                            if (Kind == "Publisher" && window.Current.Name == "Avis de sécurité pour Microsoft Publisher")
                            {
                                var disable = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Désactiver les macros"));
                                if (disable != null) ((InvokePattern)disable.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                                continue;
                            }
                            if (Kind != "Access") continue;
                            var edit = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "2020"));
                            if (edit == null) continue;
                            string name = ((ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value;
                            if (name != "VBAiOfficeModule" && name != "VBAiOfficeClass" && name != "VBAiOfficeForm") continue;
                            var button = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "1"));
                            if (button != null && button.Current.Name == "OK") ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                        }
                    }
                    catch (ElementNotAvailableException) { }
                    catch (InvalidOperationException) { }
                    Thread.Sleep(250);
                }
            }) { IsBackground = true };
            dialogThread.SetApartmentState(ApartmentState.STA); dialogThread.Start();
        }
        /// <summary>Closes only the owned process and preserves evidence files.</summary>
        public void Dispose()
        {
            CloseOwnedHost();
            if (Root != null) File.WriteAllText(Path.Combine(Root, "qualification.json"), new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 }.Serialize(new { Host = Kind, ProcessId, DocumentPath, Project, Failures, Steps = steps }));
        }
        private void CloseOwnedHost()
        {
            if (owned)
            {
                try
                {
                    if (Kind == "Word") { if (document != null) ((dynamic)document).Close(0); ((dynamic)application).Quit(0); }
                    else if (Kind == "PowerPoint") { if (document != null) { ((dynamic)document).Saved = -1; ((dynamic)document).Close(); } ((dynamic)application).Quit(); }
                    else if (Kind == "Access") { ((dynamic)application).CloseCurrentDatabase(); ((dynamic)application).Quit(2); }
                    else { if (document != null) ((dynamic)document).Close(); ((dynamic)application).Quit(); }
                }
                catch (Exception error) { steps.Add(new { CleanupError = error.Message }); }
            }
            Release(document); Release(application); document = application = null;
            if (owned)
                try { using (var process = Process.GetProcessById(ProcessId)) if (!process.WaitForExit(5000)) { process.Kill(); process.WaitForExit(5000); } }
                catch (ArgumentException) { }
            owned = false; stopDialogs = true;
            if (dialogThread != null) { dialogThread.Join(2000); dialogThread = null; }
        }
        private static void Release(object item) { if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item); }
    }
}
