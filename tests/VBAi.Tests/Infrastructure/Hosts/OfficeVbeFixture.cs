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
    internal sealed partial class OfficeVbeFixture : IDisposable
    {
        private object application, document;
        private bool owned;
        private Process ownedProcess;
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
                    result.CaptureOwnedProcess();
                }
                finally { foreach (var process in launched) process.Dispose(); }
                if (kind == "Access" || kind == "Publisher") result.StartOwnedDialogHandler();
                dynamic app = result.application;
                if (kind == "Word")
                {
                    app.Visible = true; app.DisplayAlerts = 0; app.AutomationSecurity = 3;
                    result.document = result.CreateOrOpenDocument(false);
                }
                else if (kind == "PowerPoint")
                {
                    app.Visible = -1; app.AutomationSecurity = 3;
                    result.document = result.CreateOrOpenDocument(false);
                }
                else if (kind == "Access")
                {
                    app.Visible = true; app.NewCurrentDatabase(result.DocumentPath);
                }
                else
                {
                    result.document = app.NewDocument(); result.ShowPublisherWindow();
                }
                result.SaveNative();
                result.ShowVbe();
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
                result.RecordNativeProjectPath();
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

        private void RecordNativeProjectPath()
        {
            if (Kind != "Word" && Kind != "PowerPoint") return;
            object project = null;
            try
            {
                project = ((dynamic)document).VBProject;
                try { steps.Add(new { NativeProjectFileName = (string)((dynamic)project).FileName }); }
                catch (Exception error)
                {
                    steps.Add(new { NativeProjectFileNameError = error.Message,
                        ExceptionType = error.GetType().FullName, HResult = "0x" + unchecked((uint)error.HResult).ToString("X8") });
                }
            }
            finally { Release(project); }
        }

        /// <summary>Records independent native save readback before any helper save can change the evidence.</summary>
        internal void RecordNativePersistence(string phase)
        {
            if (Kind != "Word" && Kind != "PowerPoint") return;
            var fields = new Dictionary<string, object> { ["NativePersistencePhase"] = phase };
            object project = null, editor = null, selected = null;
            try
            {
                dynamic nativeDocument = document;
                fields["DocumentPath"] = (string)nativeDocument.FullName;
                fields["DocumentSaved"] = Kind == "Word" ? (bool)nativeDocument.Saved : Convert.ToInt32(nativeDocument.Saved) == -1;
                fields["DocumentReadOnly"] = Kind == "Word" ? (bool)nativeDocument.ReadOnly : Convert.ToInt32(nativeDocument.ReadOnly) != 0;
                fields["DocumentFormat"] = Kind == "Word" ? (object)Convert.ToInt32(nativeDocument.SaveFormat) : null;
                fields["FileLength"] = File.Exists(DocumentPath) ? new FileInfo(DocumentPath).Length : 0;
                project = nativeDocument.VBProject;
                fields["ProjectSaved"] = (bool)((dynamic)project).Saved;
                try { fields["ProjectPath"] = (string)((dynamic)project).FileName; }
                catch (Exception error) { fields["ProjectPathError"] = error.GetType().FullName + " / 0x" + unchecked((uint)error.HResult).ToString("X8") + ": " + error.Message; }
                fields["SourceSha256"] = typeof(VbeProjectComponents).GetMethod("OtherHostSourceSha",
                    System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).Invoke(null, new[] { project });
                editor = ((dynamic)application).VBE;
                selected = VbeProjectResolver.Resolve((dynamic)editor, Project);
                IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
                try
                {
                    first = Marshal.GetIUnknownForObject(project); second = Marshal.GetIUnknownForObject(selected);
                    fields["SelectedProjectIdentityMatches"] = first == second;
                }
                finally { if (second != IntPtr.Zero) Marshal.Release(second); if (first != IntPtr.Zero) Marshal.Release(first); }
            }
            catch (Exception error) { fields["ReadbackError"] = error.ToString(); }
            finally
            {
                // Selected project and document.VBProject may share the same RCW; balance each acquisition.
                foreach (var value in new[] { selected, editor, project })
                    if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                steps.Add(fields);
            }
        }
        /// <summary>Records unavailable product capabilities separately from successful fail-closed policy checks.</summary>
        internal void CompatibilityGap(string capability, string reason)
        {
            steps.Add(new { Capability = capability, Result = "NOT_QUALIFIED", Reason = reason });
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
            SaveNative();
            ReopenCore();
        }
        /// <summary>Reopens a Word/PowerPoint file without a helper save, so only the adapter can have persisted edits.</summary>
        internal void ReopenFromDisk()
        {
            if (Kind != "Word" && Kind != "PowerPoint")
                throw new InvalidOperationException("Adapter-only reopen is qualified only for Word and PowerPoint.");
            ReopenCore();
        }
        private void ReopenCore()
        {
            dynamic app = application;
            if (Kind == "Word") { ((dynamic)document).Close(0); Release(document); document = null; document = CreateOrOpenDocument(true); }
            else if (Kind == "PowerPoint") { ((dynamic)document).Saved = -1; ((dynamic)document).Close(); Release(document); document = null; document = CreateOrOpenDocument(true); }
            else
            {
                // Publisher requires a new application for Open; Access can retain a VBIDE database reference.
                int failuresBeforeClose = Failures.Count;
                CloseOwnedHost();
                Assert.AreEqual(failuresBeforeClose, Failures.Count, "Cannot reopen after an unsuccessful host shutdown: " + string.Join(Environment.NewLine, Failures));
                application = Activator.CreateInstance(Type.GetTypeFromProgID(Kind + ".Application"));
                var processes = Process.GetProcessesByName(Kind == "Access" ? "MSACCESS" : "MSPUB");
                try { Assert.AreEqual(1, processes.Length); ProcessId = processes[0].Id; owned = true; CaptureOwnedProcess(); }
                finally { foreach (var process in processes) process.Dispose(); }
                app = application;
                StartOwnedDialogHandler();
                if (Kind == "Publisher") { document = app.Open(DocumentPath, false, false); ShowPublisherWindow(); }
                else { app.Visible = true; app.OpenCurrentDatabase(DocumentPath); }
                ShowVbe();
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
            try { CloseOwnedHost(); }
            finally
            {
                if (Root != null) File.WriteAllText(Path.Combine(Root, "qualification.json"), new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 }.Serialize(new { Host = Kind, ProcessId, DocumentPath, Project, Failures, Steps = steps }));
            }
            Assert.AreEqual(0, Failures.Count, string.Join(Environment.NewLine, Failures));
        }
        /// <summary>Releases the collection RCW explicitly instead of leaving a chained COM temporary alive.</summary>
        private object CreateOrOpenDocument(bool reopen)
        {
            object collection = null;
            try
            {
                collection = Kind == "Word" ? (object)((dynamic)application).Documents : (object)((dynamic)application).Presentations;
                if (Kind == "Word") return reopen ? ((dynamic)collection).Open(DocumentPath) : ((dynamic)collection).Add();
                return reopen ? ((dynamic)collection).Open(DocumentPath, 0, 0, -1) : ((dynamic)collection).Add(-1);
            }
            finally { Release(collection); }
        }
        private void ShowPublisherWindow()
        {
            object window = null;
            try { window = ((dynamic)application).ActiveWindow; ((dynamic)window).Visible = true; }
            finally { Release(window); }
        }
        private void ShowVbe()
        {
            object bars = null, editor = null, window = null;
            try
            {
                if (Kind == "Access")
                {
                    editor = ((dynamic)application).VBE;
                    window = ((dynamic)editor).MainWindow;
                    ((dynamic)window).Visible = true;
                }
                else
                {
                    bars = ((dynamic)application).CommandBars;
                    ((dynamic)bars).ExecuteMso("VisualBasic");
                }
            }
            finally { Release(window); Release(editor); Release(bars); }
        }
        private void CaptureOwnedProcess()
        {
            ownedProcess = Process.GetProcessById(ProcessId);
            _ = ownedProcess.Handle; // Preserve teardown/crash evidence after the process exits.
        }
        private void RecordCleanupFailure(string reason)
        {
            Failures.Add("Host cleanup: " + reason);
            steps.Add(new { CleanupError = reason, Result = "FAIL" });
        }
        private void CloseOwnedHost()
        {
            if (owned)
            {
                try
                {
                    if (Kind == "Word") { if (document != null) ((dynamic)document).Close(0); }
                    else if (Kind == "PowerPoint") { if (document != null) { ((dynamic)document).Saved = -1; ((dynamic)document).Close(); } }
                    else if (Kind == "Access") { ((dynamic)application).CloseCurrentDatabase(); ((dynamic)application).Quit(2); }
                    else { if (document != null) ((dynamic)document).Close(); }
                }
                catch (Exception error) { RecordCleanupFailure(error.Message); }
            }
            try { Release(document); } catch (Exception error) { RecordCleanupFailure(error.Message); }
            document = null;
            if (owned && Kind != "Access")
                try { if (Kind == "Word") ((dynamic)application).Quit(0); else ((dynamic)application).Quit(); }
                catch (Exception error) { RecordCleanupFailure(error.Message); }
            try { Release(application); } catch (Exception error) { RecordCleanupFailure(error.Message); }
            application = null;
            var process = ownedProcess;
            ownedProcess = null;
            if (process != null)
                using (process)
                    try
                    {
                        bool forced = !process.WaitForExit(5000);
                        if (forced)
                        {
                            RecordCleanupFailure("The owned host did not exit after Quit and COM release; forced termination was required. PID=" + ProcessId);
                            process.Kill();
                        }
                        if (!process.WaitForExit(5000)) RecordCleanupFailure("The owned host still has not exited. PID=" + ProcessId);
                        else
                        {
                            steps.Add(new { ShutdownProcessId = ProcessId, ExitCode = process.ExitCode, ForcedTermination = forced });
                            if (process.ExitCode != 0) RecordCleanupFailure("Abnormal exit code 0x" + unchecked((uint)process.ExitCode).ToString("X8") + ". PID=" + ProcessId);
                        }
                    }
                    catch (Exception error) { RecordCleanupFailure(error.Message); }
            else if (owned) RecordCleanupFailure("No retained process handle was available to verify host shutdown. PID=" + ProcessId);
            owned = false; stopDialogs = true;
            if (dialogThread != null) { dialogThread.Join(2000); dialogThread = null; }
        }
        private static void Release(object item) { if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item); }
    }
}
