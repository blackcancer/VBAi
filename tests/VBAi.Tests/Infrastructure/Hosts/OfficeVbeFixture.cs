using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
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
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        private delegate bool DialogWindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumWindows(DialogWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int count);
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr dialog, int id);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr control);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)]
        private static extern IntPtr ReadDialogText(IntPtr window, uint message, UIntPtr capacity, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")]
        private static extern IntPtr SendDialogCommand(IntPtr window, uint message, UIntPtr command, IntPtr control, uint flags, uint timeout, out UIntPtr result);
        private OwnedDialogWorker dialogWorker;
        /// <summary>Captures one host generation; stopping it cannot be undone by a later worker.</summary>
        private sealed class OwnedDialogWorker
        {
            internal readonly int ProcessId;
            internal readonly Process HostProcess;
            internal readonly string Kind;
            internal volatile bool StopRequested;
            internal Thread Thread;
            internal Exception Failure;
            internal readonly HashSet<IntPtr> InvokedDialogs = new HashSet<IntPtr>();
            internal OwnedDialogWorker(Process hostProcess, string kind) { HostProcess = hostProcess; ProcessId = hostProcess.Id; Kind = kind; }
            internal bool HostAlive { get { try { return !HostProcess.HasExited; } catch (InvalidOperationException) { return false; } } }
        }
        private bool hostTeardownRefused;
        // Retain refused native instances explicitly; final RCW release could otherwise close them implicitly.
        private static readonly List<OfficeVbeFixture> retainedOfficeFixtures = new List<OfficeVbeFixture>();
        private readonly List<object> steps = new List<object>();
        private readonly OfficeCommandContainment commandContainment = new OfficeCommandContainment();
        internal Func<int, object, IDictionary<string, object>> Dispatch = (pid, request) => VbeBridgeClient.Read(pid, request);
        internal string Kind { get; private set; }
        private string hostProgId;
        internal int ProcessId { get; private set; }
        internal string Root { get; private set; }
        internal string DocumentPath { get; private set; }
        internal string Project { get; private set; }
        internal readonly List<string> Failures = new List<string>();
        private OfficeVbeFixture() { }

        /// <summary>Creates only a new process; existing host sessions are preserved.</summary>
        internal static OfficeVbeFixture Start(string kind)
        {
            return Start(kind, false);
        }

        /// <summary>Preserves an uncertain startup for a read-only identity investigation.</summary>
        internal static OfficeVbeFixture StartAccessIdentityProbe() { return Start("Access", true); }

        private static OfficeVbeFixture Start(string kind, bool preserveStartupFailure)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_TESTS=1 to qualify installed Office hosts.");
            string executable = kind == "Word" ? "WINWORD" : kind == "PowerPoint" ? "POWERPNT" : kind == "Access" ? "MSACCESS" : "MSPUB";
            var processes = Process.GetProcessesByName(executable);
            int[] existing = processes.Select(p => p.Id).ToArray();
            foreach (var process in processes) process.Dispose();
            if (existing.Length != 0 && kind != "Word") Assert.Inconclusive("Close existing " + kind + " instances before isolated qualification.");
            string progId = ResolveHostProgId(kind);
            var type = Type.GetTypeFromProgID(progId);
            if (type == null) Assert.Inconclusive(kind + " is not installed.");
            var result = new OfficeVbeFixture { Kind = kind, hostProgId = progId };
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
                    result.RecordPublisherActivationCanary();
                    result.document = app.NewDocument();
                    result.RequirePublisherPublication("AfterNewDocumentBeforeBootstrapSave", false);
                    result.ShowPublisherWindow();
                }
                result.SaveNative();
                if (kind == "Publisher") result.RequirePublisherPublication("AfterBootstrapSave", true);
                result.ShowVbe();
                var status = result.Data("status");
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"], "Another add-in build is installed.");
                Assert.AreEqual(result.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                var projects = result.Items("list_projects");
                result.BindStartupProject(projects);
                result.Items("list_modules");
                result.RecordNativeProjectPath();
                return result;
            }
            catch (Exception startupError)
            {
                result.Failures.Add("Host startup: " + startupError);
                result.steps.Add(new { StartupError = startupError.ToString(), HostProgId = progId, Result = "FAIL" });
                if (preserveStartupFailure || kind == "Publisher")
                {
                    result.RetainUncertainOffice();
                    try { result.FlushAdapterEvidence(); }
                    catch (Exception evidenceError) { throw new AggregateException("Read-only startup and evidence persistence both failed; host retained.", startupError, evidenceError); }
                    throw;
                }
                try { result.Dispose(false); }
                catch (Exception cleanupError) { throw new AggregateException("Office startup and cleanup both failed.", startupError, cleanupError); }
                throw;
            }
        }

        /// <summary>Binds the project used by the disposable baseline; kept separate for fake-native startup regressions.</summary>
        internal void BindStartupProject(IDictionary<string, object>[] projects)
        {
            if (Kind == "Publisher") { BindPublisherStartupProject(projects); return; }
            var matched = projects.Where(p => string.Equals(p["FileName"] as string, DocumentPath, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matched.Length == 1) Project = DocumentPath;
            else
            {
                var candidates = projects.Where(p => !string.Equals(p["Name"] as string, "Normal", StringComparison.OrdinalIgnoreCase)).ToArray();
                Assert.AreEqual(1, candidates.Length, "No unique disposable document project; no mutation is permitted.");
                Project = (string)candidates[0]["Name"];
            }
        }

        /// <summary>Accepts only the host's generic or numeric versioned ProgID, without changing registration.</summary>
        private static string ResolveHostProgId(string kind)
        {
            Assert.IsTrue(new[] { "Word", "PowerPoint", "Access", "Publisher" }.Contains(kind), "Unknown Office host.");
            string generic = kind + ".Application";
            string configured = Environment.GetEnvironmentVariable("VBAi_OFFICE_" + kind.ToUpperInvariant() + "_PROGID");
            if (configured == null) return generic;
            string prefix = generic + ".";
            bool versioned = configured.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && configured.Length > prefix.Length &&
                configured.Substring(prefix.Length).All(c => c >= '0' && c <= '9');
            Assert.IsTrue(string.Equals(configured, generic, StringComparison.OrdinalIgnoreCase) || versioned,
                "The Office ProgID override must be " + generic + " or " + generic + ".<digits>.");
            return configured;
        }

        /// <summary>Reads a command response without retries after emission.</summary>
        internal IDictionary<string, object> Response(string name, params object[] pairs)
        {
            var request = new Dictionary<string, object> { ["Command"] = name };
            if (Project != null) request["Project"] = Project;
            for (int i = 0; i < pairs.Length; i += 2) request[(string)pairs[i]] = pairs[i + 1];
            return commandContainment.Send(name, request, record => steps.Add(record), FlushAdapterEvidence,
                () => Dispatch(ProcessId, request), RetainUncertainOffice);
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
            commandContainment.RequireTerminal();
            if (Kind == "Access" || Kind == "Publisher")
            {
                object current = null, database = null;
                var observations = new Dictionary<string, object> { ["NativePersistencePhase"] = phase };
                try
                {
                    RequireOwnedDocument();
                    current = Kind == "Access" ? (object)((dynamic)application).CurrentProject : document;
                    observations["DocumentPath"] = (string)((dynamic)current).FullName;
                    observations["DocumentSaved"] = Kind == "Access" ? null : (object)(bool)((dynamic)current).Saved;
                    if (Kind == "Access")
                    {
                        database = ((dynamic)application).CurrentDb();
                        observations["DocumentReadOnly"] = !(bool)((dynamic)database).Updatable;
                        observations["DocumentFormat"] = Convert.ToInt32(((dynamic)current).FileFormat);
                    }
                    else
                    {
                        observations["DocumentReadOnly"] = (bool)((dynamic)current).ReadOnly;
                        observations["DocumentFormat"] = Convert.ToInt32(((dynamic)current).SaveFormat);
                    }
                    observations["FileLength"] = File.Exists(DocumentPath) ? new FileInfo(DocumentPath).Length : 0;
                    observations["BoundProjectPersistence"] = Data("project_persistence_status");
                    observations["ProjectProperties"] = Data("project_properties");
                }
                catch (Exception error) { observations["ReadbackError"] = error.ToString(); }
                finally { Release(database); if (Kind == "Access") Release(current); steps.Add(observations); }
                return;
            }
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
            commandContainment.RequireTerminal();
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
        /// <summary>Reopens the owned Office file without a helper save; discard-on-close cannot mask an adapter failure.</summary>
        internal void ReopenFromDisk(Action afterOwnedClose = null)
        {
            Assert.IsTrue(File.Exists(DocumentPath), "The adapter must have left an existing native file.");
            Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), Data("status")["AssemblyModuleVersionId"]);
            ReopenCore(true, afterOwnedClose);
        }
        private void ReopenCore(bool adapterOnly = false, Action afterOwnedClose = null)
        {
            commandContainment.RequireTerminal();
            dynamic app = application;
            if (Kind == "Word") { ((dynamic)document).Close(0); Release(document); document = null; document = CreateOrOpenDocument(true); }
            else if (Kind == "PowerPoint") { ((dynamic)document).Saved = -1; ((dynamic)document).Close(); Release(document); document = null; document = CreateOrOpenDocument(true); }
            else
            {
                // Full process exit removes stale Access VBIDE storage; Publisher Open requires a new instance.
                Assert.IsFalse(hostTeardownRefused, "Cannot reopen an Office fixture whose previous shutdown was refused.");
                StopOwnedDialogHandler();
                int failuresBeforeClose = Failures.Count;
                CloseOwnedHost(adapterOnly);
                Assert.AreEqual(failuresBeforeClose, Failures.Count, "Cannot reopen after an unsuccessful host shutdown: " + string.Join(Environment.NewLine, Failures));
                // Optional read-only disk evidence is captured after normal exit, before a fresh host can write.
                afterOwnedClose?.Invoke();
                var existing = Process.GetProcessesByName(Kind == "Access" ? "MSACCESS" : "MSPUB");
                try { Assert.AreEqual(0, existing.Length, "An unrelated Office instance appeared; no reopen is permitted."); }
                finally { foreach (var process in existing) process.Dispose(); }
                application = Activator.CreateInstance(Type.GetTypeFromProgID(hostProgId));
                var processes = Process.GetProcessesByName(Kind == "Access" ? "MSACCESS" : "MSPUB");
                try { Assert.AreEqual(1, processes.Length); ProcessId = processes[0].Id; owned = true; CaptureOwnedProcess(); }
                finally { foreach (var process in processes) process.Dispose(); }
                app = application;
                StartOwnedDialogHandler();
                if (Kind == "Publisher")
                {
                    Project = null;
                    document = app.Open(DocumentPath, false, false);
                    RequirePublisherPublication("AfterFreshDiskOpen", true);
                    ShowPublisherWindow();
                }
                else { app.Visible = true; app.OpenCurrentDatabase(DocumentPath); }
                ShowVbe();
                var status = Data("status");
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(ProcessId, Convert.ToInt32(status["HostProcessId"]));
                if (Kind == "Publisher") BindStartupProject(Items("list_projects"));
                RequireOwnedDocument();
                steps.Add(new { ReopenFromDisk = adapterOnly, HelperSaveInvoked = !adapterOnly, ProcessId, DocumentPath });
            }
        }
        /// <summary>Handles only owned test dialogs: Access component saves and Publisher macro disabling.</summary>
        private void StartOwnedDialogHandler()
        {
            StopOwnedDialogHandler();
            var worker = new OwnedDialogWorker(ownedProcess, Kind);
            dialogWorker = worker;
            worker.Thread = new Thread(() => {
                while (!worker.StopRequested && worker.HostAlive)
                {
                    try
                    {
                        if (worker.Kind == "Access")
                        {
                            HandleAccessSaveDialogs(worker);
                            if (!worker.StopRequested) Thread.Sleep(250);
                            continue;
                        }
                        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ProcessIdProperty, worker.ProcessId));
                        foreach (AutomationElement window in windows)
                        {
                            if (worker.StopRequested) break;
                            if (worker.Kind == "Publisher" && window.Current.Name == "Avis de sécurité pour Microsoft Publisher")
                            {
                                var disable = window.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Désactiver les macros"));
                                if (disable != null && disable.Current.Name == "Désactiver les macros") InvokeOwnedDialogButton(worker, window, disable);
                                continue;
                            }
                        }
                    }
                    catch (ElementNotAvailableException) { }
                    catch (InvalidOperationException) { }
                    if (!worker.StopRequested) Thread.Sleep(250);
                }
            }) { IsBackground = true };
            worker.Thread.SetApartmentState(ApartmentState.STA); worker.Thread.Start();
        }

        /// <summary>Reads only owned native Access save dialogs with bounded cross-process messages.</summary>
        private static void HandleAccessSaveDialogs(OwnedDialogWorker worker)
        {
            var scan = Stopwatch.StartNew();
            EnumWindows((window, parameter) => {
                if (worker.StopRequested || scan.ElapsedMilliseconds >= 1000) return false;
                if (!OwnedDialogControl(worker, window, "#32770", null) || !IsWindowVisible(window) || worker.InvokedDialogs.Contains(window)) return true;
                IntPtr edit = GetDlgItem(window, 2020), button = GetDlgItem(window, 1);
                if (!OwnedDialogControl(worker, edit, "Edit", 2020) || !OwnedDialogControl(worker, button, "Button", 1) || !IsWindowEnabled(button)) return true;
                string name = BoundedDialogText(edit);
                if (name != "VBAiOfficeModule" && name != "VBAiOfficeClass" && name != "VBAiOfficeForm") return true;
                if (BoundedDialogText(button) != "OK") return true;
                if (worker.StopRequested || GetDlgItem(window, 2020) != edit || GetDlgItem(window, 1) != button ||
                    !OwnedDialogControl(worker, window, "#32770", null) || !OwnedDialogControl(worker, edit, "Edit", 2020) ||
                    !OwnedDialogControl(worker, button, "Button", 1) || !IsWindowEnabled(button)) return true;
                worker.InvokedDialogs.Add(window); // Never repeat an uncertain save-dialog mutation.
                UIntPtr result;
                if (SendDialogCommand(window, 0x0111, new UIntPtr(1), button, 0x23, 750, out result) == IntPtr.Zero)
                {
                    worker.Failure = new InvalidOperationException("The owned Access save-dialog command was uncertain; it will not be retried.");
                    worker.StopRequested = true;
                }
                return !worker.StopRequested;
            }, IntPtr.Zero);
        }

        private static bool OwnedDialogControl(OwnedDialogWorker worker, IntPtr window, string expectedClass, int? id)
        {
            if (window == IntPtr.Zero || !worker.HostAlive) return false;
            uint pid; GetWindowThreadProcessId(window, out pid);
            var name = new StringBuilder(64); GetClassName(window, name, name.Capacity);
            return pid == (uint)worker.ProcessId && name.ToString() == expectedClass && (!id.HasValue || GetDlgCtrlID(window) == id.Value);
        }

        private static string BoundedDialogText(IntPtr window)
        {
            var text = new StringBuilder(128); UIntPtr result;
            return ReadDialogText(window, 0x000D, new UIntPtr((uint)text.Capacity), text, 0x23, 150, out result) == IntPtr.Zero ? null : text.ToString();
        }

        /// <summary>Rechecks cancellation after provider calls and immediately before the owned button invocation.</summary>
        private static void InvokeOwnedDialogButton(OwnedDialogWorker worker, AutomationElement window, AutomationElement button)
        {
            if (worker.StopRequested || !worker.HostAlive) return;
            var pattern = (InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern);
            if (window.Current.ProcessId != worker.ProcessId || button.Current.ProcessId != worker.ProcessId) return;
            if (worker.StopRequested || !worker.HostAlive) return;
            pattern.Invoke();
        }

        /// <summary>Preserves a blocked worker reference and refuses reuse until its native provider call has returned.</summary>
        private void StopOwnedDialogHandler()
        {
            var worker = dialogWorker;
            if (worker == null) return;
            worker.StopRequested = true;
            Assert.IsTrue(worker.Thread.Join(2000),
                "The owned " + worker.Kind + " dialog worker did not stop; no new worker or reopen is permitted. PID=" + worker.ProcessId);
            Assert.IsNull(worker.Failure, worker.Failure?.Message);
            Assert.IsFalse(worker.InvokedDialogs.Any(window => OwnedDialogControl(worker, window, "#32770", null) && IsWindowVisible(window)),
                "An owned Access save dialog remains after handler shutdown; no fresh mutation or reopen is permitted.");
            dialogWorker = null;
        }

        /// <summary>Stops Access save-dialog approval before the final unsaved code mutation.</summary>
        internal void StopAccessSaveDialogHandler()
        {
            if (Kind != "Access") return;
            StopOwnedDialogHandler();
        }
        /// <summary>Closes only the owned process and preserves evidence files.</summary>
        public void Dispose() { Dispose(true); }

        /// <summary>During startup failure, compares only newly introduced cleanup failures.</summary>
        private void Dispose(bool includeExistingFailures)
        {
            int failuresBeforeCleanup = Failures.Count;
            Exception cleanupFailure = null;
            try
            {
                CloseOwnedHost(adapterOnlyCleanup);
                Assert.AreEqual(includeExistingFailures ? 0 : failuresBeforeCleanup, Failures.Count,
                    string.Join(Environment.NewLine, includeExistingFailures ? Failures : Failures.Skip(failuresBeforeCleanup)));
            }
            catch (Exception error)
            {
                cleanupFailure = commandContainment.Failure != null && commandContainment.Uncertain
                    ? new AggregateException("The original Office dispatch and its cleanup refusal are both retained.", commandContainment.Failure, error)
                    : error;
            }
            try
            {
                if (Root != null) File.WriteAllText(Path.Combine(Root, "qualification.json"), new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 }.Serialize(new { Host = Kind, HostProgId = hostProgId, ProcessId, DocumentPath, Project,
                    PendingCommand = commandContainment.Command, CommandPending = commandContainment.Pending,
                    DeliveryUncertain = commandContainment.Uncertain, Failures, Steps = steps }));
            }
            catch (Exception evidence)
            {
                if (cleanupFailure != null) throw new AggregateException("Office cleanup/refusal and final evidence persistence both failed.", cleanupFailure, evidence);
                throw;
            }
            if (cleanupFailure != null) ExceptionDispatchInfo.Capture(cleanupFailure).Throw();
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
        /// <summary>Retains native ownership without invoking COM or waiting on a stalled host.</summary>
        private void RetainUncertainOffice()
        {
            if (!hostTeardownRefused)
            {
                hostTeardownRefused = true;
                lock (retainedOfficeFixtures) retainedOfficeFixtures.Add(this);
            }
            if (dialogWorker != null) dialogWorker.StopRequested = true;
        }
        private void CloseOwnedHost(bool adapterOnly = false)
        {
            if (commandContainment.Pending || commandContainment.Uncertain)
            {
                RetainUncertainOffice();
                RecordCleanupFailure("Bridge delivery remains pending/uncertain; process, COM references and evidence retained. No Close/Quit/reset/release/reopen was emitted. PID=" + ProcessId);
                return;
            }
            if (hostTeardownRefused) return;
            bool nativeIdentityVerified = Kind != "Access" && Kind != "Publisher";
            if (owned)
            {
                try
                {
                    if (Kind == "Access" || Kind == "Publisher")
                    {
                        RequireOwnedDocument(); nativeIdentityVerified = true;
                        StopOwnedDialogHandler();
                    }
                    if (Kind == "Word") { if (document != null) ((dynamic)document).Close(0); }
                    else if (Kind == "PowerPoint") { if (document != null) { ((dynamic)document).Saved = -1; ((dynamic)document).Close(); } }
                    else if (Kind == "Access")
                    {
                        if (adapterOnly)
                        {
                            StopAccessSaveDialogHandler();
                            CloseAccessObjectsWithoutSaving();
                        }
                        ((dynamic)application).CloseCurrentDatabase(); ((dynamic)application).Quit(2);
                    }
                    else if (Kind == "Publisher")
                    {
                        Assert.IsNotNull(document, "The owned Publisher document must be identified before Quit.");
                        RequireSoleOwnedSavedPublisherBeforeQuit();
                        ((dynamic)application).Quit();
                    }
                }
                catch (Exception error)
                {
                    if (Kind == "Publisher" || Kind == "Access")
                    {
                        hostTeardownRefused = true;
                        lock (retainedOfficeFixtures) retainedOfficeFixtures.Add(this);
                        RecordCleanupFailure(Kind + " close/Quit refused or failed; instance and COM references retained without retry: " + error.Message);
                        if (dialogWorker != null) dialogWorker.StopRequested = true;
                        return;
                    }
                    RecordCleanupFailure(error.Message);
                }
            }
            try { Release(document); } catch (Exception error) { RecordCleanupFailure(error.Message); }
            document = null;
            if (owned && nativeIdentityVerified && Kind != "Access" && Kind != "Publisher")
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
                        bool exited = process.WaitForExit(5000);
                        if (!exited)
                        {
                            RecordCleanupFailure("The owned host did not exit after Quit and COM release; it was retained without forced termination. PID=" + ProcessId);
                        }
                        else
                        {
                            steps.Add(new { ShutdownProcessId = ProcessId, ExitCode = process.ExitCode, ForcedTermination = false, AdapterOnlyClose = adapterOnly });
                            if (process.ExitCode != 0) RecordCleanupFailure("Abnormal exit code 0x" + unchecked((uint)process.ExitCode).ToString("X8") + ". PID=" + ProcessId);
                        }
                    }
                    catch (Exception error) { RecordCleanupFailure(error.Message); }
            else if (owned) RecordCleanupFailure("No retained process handle was available to verify host shutdown. PID=" + ProcessId);
            owned = false;
            StopOwnedDialogHandler();
        }

        /// <summary>Refuses Quit unless the only publication is the saved disposable document in the retained process.</summary>
        private void RequireSoleOwnedSavedPublisherBeforeQuit()
        {
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited, "The retained Publisher process must be alive.");
            if (Project != null)
            {
                var request = new { Command = "project_persistence_status", Project };
                var reply = commandContainment.Send("project_persistence_status", request, record => steps.Add(record), FlushAdapterEvidence,
                    () => VbeBridgeClient.Read("VBAi." + ProcessId, request, 5000, 250, 1, 0), RetainUncertainOffice);
                steps.Add(new { PublisherQuitProjectPersistence = reply, BridgeAvailable = reply != null });
                Assert.IsNotNull(reply, "The identified Publisher project's bridge must be available before Quit.");
                Assert.AreEqual(true, reply["Ok"], "Publisher project persistence must be readable before Quit.");
                Assert.AreEqual(true, VbeBridgeClient.Object(reply["Data"])["ProjectSaved"],
                    "Publisher's VBA project must already be saved before Quit.");
            }
            object collection = null, window = null, current = null;
            try
            {
                RequireOwnedDocument();
                collection = ((dynamic)application).Documents;
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)collection).Count), "Publisher must expose only the owned publication before Quit.");
                current = ((dynamic)collection)[1];
                IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
                try
                {
                    first = Marshal.GetIUnknownForObject(current); second = Marshal.GetIUnknownForObject(document);
                    Assert.AreEqual(first, second, "Publisher's sole publication is not the owned document.");
                }
                finally { if (second != IntPtr.Zero) Marshal.Release(second); if (first != IntPtr.Zero) Marshal.Release(first); }
                Assert.IsTrue((bool)((dynamic)current).Saved, "Publisher must already be saved before Quit.");
                window = ((dynamic)application).ActiveWindow;
                uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).hWnd)), out pid);
                Assert.AreEqual((uint)ProcessId, pid, "The Publisher application no longer belongs to the retained process.");
                RequireOwnedDocument();
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)collection).Count), "Publisher's publication collection changed before Quit.");
                Assert.IsTrue((bool)((dynamic)document).Saved, "Publisher changed before Quit.");
            }
            finally
            {
                // The collection item shares document's RCW; balance this acquisition without invalidating the retained document.
                if (current != null && Marshal.IsComObject(current)) Marshal.ReleaseComObject(current);
                Release(window); Release(collection);
            }
        }

        /// <summary>Verifies the external COM application and disposable document before native close/reopen.</summary>
        private void RequireOwnedDocument()
        {
            commandContainment.RequireTerminal();
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited, "The retained owned process must still be alive.");
            object current = null, window = null;
            try
            {
                IntPtr handle;
                if (Kind == "Access")
                {
                    current = ((dynamic)application).CurrentProject;
                    handle = new IntPtr(Convert.ToInt64(((dynamic)application).hWndAccessApp()));
                }
                else
                {
                    current = ((dynamic)application).ActiveDocument;
                    window = ((dynamic)application).ActiveWindow;
                    handle = new IntPtr(Convert.ToInt64(((dynamic)window).hWnd));
                    IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
                    try
                    {
                        first = Marshal.GetIUnknownForObject(current); second = Marshal.GetIUnknownForObject(document);
                        Assert.AreEqual(first, second, "The active Publisher document changed.");
                    }
                    finally { if (second != IntPtr.Zero) Marshal.Release(second); if (first != IntPtr.Zero) Marshal.Release(first); }
                }
                uint pid; GetWindowThreadProcessId(handle, out pid); Assert.AreEqual((uint)ProcessId, pid);
                Assert.AreEqual(Path.GetFullPath(DocumentPath), Path.GetFullPath((string)((dynamic)current).FullName), true,
                    "The native document is not the owned disposable file.");
            }
            finally
            {
                // These getters can share document's RCW; balance acquisitions instead of invalidating it.
                if (current != null && Marshal.IsComObject(current)) Marshal.ReleaseComObject(current);
                Release(window);
            }
        }

        /// <summary>Closes only loaded objects of the verified disposable Access database with explicit acSaveNo.</summary>
        private void CloseAccessObjectsWithoutSaving()
        {
            RequireOwnedDocument();
            object current = null, commands = null;
            try
            {
                current = ((dynamic)application).CurrentProject; commands = ((dynamic)application).DoCmd;
                foreach (var kind in new[] { 2, 3, 5 }) // acForm, acReport, acModule
                {
                    object collection = null;
                    var names = new List<string>();
                    try
                    {
                        collection = kind == 2 ? (object)((dynamic)current).AllForms : kind == 3 ? (object)((dynamic)current).AllReports : (object)((dynamic)current).AllModules;
                        // A COM foreach creates an unnamed IEnumVARIANT acquisition that can
                        // keep Access alive until testhost exits. Balance explicit indexed reads.
                        int objectCount = Convert.ToInt32(((dynamic)collection).Count);
                        Assert.IsTrue(objectCount >= 0 && objectCount <= 1000, "Unexpected owned Access object count.");
                        for (int index = 0; index < objectCount; index++)
                        {
                            object item = null;
                            try
                            {
                                item = ((dynamic)collection)[index]; // Access AllObjects collections are zero-based.
                                if ((bool)((dynamic)item).IsLoaded) names.Add((string)((dynamic)item).Name);
                            }
                            finally { Release(item); }
                        }
                    }
                    finally { Release(collection); }
                    foreach (string name in names)
                    {
                        RequireOwnedDocument();
                        ((dynamic)commands).Close(kind, name, 2); // acSaveNo; never allow a close helper to persist the test edit.
                        object verificationCollection = null, closedObject = null;
                        try
                        {
                            verificationCollection = kind == 2 ? (object)((dynamic)current).AllForms : kind == 3 ? (object)((dynamic)current).AllReports : (object)((dynamic)current).AllModules;
                            closedObject = ((dynamic)verificationCollection)[name];
                            Assert.IsFalse((bool)((dynamic)closedObject).IsLoaded, "Access did not close the object with acSaveNo: " + name);
                        }
                        finally { Release(closedObject); Release(verificationCollection); }
                        steps.Add(new { AccessObjectClosedWithoutSave = name, ObjectType = kind, SaveOption = 2 });
                    }
                }
            }
            finally { Release(commands); Release(current); }
        }
        private void Release(object item)
        {
            if (commandContainment.Pending || commandContainment.Uncertain)
            {
                if (item != null) retainedDiagnosticReferences.Add(item);
                return;
            }
            if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
        }
    }
}
