using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Session-owned live VBE test operations shared by the UI and guarded tool routes.</summary>
    internal sealed partial class VbeTestExplorerService : IVbaTestExplorerService, IVbaTestExecutionHost, IVbaTestCoverageExplorerService, IDisposable
    {
        private readonly dynamic vbe;
        private readonly Control dispatcher;
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private readonly Dictionary<string, object> projects = new Dictionary<string, object>();
        private readonly Dictionary<string, RunEntry> runs = new Dictionary<string, RunEntry>();
        private readonly VbaTestRunner runner;
        private bool disposed, outcomeUnknown;
        private RunEntry active;
        private readonly List<PendingCall> pendingCalls = new List<PendingCall>();
        internal Func<string> BackupRoot = () => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "TestSupportBackups");
        internal VbeDebug.IProcedureValuesHost Host = CreateReturnedValuesHost();
        internal Func<bool> IsExecutionHost = () => { using (var process = Process.GetCurrentProcess()) return process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
            || process.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)
            || process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase); };

        private static VbeDebug.IProcedureValuesHost CreateReturnedValuesHost()
        {
            using (var process = Process.GetCurrentProcess())
            {
                if (process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return new VbaTestWordValuesHost();
                return process.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)
                    ? (VbeDebug.IProcedureValuesHost)new VbaTestPowerPointValuesHost() : new VbeDebug.NativeProcedureValuesHost();
            }
        }
        internal Func<VbaTestCatalog, string, string, bool> ConfirmSupport;
        internal Func<string, object> ShowExplorer;
        internal Func<bool> IsNativeExecutionHost = () => {
            using (var process = Process.GetCurrentProcess()) return new[] { "EXCEL", "WINWORD", "POWERPNT", "MSACCESS", "MSPUB", "OUTLOOK", "SLDWORKS", "VISIO", "WINPROJ" }
                .Contains(process.ProcessName.ToUpperInvariant());
        };
        private VbaNativeTestExecutionHost nativeExecutionHost;
        private readonly VbaTestResultSink resultSink = new VbaTestResultSink();
        private string nativePhase;

        private sealed class RunEntry
        {
            internal VbaTestRun Run;
            internal string ProjectId;
            internal string State = "Running";
            internal CancellationTokenSource Stop = new CancellationTokenSource();
            internal Task<VbaTestRun> Completion;
            internal Action ExecutionGuard;
            internal bool ReturnedValues;
        }

        private sealed class PendingCall
        {
            internal readonly TaskCompletionSource<VbaTestResult> Completion = new TaskCompletionSource<VbaTestResult>();
            internal bool Invoked;
        }

        internal VbeTestExplorerService(object vbe, Control dispatcher)
        {
            this.vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
            this.dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            InitializeOwnerContinuations();
            runner = new VbaTestRunner(this);
            defaultCompileCoverageDelegate = CompileCoverageClone;
            CompileCoverageProject = defaultCompileCoverageDelegate;
            ConfirmSupport = (catalog, before, after) => TestSupportReviewDialog.Confirm(dispatcher, catalog.Project.Name, before, after);
            nativeExecutionHost = new VbaNativeTestExecutionHost(vbe, dispatcher, resultSink,
                catalog => ResolveLive(catalog.Project.Id), Validate, NativeExecutionGuard,
                VbaTestRuntimeSource.DispatchSignature, () => active?.Run.Id ?? throw new InvalidOperationException("No authorized test run is active."));
        }

        private void RequireOwner()
        {
            if (disposed || dispatcher.IsDisposed) throw new ObjectDisposedException(nameof(VbeTestExplorerService));
            if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("Test operations require the owning VBE thread.");
        }

        public IReadOnlyList<VbaTestProjectSnapshot> ReadProjects()
        {
            RequireOwner();
            var result = new List<VbaTestProjectSnapshot>();
            var liveIds = new HashSet<string>();
            foreach (object project in vbe.VBProjects)
            {
                string id = FindIdentity(project);
                liveIds.Add(id);
                try { result.Add(Snapshot(id, project)); }
                catch (Exception)
                {
                    result.Add(new VbaTestProjectSnapshot { Id = id, Name = Convert.ToString(((dynamic)project).Name), Selector = id,
                        Modules = new VbaTestModuleSnapshot[0] });
                }
            }
            foreach (string id in projects.Keys.Where(id => !liveIds.Contains(id)).ToArray()) projects.Remove(id);
            return result;
        }

        private string FindIdentity(object project)
        {
            foreach (var entry in projects)
                if (ReferenceEquals(entry.Value, project) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(entry.Value, project)) return entry.Key;
            string id = Guid.NewGuid().ToString("N");
            projects.Add(id, project);
            return id;
        }

        internal VbaTestCatalog DiscoverSelector(string selector)
        {
            RequireOwner();
            object project = VbeProjectResolver.Resolve(vbe, selector);
            return Discover(FindIdentity(project));
        }

        public VbaTestCatalog Discover(string projectId)
        {
            RequireOwner();
            object project = ResolveLive(projectId);
            return VbaTestDiscovery.Discover(Snapshot(projectId, project));
        }

        private object ResolveLive(string id)
        {
            if (string.IsNullOrWhiteSpace(id) || !projects.TryGetValue(id, out object wanted)) throw new InvalidOperationException("The selected test project is no longer available.");
            foreach (object project in vbe.VBProjects)
                if (ReferenceEquals(project, wanted) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, wanted)) return project;
            throw new InvalidOperationException("The selected project was closed or replaced.");
        }

        private VbaTestProjectSnapshot Snapshot(string id, dynamic project)
        {
            if ((int)project.Protection != 0) throw new InvalidOperationException("The selected VBA project is protected.");
            string name = (string)project.Name, path = null;
            try { path = VbeProjectHostPath.Read((object)project); } catch (Exception) { }
            var modules = new List<VbaTestModuleSnapshot>();
            var identity = new StringBuilder().Append(id).Append('\n').Append(name).Append('\n').Append(path).Append('\n');
            foreach (dynamic component in project.VBComponents)
            {
                dynamic code = component.CodeModule;
                int count = (int)code.CountOfLines;
                if (count > 200000 || modules.Count >= 1000) throw new InvalidOperationException("The project exceeds the test discovery limits.");
                string source = count == 0 ? "" : (string)code.Lines[1, count];
                modules.Add(new VbaTestModuleSnapshot { Name = (string)component.Name, Source = source, Hash = Hash(source), ComponentType = (int)component.Type });
            }
            foreach (var module in modules.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
                identity.Append(module.Name).Append('|').Append(module.ComponentType).Append('|').Append(module.Hash).Append('\n');
            // References are fingerprinted as well as code: a changed dependency invalidates the plan.
            var references = new StringBuilder();
            foreach (dynamic reference in project.References)
                references.Append((string)reference.Name).Append('|').Append((string)reference.Guid).Append('|')
                    .Append((int)reference.Major).Append('|').Append((int)reference.Minor).Append('|').Append((bool)reference.IsBroken).Append('\n');
            identity.Append(references);
            return new VbaTestProjectSnapshot { Id = id, Name = name, Selector = !string.IsNullOrEmpty(path) && Path.IsPathRooted(path) ? path : name,
                HostPath = path, Revision = Hash(identity.ToString()), ReferencesHash = Hash(references.ToString()), Modules = modules.ToArray() };
        }

        internal static string Hash(string source)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(source ?? ""))).Replace("-", ""); }

        private static string Canonical(string source) => (source ?? "").Replace("\r\n", "\n").TrimEnd('\n');

        public string ExecutionUnavailableReason(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (catalog == null) return "Select and refresh a VBA project.";
            if (outcomeUnknown) return "A native call has an uncertain outcome. Reconnect VBAi after inspecting the host; automatic retry is disabled.";
            bool returnedValues = active?.ReturnedValues ?? PreferReturnedValues();
            if (!returnedValues && !IsNativeExecutionHost()) return "Verified test execution requires an owned in-process VBE host. Discovery remains available.";
            if (returnedValues && (string.IsNullOrEmpty(catalog.Project.HostPath) || !Path.IsPathRooted(catalog.Project.HostPath))) return "Save the selected macro-enabled host document before running tests.";
            try
            {
                var support = catalog.Project.Modules.SingleOrDefault(module => module.Name.Equals(VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase));
                if (support == null) return "Review and install the project-local test support module before running tests.";
                if (Canonical(support.Source) != Canonical(VbaTestRuntimeSource.Generate(catalog))) return "The test support module is outdated or changed. Review its update before running tests.";
                Validate(catalog);
                if (returnedValues) Host.ResolveTarget(ResolveLive(catalog.Project.Id), catalog.Project.HostPath);
                else
                {
                    string registration = NativeRuntimeRegistrationReason();
                    if (registration != null) return registration;
                    nativeExecutionHost.Validate(catalog);
                }
            }
            catch (Exception error) { return error.Message; }
            return null;
        }

        public void Validate(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (outcomeUnknown) throw new VbaTestInvocationException("The preceding native outcome is uncertain; inspect the host before reconnecting.", false);
            dynamic project = ResolveLive(catalog.Project.Id);
            if ((int)project.Mode != 2) throw new VbaTestInvocationException("The selected VBA project is no longer in design mode.", false);
            if (Snapshot(catalog.Project.Id, project).Revision != catalog.Project.Revision)
                throw new VbaTestInvocationException("The project changed since test discovery. Refresh before execution.", false);
        }

        public Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
        {
            RequireOwner();
            if (active == null || active.ProjectId != catalog.Project.Id)
                throw new VbaTestInvocationException("Only the active authorized test run may invoke a procedure.", false);
            if (!active.ReturnedValues) return InvokeNativeAsync(catalog, procedure, phase);
            var entry = active;
            var pending = new PendingCall();
            pendingCalls.Add(pending);
            var completion = pending.Completion;
            dispatcher.BeginInvoke(new Action(() => {
                if (completion.Task.IsCompleted) return;
                bool invoked = false;
                try
                {
                    entry.ExecutionGuard?.Invoke();
                    Validate(catalog);
                    string reason = ExecutionUnavailableReason(catalog);
                    if (reason != null) throw new VbaTestInvocationException(reason, false);
                    object target = Host.ResolveTarget(ResolveLive(catalog.Project.Id), catalog.Project.HostPath);
                    Validate(catalog);
                    entry.ExecutionGuard?.Invoke();
                    if (entry.Stop.IsCancellationRequested && phase != "TestCleanup" && phase != "ModuleCleanup")
                        throw new VbaTestInvocationException("The test run was stopped before this call was dispatched.", false);
                    invoked = true;
                    pending.Invoked = true;
                    object returned = Host.Invoke(target, VbaTestRuntimeSource.ModuleName, VbaTestRuntimeSource.DispatcherProcedure, new object[] { procedure.Module, procedure.Procedure });
                    var result = VbaTestRuntimeSource.Decode(procedure, returned);
                    result.Phase = phase;
                    completion.TrySetResult(result);
                }
                catch (Exception error)
                {
                    if (invoked) outcomeUnknown = true;
                    completion.TrySetException(new VbaTestInvocationException(error.Message, invoked, error));
                }
                finally { pendingCalls.Remove(pending); }
            }));
            return completion.Task;
        }

        public Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation)
        { return BeginRun(catalog, tests, onResult, cancellation, null); }

        private Task<VbaTestRun> BeginRun(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation, Action executionGuard, bool measureCoverage = false)
        {
            RequireOwner();
            string reason = measureCoverage ? CoverageUnavailableReason(catalog) : ExecutionUnavailableReason(catalog);
            if (reason != null) throw new InvalidOperationException(reason);
            if (active != null) throw new InvalidOperationException("A test run is already active in this session.");
            if (tests == null || tests.Count == 0) throw new ArgumentException("Select explicit discovered tests.");
            tests = tests.Select(test => catalog.Tests.Single(item => item.Id == test.Id && item.Module == test.Module && item.Procedure == test.Procedure))
                .GroupBy(test => test.Id).Select(group => group.First()).ToArray();
            var entry = new RunEntry { ProjectId = catalog.Project.Id, ExecutionGuard = executionGuard, ReturnedValues = PreferReturnedValues(),
                Run = new VbaTestRun { Id = Guid.NewGuid().ToString("N"), Project = catalog.Project.Name, Revision = catalog.Project.Revision } };
            runs.Add(entry.Run.Id, entry);
            foreach (string expired in runs.Keys.Where(id => id != entry.Run.Id).Take(Math.Max(0, runs.Count - 20)).ToArray()) { runs[expired].Stop.Dispose(); runs.Remove(expired); }
            active = entry;
            entry.Completion = CompleteRun(entry, catalog, tests, onResult, cancellation, measureCoverage);
            return entry.Completion;
        }

        private async Task<VbaTestRun> CompleteRun(RunEntry entry, VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> progress, CancellationToken cancellation, bool measureCoverage)
        {
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, entry.Stop.Token))
            {
                try
                {
                    Action<VbaTestResult> publish = result => {
                        entry.Run.Results.Add(result);
                        if (result.Outcome == VbaTestOutcome.OutcomeUnknown) entry.Run.OutcomeUnknown = true;
                        progress?.Invoke(result);
                    };
                    var run = measureCoverage
                        ? await AwaitOwner(ExecuteCoverageAsync(catalog, tests, publish, linked.Token, entry.ExecutionGuard))
                        : await AwaitOwner(runner.RunAsync(catalog, tests, publish, linked.Token));
                    run.Id = entry.Run.Id;
                    entry.Run = run;
                    entry.State = run.OutcomeUnknown ? "OutcomeUnknown" : linked.IsCancellationRequested ? "Cancelled" : run.Error != null ? "Aborted" : "Completed";
                    return run;
                }
                catch (Exception error) { entry.State = "Aborted"; entry.Run.Error = error.Message; throw; }
                finally { entry.ExecutionGuard = null; active = null; ReleaseOwnerContinuationsWhenIdle(); }
            }
        }

        internal object StartRun(string selector, string revision, string[] ids, Action executionGuard = null, bool measureCoverage = false)
        {
            var catalog = DiscoverSelector(selector);
            if (string.IsNullOrEmpty(revision) || catalog.Project.Revision != revision) throw new InvalidOperationException("ExpectedProjectVersion must match the discovered test revision.");
            if (ids == null || ids.Length == 0) throw new ArgumentException("Items must explicitly select discovered test IDs.");
            var selected = ids.Select(id => catalog.Tests.Single(test => test.Id.Equals(id, StringComparison.OrdinalIgnoreCase))).ToArray();
            executionGuard?.Invoke();
            var task = BeginRun(catalog, selected, null, CancellationToken.None, executionGuard, measureCoverage);
            var entry = active ?? runs.Values.Single(item => ReferenceEquals(item.Completion, task));
            return RunStatus(selector, entry.Run.Id, "compact");
        }

        internal object RunStatus(string selector, string id, string format, int offset = 0, int limit = 0)
        {
            RequireOwner();
            VbaTestReports.PageLimit(offset, limit);
            var catalog = DiscoverSelector(selector);
            if (!runs.TryGetValue(id ?? "", out RunEntry entry) || entry.ProjectId != catalog.Project.Id) throw new InvalidOperationException("Unknown test run in this project/session.");
            return new { Query = id, entry.State, Pending = entry.State == "Running" || entry.State == "StopRequested",
                Stale = catalog.Project.Revision != entry.Run.Revision,
                Report = format == "human" ? (object)VbaTestReports.HumanPage(entry.Run, offset, limit) : new System.Web.Script.Serialization.JavaScriptSerializer()
                    { MaxJsonLength = 10 * 1024 * 1024 }.DeserializeObject(VbaTestReports.CompactPage(entry.Run, offset, limit)) };
        }

        internal object StopRun(string selector, string id)
        {
            var catalog = DiscoverSelector(selector);
            if (!runs.TryGetValue(id ?? "", out RunEntry entry) || entry.ProjectId != catalog.Project.Id) throw new InvalidOperationException("Unknown test run in this project/session.");
            if (entry.State == "Running") { entry.State = "StopRequested"; entry.Stop.Cancel(); }
            return RunStatus(selector, id, "compact");
        }

        public void Navigate(VbaTestCatalog catalog, VbaTestDescriptor test)
        {
            RequireOwner();
            dynamic project = ResolveLive(catalog.Project.Id);
            var live = Snapshot(catalog.Project.Id, project);
            if (live.Revision != catalog.Project.Revision) throw new InvalidOperationException("The test source changed. Refresh before navigating.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, test.Module, StringComparison.OrdinalIgnoreCase))
                { dynamic pane = component.CodeModule.CodePane; pane.Show(); pane.SetSelection(test.Line, 1, test.Line, 1); return; }
            throw new InvalidOperationException("The test module is unavailable.");
        }

        public void InstallSupport(VbaTestCatalog catalog)
        {
            RequireOwner();
            Validate(catalog);
            if (active != null) throw new InvalidOperationException("Stop the active test run before updating test support.");
            string after = VbaTestRuntimeSource.Generate(catalog);
            string before = catalog.Project.Modules.FirstOrDefault(module => module.Name.Equals(VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase))?.Source ?? "";
            if (!ConfirmSupport(catalog, before, after)) return;
            ApplySupport(catalog, after);
        }

        internal object PreviewSupport(string selector)
        {
            var catalog = DiscoverSelector(selector);
            return new { Project = selector, ExpectedProjectVersion = catalog.Project.Revision, Module = VbaTestRuntimeSource.ModuleName,
                Text = VbaTestRuntimeSource.Generate(catalog), Coverage = "Unavailable" };
        }

        /// <summary>Bridge operations; assistant permission checks remain in the tool gateway.</summary>
        internal object Command(Request request, Action executionGuard = null)
        {
            RequireOwner();
            if (request.Command == "vba_test_run_status" || request.Command == "vba_test_coverage")
                VbaTestReports.PageLimit(request.Offset, request.Limit);
            if (request.Command == "show_vba_test_explorer")
            {
                var requestedProject = DiscoverSelector(request.Project);
                if (ShowExplorer == null) throw new InvalidOperationException("The test explorer UI is unavailable in this session.");
                return ShowExplorer(requestedProject.Project.Id);
            }
            if (request.Command == "vba_test_run_status")
            {
                if (!string.IsNullOrEmpty(request.Action) && request.Action != "human" && request.Action != "compact")
                    throw new ArgumentException("Action must be human or compact.");
                return RunStatus(request.Project, request.Query, request.Action, request.Offset, request.Limit);
            }
            if (request.Command == "stop_vba_tests") return StopRun(request.Project, request.Query);
            if (request.Command == "preview_vba_test_support") return PreviewSupport(request.Project);
            var catalog = DiscoverSelector(request.Project);
            if (request.Command == "discover_vba_tests") return new { Project = request.Project, ExpectedProjectVersion = catalog.Project.Revision,
                Modules = catalog.Modules.Select(module => new { module.Name, module.Diagnostic, module.ModuleInitialize, module.ModuleCleanup,
                    module.TestInitialize, module.TestCleanup, module.Tests }).ToArray(), catalog.Diagnostics,
                ExecutionUnavailableReason = ExecutionUnavailableReason(catalog), Coverage = PreviewCoverage(catalog) };
            if (request.Command == "vba_test_coverage")
            {
                if (!string.IsNullOrEmpty(request.Query))
                {
                    if (!runs.TryGetValue(request.Query, out RunEntry covered) || covered.ProjectId != catalog.Project.Id)
                        throw new InvalidOperationException("Unknown coverage run in this project/session.");
                    return new { Project = request.Project, covered.State, Stale = covered.Run.Revision != catalog.Project.Revision,
                        Uncertain = covered.Run.OutcomeUnknown, Error = covered.Run.Error,
                        Available = covered.Run.Coverage?.Available == true,
                        Report = VbaTestReports.CoveragePage(covered.Run.Coverage, request.Offset, request.Limit) };
                }
                return PreviewCoverage(catalog, request.Offset, request.Limit);
            }
            if (string.IsNullOrEmpty(request.ExpectedProjectVersion) || request.ExpectedProjectVersion != catalog.Project.Revision)
                throw new InvalidOperationException("ExpectedProjectVersion must match the discovered test project revision.");
            if (request.Command == "navigate_vba_test")
            {
                if (request.Items == null || request.Items.Length != 1) throw new ArgumentException("Items must identify exactly one discovered test.");
                var test = catalog.Tests.Single(item => item.Id.Equals(request.Items[0], StringComparison.OrdinalIgnoreCase));
                Navigate(catalog, test);
                return new { Navigated = true, test.Module, test.Procedure, test.Line };
            }
            if (request.ExpectedMode != 2) throw new ArgumentException("ExpectedMode=2 is required for a test mutation or run.");
            if (request.Command == "install_vba_test_support") return ApplySupport(catalog, request.Text);
            if (request.Command == "run_vba_tests")
            {
                if (!string.IsNullOrEmpty(request.Action) && request.Action != "coverage") throw new ArgumentException("Action must be coverage or omitted.");
                return StartRun(request.Project, request.ExpectedProjectVersion, request.Items, executionGuard, request.Action == "coverage");
            }
            throw new ArgumentException("Unknown VBA test command.");
        }

        internal object ApplySupport(VbaTestCatalog catalog, string approvedSource)
        {
            RequireOwner(); Validate(catalog);
            if (active != null) throw new InvalidOperationException("A test run is active.");
            string generated = VbaTestRuntimeSource.Generate(catalog);
            if (approvedSource != generated) throw new InvalidOperationException("Only the exact generated support source can be installed.");
            dynamic project = ResolveLive(catalog.Project.Id), component = null;
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase)) { component = candidate; break; }
            string before = component == null || (int)component.CodeModule.CountOfLines == 0 ? "" : (string)component.CodeModule.Lines[1, (int)component.CodeModule.CountOfLines];
            if (component != null && ((int)component.Type != 1 || !VbaTestRuntimeSource.IsOwned(before))) throw new InvalidOperationException("The support module name is occupied by user code.");
            string folder = Path.Combine(BackupRoot(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string backup = Path.Combine(folder, VbaTestRuntimeSource.ModuleName + ".bas");
            File.WriteAllText(backup, before, new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(folder, "proposed.bas"), generated, new UTF8Encoding(false));
            Validate(catalog);
            try
            {
                if (component == null) { component = project.VBComponents.Add(1); component.Name = VbaTestRuntimeSource.ModuleName; }
                dynamic code = component.CodeModule;
                int count = (int)code.CountOfLines;
                if (count > 0) code.DeleteLines(1, count);
                code.AddFromString(generated);
                string observed = (string)code.Lines[1, (int)code.CountOfLines];
                if (Canonical(observed) != Canonical(generated)) throw new InvalidOperationException("The support module did not match the reviewed source after writing.");
                return new { Applied = true, Backup = backup, Project = catalog.Project.Selector, SavedToDisk = false };
            }
            catch (Exception error)
            { throw new InvalidOperationException("Test support may be partially written. No automatic retry or restoration was attempted. Preserved backup: " + backup + ". " + error.Message, error); }
        }

        public void Dispose()
        {
            if (disposed) return;
            RequireContinuationOwner();
            active?.Stop.Cancel();
            disposed = true;
            coverageService?.Dispose();
            nativeExecutionHost?.Dispose(); nativeExecutionHost = null;
            resultSink.Dispose();
            foreach (var pending in pendingCalls.ToArray())
                pending.Completion.TrySetException(new VbaTestInvocationException("The VBE session disconnected before completion could be verified.", pending.Invoked));
            pendingCalls.Clear();
            projects.Clear();
            foreach (var entry in runs.Values) entry.Stop.Dispose();
            runs.Clear();
            ReleaseOwnerContinuationsWhenIdle();
        }
    }
}
