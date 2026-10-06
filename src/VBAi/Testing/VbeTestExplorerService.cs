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

        /// <summary>Live VBE object accessed only on the thread that constructed this service.</summary>
        private readonly dynamic vbe;

        /// <summary>WinForms control used to marshal asynchronous continuations to the owning VBE thread.</summary>
        private readonly Control dispatcher;

        /// <summary>Managed thread ID captured at construction and required for every live VBE operation.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Known project COM identities keyed by stable explorer project IDs.</summary>
        private readonly Dictionary<string, object> projects = new Dictionary<string, object>();

        /// <summary>Active or recently completed runs keyed by their caller-visible run IDs.</summary>
        private readonly Dictionary<string, RunEntry> runs = new Dictionary<string, RunEntry>();

        /// <summary>Orchestrates setup, test, and cleanup phases through this execution host.</summary>
        private readonly VbaTestRunner runner;

        /// <summary>Disposed state and fail-closed latch set when native mutation outcome becomes uncertain.</summary>
        private bool disposed, outcomeUnknown;

        /// <summary>Run currently authorized to issue native test callbacks.</summary>
        private RunEntry active;

        /// <summary>Calls awaiting completion of native callbacks dispatched to the owner thread.</summary>
        private readonly List<PendingCall> pendingCalls = new List<PendingCall>();

        /// <summary>Per-user directory used to retain pre-install copies of test support modules.</summary>
        internal Func<string> BackupRoot = () => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "TestSupportBackups");

        /// <summary>Returned-value adapter selected for the current Office process.</summary>
        internal VbeDebug.IProcedureValuesHost Host = CreateReturnedValuesHost(CurrentProcessName());

        /// <summary>Overrideable check indicating whether the current process supports returned-value execution.</summary>
        internal Func<bool> IsExecutionHost = () => IsReturnedValuesHost(CurrentProcessName());

        /// <summary>Overrideable registration check; non-null text blocks native dispatch.</summary>
        internal Func<string> NativeRuntimeReason = () => NativeRuntimeRegistrationReason();

        /// <summary>Reads the executable name of the current process without its path or extension.</summary>
        /// <returns>Current process name.</returns>
        private static string CurrentProcessName()
        { using (var process = Process.GetCurrentProcess()) return process.ProcessName; }

        /// <summary>Checks whether the process is one of the Office hosts with a returned-value adapter.</summary>
        /// <param name="processName">Executable name, compared case-insensitively.</param>
        /// <returns>True for Excel, PowerPoint, or Word.</returns>
        internal static bool IsReturnedValuesHost(string processName) => processName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)
            || processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase);

        /// <summary>Selects the returned-value adapter for the current Office executable.</summary>
        /// <param name="processName">Executable name, compared case-insensitively.</param>
        /// <returns>Word or PowerPoint adapter for those hosts; otherwise the generic native values host.</returns>
        internal static VbeDebug.IProcedureValuesHost CreateReturnedValuesHost(string processName)
        {
            if (processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase)) return new VbaTestWordValuesHost();
            return processName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase)
                ? (VbeDebug.IProcedureValuesHost)new VbaTestPowerPointValuesHost() : new VbeDebug.NativeProcedureValuesHost();
        }

        /// <summary>Confirmation callback for reviewing existing test-support code before replacement.</summary>
        internal Func<VbaTestCatalog, string, string, bool> ConfirmSupport;

        /// <summary>Callback that opens the test explorer and returns its disposable window lifetime.</summary>
        internal Func<string, object> ShowExplorer;

        /// <summary>Overrideable check for a process that can dispatch native callbacks through VBE.</summary>
        internal Func<bool> IsNativeExecutionHost = () => {
            using (var process = Process.GetCurrentProcess()) return new[] { "EXCEL", "WINWORD", "POWERPNT", "MSACCESS", "MSPUB", "OUTLOOK", "SLDWORKS", "VISIO", "WINPROJ" }
                .Contains(process.ProcessName.ToUpperInvariant());
        };

        /// <summary>Owner-thread dispatcher for the registered native test callback.</summary>
        private VbaNativeTestExecutionHost nativeExecutionHost;

        /// <summary>Validates callback run identity and phase before accepting native result receipts.</summary>
        private readonly VbaTestResultSink resultSink = new VbaTestResultSink();

        /// <summary>Phase currently authorized for native result callbacks, or null between dispatches.</summary>
        private string nativePhase;

        /// <summary>Mutable cancellation and completion state retained for one test run.</summary>
        private sealed class RunEntry
        {

            /// <summary>Public run record containing IDs, status, and accumulated results.</summary>
            internal VbaTestRun Run;

            /// <summary>Stable project ID against which this run was authorized.</summary>
            internal string ProjectId;

            /// <summary>Lifecycle label used while the run is in progress or awaiting cleanup.</summary>
            internal string State = "Running";

            /// <summary>Cooperative cancellation source checked between safe execution phases.</summary>
            internal CancellationTokenSource Stop = new CancellationTokenSource();

            /// <summary>Task representing completion of the full run, including cleanup.</summary>
            internal Task<VbaTestRun> Completion;

            /// <summary>Revalidates run authorization and project identity before native dispatch.</summary>
            internal Action ExecutionGuard;

            /// <summary>Whether this run uses the host's returned-value execution adapter.</summary>
            internal bool ReturnedValues;
        }

        /// <summary>Completion receipt for one callback posted to the owner-thread dispatcher.</summary>
        private sealed class PendingCall
        {

            /// <summary>Asynchronous result completed by the native callback or dispatch failure.</summary>
            internal readonly TaskCompletionSource<VbaTestResult> Completion = new TaskCompletionSource<VbaTestResult>();

            /// <summary>True after the callback entered; distinguishes pre-dispatch failure from uncertain completion.</summary>
            internal bool Invoked;
        }

        /// <summary>Creates a session-bound test service on the thread that owns the supplied VBE.</summary>
        /// <param name="vbe">Live VBE automation object; all access remains on the constructing thread.</param>
        /// <param name="dispatcher">Created WinForms control used to schedule continuations on that thread.</param>
        /// <param name="continuationFactory">Optional factory for additional owner-thread continuation controls.</param>
        internal VbeTestExplorerService(object vbe, Control dispatcher, Func<Control> continuationFactory = null)
        {
            this.vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
            this.dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));
            InitializeOwnerContinuations(continuationFactory ?? (() => new Control()));
            runner = new VbaTestRunner(this);
            defaultCompileCoverageDelegate = CompileCoverageClone;
            CompileCoverageProject = defaultCompileCoverageDelegate;
            ConfirmSupport = (catalog, before, after) => TestSupportReviewDialog.Confirm(dispatcher, catalog.Project.Name, before, after);
            nativeExecutionHost = new VbaNativeTestExecutionHost(vbe, dispatcher, resultSink,
                catalog => ResolveLive(catalog.Project.Id), Validate, NativeExecutionGuard,
                VbaTestRuntimeSource.DispatchSignature, () => active?.Run.Id ?? throw new InvalidOperationException("No authorized test run is active."));
        }

        /// <summary>Rejects disposed services and calls made outside the VBE-owning thread.</summary>
        private void RequireOwner()
        {
            if (disposed || dispatcher.IsDisposed) throw new ObjectDisposedException(nameof(VbeTestExplorerService));
            if (Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("Test operations require the owning VBE thread.");
        }

        /// <summary>Captures each currently open VBA project and removes identities for projects that have closed.</summary>
        /// <returns>Project snapshots; projects that cannot be fully read remain listed with no module snapshots.</returns>
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

        /// <summary>Reuses or assigns a session-local ID for one COM project identity.</summary>
        /// <param name="project">Live VBProject object to identify.</param>
        /// <returns>Opaque ID retained while the same project remains open in this service.</returns>
        private string FindIdentity(object project)
        {
            foreach (var entry in projects)
                if (ReferenceEquals(entry.Value, project) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(entry.Value, project)) return entry.Key;
            string id = Guid.NewGuid().ToString("N");
            projects.Add(id, project);
            return id;
        }

        /// <summary>Resolves a project selector and discovers tests from its current source snapshot.</summary>
        /// <param name="selector">Host path or project selector accepted by <see cref="VbeProjectResolver"/>.</param>
        /// <returns>Catalog bound to the selected project's source and reference revision.</returns>
        internal VbaTestCatalog DiscoverSelector(string selector)
        {
            RequireOwner();
            object project = VbeProjectResolver.Resolve(vbe, selector);
            return Discover(FindIdentity(project));
        }

        /// <summary>Discovers eligible procedures in the exact live project previously assigned this ID.</summary>
        /// <param name="projectId">Session-local project ID returned in a project snapshot.</param>
        /// <returns>Discovery catalog with a revision fingerprint over modules and references.</returns>
        public VbaTestCatalog Discover(string projectId)
        {
            RequireOwner();
            object project = ResolveLive(projectId);
            return VbaTestDiscovery.Discover(Snapshot(projectId, project));
        }

        /// <summary>Reacquires the exact project identity from the current VBE collection.</summary>
        /// <param name="id">Session-local ID previously assigned to the project.</param>
        /// <returns>Current live VBProject COM object.</returns>
        /// <exception cref="InvalidOperationException">The ID is unknown or its original project was closed or replaced.</exception>
        private object ResolveLive(string id)
        {
            if (string.IsNullOrWhiteSpace(id) || !projects.TryGetValue(id, out object wanted)) throw new InvalidOperationException("The selected test project is no longer available.");
            foreach (object project in vbe.VBProjects)
                if (ReferenceEquals(project, wanted) || VbeDebug.NativeProcedureValuesHost.SameComIdentity(project, wanted)) return project;
            throw new InvalidOperationException("The selected project was closed or replaced.");
        }

        /// <summary>Reads source, component types, host path, and references into a revision-bound project snapshot.</summary>
        /// <param name="id">Session-local identity associated with this COM project.</param>
        /// <param name="project">Live project read on the owning VBE thread.</param>
        /// <returns>Snapshot with module hashes and a revision that changes when source or references change.</returns>
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

        /// <summary>Computes a SHA-256 fingerprint of UTF-8 source text.</summary>
        /// <param name="source">Text to fingerprint; null is treated as empty.</param>
        /// <returns>Uppercase hexadecimal SHA-256 string.</returns>
        internal static string Hash(string source)
        { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(source ?? ""))).Replace("-", ""); }

        /// <summary>Normalizes line endings and removes trailing LF characters for generated-module comparison.</summary>
        /// <param name="source">Source text, with null treated as empty.</param>
        /// <returns>Text using LF line endings without trailing LF.</returns>
        private static string Canonical(string source) => (source ?? "").Replace("\r\n", "\n").TrimEnd('\n');

        /// <summary>Reports the first current host, support-module, revision, or native-registration blocker for execution.</summary>
        /// <param name="catalog">Discovery catalog whose project and runtime module are being validated.</param>
        /// <returns>Null when execution is available; otherwise a user-displayable refusal reason.</returns>
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
                if (returnedValues) ReleaseReturnedTarget(Host.ResolveTarget(ResolveLive(catalog.Project.Id), catalog.Project.HostPath), false);
                else
                {
                    string registration = NativeRuntimeReason();
                    if (registration != null) return registration;
                    nativeExecutionHost.Validate(catalog);
                }
            }
            catch (Exception error) { return error.Message; }
            return null;
        }

        /// <summary>Requires an unchanged design-mode project and a run outcome that is not uncertain.</summary>
        /// <param name="catalog">Catalog whose project ID and revision must still match the live project.</param>
        /// <exception cref="VbaTestInvocationException">The previous outcome is uncertain, mode changed, or the source/reference revision is stale.</exception>
        public void Validate(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (outcomeUnknown) throw new VbaTestInvocationException("The preceding native outcome is uncertain; inspect the host before reconnecting.", false);
            dynamic project = ResolveLive(catalog.Project.Id);
            if ((int)project.Mode != 2) throw new VbaTestInvocationException("The selected VBA project is no longer in design mode.", false);
            if (Snapshot(catalog.Project.Id, project).Revision != catalog.Project.Revision)
                throw new VbaTestInvocationException("The project changed since test discovery. Refresh before execution.", false);
        }

        /// <summary>Routes the authorized procedure phase through native VBE dispatch or a supported returned-value host.</summary>
        /// <param name="catalog">Current catalog for the active run's project and revision.</param>
        /// <param name="procedure">Discovered test or fixture descriptor selected by the runner.</param>
        /// <param name="phase">Lifecycle phase used to label the result and apply cleanup cancellation rules.</param>
        /// <returns>Task completed by the owner-thread invocation and result validation.</returns>
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
                object target = null;
                try
                {
                    entry.ExecutionGuard?.Invoke();
                    Validate(catalog);
                    string reason = ExecutionUnavailableReason(catalog);
                    if (reason != null) throw new VbaTestInvocationException(reason, false);
                    target = Host.ResolveTarget(ResolveLive(catalog.Project.Id), catalog.Project.HostPath);
                    Validate(catalog);
                    entry.ExecutionGuard?.Invoke();
                    if (entry.Stop.IsCancellationRequested && phase != "TestCleanup" && phase != "ModuleCleanup")
                        throw new VbaTestInvocationException("The test run was stopped before this call was dispatched.", false);
                    invoked = true;
                    pending.Invoked = true;
                    object returned = Host.Invoke(target, VbaTestRuntimeSource.ModuleName, VbaTestRuntimeSource.DispatcherProcedure, new object[] { procedure.Module, procedure.Procedure });
                    var result = VbaTestRuntimeSource.Decode(procedure, returned);
                    result.Phase = phase;
                    ReleaseReturnedTarget(target, false); target = null;
                    completion.TrySetResult(result);
                }
                catch (Exception error)
                {
                    if (invoked) outcomeUnknown = true;
                    Exception failure = error;
                    try { ReleaseReturnedTarget(target, invoked); }
                    catch (Exception releaseError) { failure = new AggregateException(error.Message + Environment.NewLine + releaseError.Message, error, releaseError); }
                    completion.TrySetException(new VbaTestInvocationException(failure.Message, invoked, failure));
                }
                finally { pendingCalls.Remove(pending); }
            }));
            return completion.Task;
        }

        /// <summary>Disposes a Word target or retains its COM references when invocation completion is uncertain.</summary>
        /// <param name="target">Owned host target returned by its adapter.</param>
        /// <param name="uncertain">Whether the target must remain rooted instead of being released.</param>
        private static void ReleaseReturnedTarget(object target, bool uncertain)
        {
            if (target is VbaTestWordValuesHost.OwnedTarget word)
            {
                if (uncertain) word.RetainOnUncertain();
                word.Dispose();
            }
        }

        /// <summary>Starts the selected discovery-catalog tests on the owning VBE thread and returns their run task.</summary>
        /// <param name="catalog">Current project discovery snapshot and revision.</param>
        /// <param name="tests">Explicit test descriptors from the current catalog.</param>
        /// <param name="onResult">Optional callback invoked for each completed test result.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>Task representing the selected tests and final run state.</returns>
        public Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation)
        { return BeginRun(catalog, tests, onResult, cancellation, null); }

        /// <summary>Validates and records an explicit test selection before starting one run.</summary>
        /// <param name="catalog">Current project discovery snapshot and revision.</param>
        /// <param name="tests">Explicit test descriptors from the current catalog.</param>
        /// <param name="onResult">Optional callback invoked for each completed test result.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <param name="executionGuard">Authorization and revision check rerun at guarded execution phases.</param>
        /// <param name="measureCoverage">Indicates whether measure coverage is enabled.</param>
        /// <returns>Canonical task stored in the run entry, including results and terminal state.</returns>
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

        /// <summary>Executes the run, streams results, records terminal state, and releases active-run ownership.</summary>
        /// <param name="entry">Entry whose stop token, result object, and execution state are updated.</param>
        /// <param name="catalog">Current project discovery snapshot and revision.</param>
        /// <param name="tests">Explicit test descriptors from the current catalog.</param>
        /// <param name="progress">Callback for publishing each result as it is accepted.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <param name="measureCoverage">Indicates whether measure coverage is enabled.</param>
        /// <returns>Completed run task; failures are recorded as aborted and rethrown.</returns>
        private async Task<VbaTestRun> CompleteRun(RunEntry entry, VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> progress, CancellationToken cancellation, bool measureCoverage)
        {
            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellation, entry.Stop.Token))
            {
                try
                {
                    // Return the query before native copy preparation enters the owning thread.
                    if (measureCoverage) await AwaitOwner(QueueCoverageStart());
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

        /// <summary>Starts a run from an explicit project selector, expected revision, and set of discovered test IDs.</summary>
        /// <param name="selector">Project selector resolved through the current VBE collection.</param>
        /// <param name="revision">Expected discovery revision; a mismatch refuses the run.</param>
        /// <param name="ids">Nonempty list of IDs from the current discovery catalog.</param>
        /// <param name="executionGuard">Optional authorization callback rerun before execution phases.</param>
        /// <param name="measureCoverage">Whether to execute in a disposable instrumented host-document copy.</param>
        /// <returns>Initial compact run-status object; later results are read by the generated run ID.</returns>
        internal object StartRun(string selector, string revision, string[] ids, Action executionGuard = null, bool measureCoverage = false)
        {
            var catalog = DiscoverSelector(selector);
            if (string.IsNullOrEmpty(revision) || catalog.Project.Revision != revision) throw new InvalidOperationException("ExpectedProjectVersion must match the discovered test revision.");
            if (ids == null || ids.Length == 0) throw new ArgumentException("Items must explicitly select discovered test IDs.");
            var selected = ids.Select(id => catalog.Tests.Single(test => test.Id.Equals(id, StringComparison.OrdinalIgnoreCase))).ToArray();
            executionGuard?.Invoke();
            var task = BeginRun(catalog, selected, null, CancellationToken.None, executionGuard, measureCoverage);
            var entry = active ?? runs.Values.Single(item => ReferenceEquals(item.Completion, task));
            // Do not re-enter COM discovery after publishing the preparation continuation.
            return ReportRunStatus(catalog, entry.Run.Id, entry, "compact", 0, 0);
        }

        /// <summary>Posts a continuation so the caller returns before native coverage-copy preparation begins.</summary>
        /// <returns>Task completed on a later continuation-dispatcher turn.</returns>
        private Task<bool> QueueCoverageStart()
        {
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            try { continuationDispatcher.BeginInvoke(new Action(() => completion.TrySetResult(true))); }
            catch (Exception error) { completion.TrySetException(error); }
            return completion.Task;
        }

        /// <summary>Reads a run status page only when the run belongs to the selected project's current service session.</summary>
        /// <param name="selector">Current project selector used to verify run ownership.</param>
        /// <param name="id">Run ID returned by <see cref="StartRun"/>.</param>
        /// <param name="format">Either <c>human</c> or <c>compact</c>.</param>
        /// <param name="offset">Nonnegative zero-based page offset.</param>
        /// <param name="limit">Page size from 1 through 100, or zero for the default.</param>
        /// <returns>Run state, pending/stale flags, and the requested report page.</returns>
        internal object RunStatus(string selector, string id, string format, int offset = 0, int limit = 0)
        {
            RequireOwner();
            VbaTestReports.PageLimit(offset, limit);
            var catalog = DiscoverSelector(selector);
            if (!runs.TryGetValue(id ?? "", out RunEntry entry) || entry.ProjectId != catalog.Project.Id) throw new InvalidOperationException("Unknown test run in this project/session.");
            return ReportRunStatus(catalog, id, entry, format, offset, limit);
        }

        /// <summary>Builds a query response with current run state and the requested human or compact page.</summary>
        /// <param name="catalog">Current project discovery snapshot and revision.</param>
        /// <param name="id">Run identifier returned by StartRun.</param>
        /// <param name="entry">Entry whose stop token, result object, and execution state are updated.</param>
        /// <param name="format">Requested report form: human text or compact JSON.</param>
        /// <param name="offset">Zero-based item offset for the report page.</param>
        /// <param name="limit">Page size, with zero selecting the default.</param>
        /// <returns>Status object containing run state, stale/pending flags, and serialized report page.</returns>
        private static object ReportRunStatus(VbaTestCatalog catalog, string id, RunEntry entry, string format, int offset, int limit)
        {
            return new { Query = id, entry.State, Pending = entry.State == "Running" || entry.State == "StopRequested",
                Stale = catalog.Project.Revision != entry.Run.Revision,
                Report = format == "human" ? (object)VbaTestReports.HumanPage(entry.Run, offset, limit) : new System.Web.Script.Serialization.JavaScriptSerializer()
                    { MaxJsonLength = 10 * 1024 * 1024 }.DeserializeObject(VbaTestReports.CompactPage(entry.Run, offset, limit)) };
        }

        /// <summary>Requests cooperative cancellation for the matching active run.</summary>
        /// <param name="selector">Project selector that must resolve to the run's project identity.</param>
        /// <param name="id">Run ID to stop.</param>
        /// <returns>Current compact run-status response; cleanup phases may still execute.</returns>
        internal object StopRun(string selector, string id)
        {
            var catalog = DiscoverSelector(selector);
            if (!runs.TryGetValue(id ?? "", out RunEntry entry) || entry.ProjectId != catalog.Project.Id) throw new InvalidOperationException("Unknown test run in this project/session.");
            if (entry.State == "Running") { entry.State = "StopRequested"; entry.Stop.Cancel(); }
            return RunStatus(selector, id, "compact");
        }

        /// <summary>Opens the test module at its discovered line after confirming the project revision is unchanged.</summary>
        /// <param name="catalog">Discovery catalog providing project ID and expected revision.</param>
        /// <param name="test">Discovered procedure whose source line should be selected.</param>
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

        /// <summary>Shows the proposed generated runtime module for explicit review before project-local installation.</summary>
        /// <param name="catalog">Catalog used to generate and revision-check the support module.</param>
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

        /// <summary>Returns the generated support source and expected project revision without modifying the project.</summary>
        /// <param name="selector">Project selector resolved through the current VBE collection.</param>
        /// <returns>Preview data containing support module name, source, and coverage status.</returns>
        internal object PreviewSupport(string selector)
        {
            var catalog = DiscoverSelector(selector);
            return new { Project = selector, ExpectedProjectVersion = catalog.Project.Revision, Module = VbaTestRuntimeSource.ModuleName,
                Text = VbaTestRuntimeSource.Generate(catalog), Coverage = "Unavailable" };
        }

        /// <summary>Bridge operations; assistant permission checks remain in the tool gateway.</summary>
        /// <param name="request">Validated bridge request routed to discovery, preview, run, status, stop, or explorer display.</param>
        /// <param name="executionGuard">Optional authorization callback for commands that can dispatch tests.</param>
        /// <returns>Command-specific response; permission checks are performed by the gateway before this method.</returns>
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

        /// <summary>Applies explicitly approved runtime source to the project after identity and revision checks.</summary>
        /// <param name="catalog">Discovery snapshot whose project must remain unchanged.</param>
        /// <param name="approvedSource">Exact source previously reviewed by the caller.</param>
        /// <returns>Updated discovery catalog after insertion/replacement and readback.</returns>
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

        /// <summary>Marks the session unavailable, cancels active work, and releases continuation resources when idle.</summary>
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
