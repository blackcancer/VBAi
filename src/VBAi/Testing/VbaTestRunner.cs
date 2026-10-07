using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Validated owning-thread execution boundary; no retries are permitted.</summary>
    internal interface IVbaTestExecutionHost
    {

        /// <summary>Revalidates project identity, revision, and host authorization before native execution.</summary>
        /// <param name="catalog">The exact discovered catalog whose project snapshot is about to be used.</param>
        void Validate(VbaTestCatalog catalog);

        /// <summary>Invokes one already-discovered procedure on its owning host thread.</summary>
        /// <param name="catalog">Catalog whose project and revision were validated.</param>
        /// <param name="procedure">Test or fixture descriptor to invoke.</param>
        /// <param name="phase">Lifecycle phase label attached to the result.</param>
        /// <returns>Task with a verified result; failure after dispatch may represent an uncertain outcome.</returns>
        Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase);
    }

    /// <summary>A refused preflight is distinct from a possibly executed native call.</summary>
    internal sealed class VbaTestInvocationException : Exception
    {

        /// <summary>Gets whether the native operation may have executed before the failure was observed.</summary>
        /// <value>True forbids replay because the execution outcome is unknown.</value>
        internal bool Uncertain { get; }

        /// <summary>Captures the failure phase and whether a native invocation may have completed without a receipt.</summary>
        /// <param name="message">Diagnostic describing the refusal or uncertain native completion.</param>
        /// <param name="uncertain">True only when dispatch may have occurred without a verified receipt.</param>
        /// <param name="inner">Underlying failure, if one was caught.</param>
        internal VbaTestInvocationException(string message, bool uncertain, Exception inner = null) : base(message, inner)
        { Uncertain = uncertain; }
    }

    /// <summary>Serial test and fixture scheduling with cooperative cancellation.</summary>
    internal sealed class VbaTestRunner
    {

        /// <summary>Provides preflight and owning-thread invocation for this runner session.</summary>
        private readonly IVbaTestExecutionHost host;

        /// <summary>Prevents overlapping runs through the same runner instance.</summary>
        private bool active;

        /// <summary>Creates a runner bound to the host adapter that owns validation and native invocation.</summary>
        /// <param name="host">Execution adapter that validates and invokes procedures on the correct host thread.</param>
        internal VbaTestRunner(IVbaTestExecutionHost host) { this.host = host ?? throw new ArgumentNullException(nameof(host)); }

        /// <summary>Runs a distinct selection serially with module/test fixtures, revalidation, and no-retry uncertainty handling.</summary>
        /// <param name="catalog">Discovery catalog bound to the project identity and source revision.</param>
        /// <param name="selection">Distinct descriptors from this catalog; count must be between 1 and 10,000.</param>
        /// <param name="progress">Optional callback invoked as each descriptor receives a terminal result.</param>
        /// <param name="cancellation">Cooperative cancellation token checked between native calls.</param>
        /// <returns>Run receipt containing every selected descriptor's result and any run-level error.</returns>
        /// <exception cref="ArgumentException">The catalog or selection is invalid, empty, too large, or contains foreign/duplicate descriptors.</exception>
        /// <exception cref="InvalidOperationException">Another run is already active on this runner.</exception>
        internal async Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> selection,
            Action<VbaTestResult> progress, CancellationToken cancellation)
        {
            if (active) throw new InvalidOperationException("A test run is already active in this session.");
            if (catalog?.Project == null || selection == null || selection.Count == 0 || selection.Count > 10000)
                throw new ArgumentException("Select 1..10000 discovered tests.");
            var known = catalog.Tests.ToDictionary(test => test.Id, StringComparer.OrdinalIgnoreCase);
            if (selection.Any(test => test == null || !known.ContainsKey(test.Id)) ||
                selection.Select(test => test.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != selection.Count)
                throw new ArgumentException("The selection must contain distinct tests from this catalogue.");
            var selected = selection.Select(test => known[test.Id]).OrderBy(test => test.Module, StringComparer.OrdinalIgnoreCase)
                .ThenBy(test => test.Procedure, StringComparer.OrdinalIgnoreCase).ToArray();
            var run = new VbaTestRun { Id = Guid.NewGuid().ToString("N"), Project = catalog.Project.Name, Revision = catalog.Project.Revision };
            active = true;
            try
            {
                if (!cancellation.IsCancellationRequested) Validate(catalog);
                foreach (var group in selected.GroupBy(test => test.Module, StringComparer.OrdinalIgnoreCase))
                {
                    var module = catalog.Modules.Single(item => item.Name.Equals(group.Key, StringComparison.OrdinalIgnoreCase));
                    var tests = group.ToArray();
                    if (cancellation.IsCancellationRequested) break;
                    foreach (var test in tests.Where(test => !string.IsNullOrEmpty(module.Diagnostic) || !string.IsNullOrEmpty(test.Diagnostic) || !string.IsNullOrEmpty(test.IgnoreReason)))
                        Publish(run, new VbaTestResult
                        {
                            Test = test,
                            Outcome = !string.IsNullOrEmpty(module.Diagnostic) || !string.IsNullOrEmpty(test.Diagnostic)
                            ? VbaTestOutcome.Blocked : VbaTestOutcome.Skipped,
                            Message = FirstMessage(module.Diagnostic, test.Diagnostic, test.IgnoreReason),
                            Phase = "Discovery"
                        }, progress);
                    var runnable = tests.Where(test => !run.Results.Any(result => result.Test.Id == test.Id)).ToArray();
                    if (runnable.Length == 0) continue;
                    bool stopModule = false;
                    bool safeToInvoke = true;
                    bool moduleHasEntered = false;
                    try
                    {
                        var setup = await AwaitOwner(Fixture(catalog, module.ModuleInitialize, "ModuleInitialize"));
                        moduleHasEntered = module.ModuleInitialize != null;
                        if (setup != null && setup.Outcome != VbaTestOutcome.Passed)
                        {
                            foreach (var test in runnable) Publish(run, new VbaTestResult
                            {
                                Test = test,
                                Outcome = VbaTestOutcome.Blocked,
                                Message = setup.Message,
                                ErrorNumber = setup.ErrorNumber,
                                Phase = setup.Phase
                            }, progress);
                            AddRunError(run, "Module initialization failed: " + setup.Message);
                            continue;
                        }
                        foreach (var test in runnable)
                        {
                            if (cancellation.IsCancellationRequested || stopModule) break;
                            var clock = Stopwatch.StartNew();
                            VbaTestResult result = null;
                            bool safeCleanup = true;
                            bool testHasEntered = false;
                            bool cancelledRefusal = false;
                            string currentPhase = "TestInitialize";
                            try
                            {
                                var initialized = await AwaitOwner(Fixture(catalog, module.TestInitialize, "TestInitialize"));
                                testHasEntered = module.TestInitialize != null;
                                moduleHasEntered |= testHasEntered;
                                if (initialized != null && initialized.Outcome != VbaTestOutcome.Passed)
                                    result = new VbaTestResult
                                    {
                                        Test = test,
                                        Outcome = VbaTestOutcome.Error,
                                        Message = initialized.Message,
                                        ErrorNumber = initialized.ErrorNumber,
                                        Phase = initialized.Phase
                                    };
                                else
                                {
                                    currentPhase = "Test";
                                    result = await AwaitOwner(Invoke(catalog, test, "Test"));
                                    testHasEntered = true;
                                    moduleHasEntered = true;
                                }
                            }
                            catch (Exception error)
                            {
                                cancelledRefusal = cancellation.IsCancellationRequested
                                    && error is VbaTestInvocationException refused && !refused.Uncertain;
                                safeCleanup = cancelledRefusal && testHasEntered;
                                result = Refused(test, error, currentPhase, cancelledRefusal);
                            }
                            if (safeCleanup)
                            {
                                try
                                {
                                    var cleanup = await AwaitOwner(Fixture(catalog, module.TestCleanup, "TestCleanup"));
                                    if (cleanup != null && cleanup.Outcome != VbaTestOutcome.Passed)
                                    { AppendFailure(result, cleanup); stopModule = true; }
                                }
                                catch (Exception error)
                                { AppendFailure(result, Refused(test, error, "TestCleanup")); safeCleanup = false; cancelledRefusal = false; }
                            }
                            result.Duration = clock.Elapsed;
                            Publish(run, result, progress);
                            if ((!safeCleanup && !cancelledRefusal) || result.Outcome == VbaTestOutcome.OutcomeUnknown)
                            {
                                safeToInvoke = false;
                                throw new VbaTestInvocationException(result.Message, result.Outcome == VbaTestOutcome.OutcomeUnknown);
                            }
                        }
                    }
                    catch
                    {
                        // No further call is authorized after a refused preflight or
                        // uncertain fixture, even when no individual result exists yet.
                        safeToInvoke = false;
                        throw;
                    }
                    finally
                    {
                        // Revalidation must succeed before cleanup; uncertain calls never trigger another invocation.
                        if (safeToInvoke && (!cancellation.IsCancellationRequested || moduleHasEntered))
                        {
                            var cleanup = await AwaitOwner(Fixture(catalog, module.ModuleCleanup, "ModuleCleanup"));
                            if (cleanup != null && cleanup.Outcome != VbaTestOutcome.Passed)
                                AddRunError(run, "Module cleanup failed: " + cleanup.Message
                                    + (cleanup.ErrorNumber == 0 ? "" : " (VBA error: " + cleanup.ErrorNumber + ")"));
                        }
                    }
                    if (stopModule)
                        foreach (var test in runnable.Where(test => !run.Results.Any(result => result.Test.Id == test.Id)))
                            Publish(run, new VbaTestResult
                            {
                                Test = test,
                                Outcome = VbaTestOutcome.Blocked,
                                Phase = "TestCleanup",
                                Message = "A preceding test cleanup failed."
                            }, progress);
                }
            }
            catch (Exception error)
            {
                run.OutcomeUnknown |= error is VbaTestInvocationException native && native.Uncertain;
                AddRunError(run, error.Message);
            }
            finally
            {
                try
                {
                    foreach (var test in selected.Where(test => !run.Results.Any(result => result.Test.Id == test.Id)))
                        Publish(run, new VbaTestResult
                        {
                            Test = test,
                            Outcome = cancellation.IsCancellationRequested ? VbaTestOutcome.Cancelled : VbaTestOutcome.Blocked,
                            Message = run.Error ?? "Stopped before execution.",
                            Phase = "Scheduling"
                        }, progress);
                }
                finally { active = false; }
            }
            return run;
        }

        /// <summary>Invokes a lifecycle fixture when present, preserving null as the no-hook result.</summary>
        /// <param name="catalog">Validated catalog used for the invocation.</param>
        /// <param name="fixture">Optional fixture descriptor.</param>
        /// <param name="phase">Module or test lifecycle phase label.</param>
        /// <returns>Invocation task, or a completed task containing null when no fixture exists.</returns>
        private Task<VbaTestResult> Fixture(VbaTestCatalog catalog, VbaTestDescriptor fixture, string phase)
        { return fixture == null ? Task.FromResult<VbaTestResult>(null) : Invoke(catalog, fixture, phase); }

        /// <summary>Wraps an operation so its continuation returns to the execution host's owning context when available.</summary>
        /// <typeparam name="T">Task result type.</typeparam>
        /// <param name="task">Operation whose completion must be observed.</param>
        /// <returns>Owner-aware awaitable, or a normal awaitable for hosts without a continuation dispatcher.</returns>
        private VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
        {
            return !(host is IVbaTestContinuationHost owned) ? VbaTestOwnerAwaitable<T>.Unowned(task) : owned.AwaitOwner(task);
        }

        /// <summary>Converts adapter preflight failures into explicit safe-refusal invocation failures.</summary>
        /// <param name="catalog">Catalog to validate before any native procedure call.</param>
        private void Validate(VbaTestCatalog catalog)
        {
            try { host.Validate(catalog); }
            catch (VbaTestInvocationException) { throw; }
            catch (Exception error) { throw new VbaTestInvocationException(error.Message, false, error); }
        }

        /// <summary>Revalidates and invokes one procedure, requiring a matching descriptor and verified terminal verdict.</summary>
        /// <param name="catalog">Source-bound catalog supplied to the execution adapter.</param>
        /// <param name="procedure">Procedure descriptor whose identifier must match the returned receipt.</param>
        /// <param name="phase">Phase name recorded on the result and included in diagnostics.</param>
        /// <returns>Verified result for the requested procedure.</returns>
        /// <exception cref="VbaTestInvocationException">Preflight refuses dispatch, or dispatch completion cannot be verified; uncertainty is preserved.</exception>
        private async Task<VbaTestResult> Invoke(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
        {
            try { Validate(catalog); }
            catch (VbaTestInvocationException error)
            { throw new VbaTestInvocationException(phase + " (" + procedure.Module + "." + procedure.Procedure + "): " + error.Message, error.Uncertain, error); }
            try
            {
                var result = await AwaitOwner(host.InvokeAsync(catalog, procedure, phase));
                if (result == null || result.Test == null || !string.Equals(result.Test.Id, procedure.Id, StringComparison.OrdinalIgnoreCase)
                    || (result.Outcome != VbaTestOutcome.Passed && result.Outcome != VbaTestOutcome.Failed
                        && result.Outcome != VbaTestOutcome.Error && result.Outcome != VbaTestOutcome.Inconclusive))
                    throw new VbaTestInvocationException(result?.Message ?? "The native call did not return verified completion evidence.", true);
                result.Test = procedure;
                result.Phase = phase;
                return result;
            }
            catch (VbaTestInvocationException error)
            { throw new VbaTestInvocationException(phase + " (" + procedure.Module + "." + procedure.Procedure + "): " + error.Message, error.Uncertain, error); }
            catch (Exception error)
            { throw new VbaTestInvocationException(phase + " (" + procedure.Module + "." + procedure.Procedure + "): " + error.Message, true, error); }
        }

        /// <summary>Maps a refusal exception to blocked, cancelled, or outcome-unknown without inventing a verdict.</summary>
        /// <param name="test">Descriptor the runner attempted to schedule.</param>
        /// <param name="error">Failure raised by preflight or invocation.</param>
        /// <param name="phase">Lifecycle phase in which the refusal occurred.</param>
        /// <param name="cancelled">Whether cancellation caused a safe refusal before dispatch.</param>
        /// <returns>Result preserving any possible native execution uncertainty.</returns>
        private static VbaTestResult Refused(VbaTestDescriptor test, Exception error, string phase, bool cancelled = false)
        {
            bool uncertain = error is VbaTestInvocationException invocation && invocation.Uncertain;
            return new VbaTestResult
            {
                Test = test,
                Outcome = uncertain ? VbaTestOutcome.OutcomeUnknown
                : cancelled ? VbaTestOutcome.Cancelled : VbaTestOutcome.Blocked,
                Message = error.Message,
                Phase = phase
            };
        }

        /// <summary>Combines a cleanup failure with the preceding test verdict, promoting uncertain cleanup to unknown.</summary>
        /// <param name="result">Test result to update in place.</param>
        /// <param name="cleanup">Cleanup result whose diagnostic and error number are appended.</param>
        private static void AppendFailure(VbaTestResult result, VbaTestResult cleanup)
        {
            result.Message = (result.Message ?? "") + "\r\n" + cleanup.Phase + ": " + cleanup.Message;
            result.Outcome = cleanup.Outcome == VbaTestOutcome.OutcomeUnknown ? VbaTestOutcome.OutcomeUnknown : VbaTestOutcome.Error;
            result.Phase = cleanup.Phase;
            if (cleanup.ErrorNumber != 0) result.ErrorNumber = cleanup.ErrorNumber;
        }

        /// <summary>Returns the first nonempty diagnostic from an ordered set of failure sources.</summary>
        /// <param name="messages">Candidate messages in precedence order.</param>
        /// <returns>First nonempty message, or an empty string when none are present.</returns>
        private static string FirstMessage(params string[] messages) => messages.FirstOrDefault(message => !string.IsNullOrEmpty(message));

        /// <summary>Appends a run-level diagnostic without discarding earlier failures.</summary>
        /// <param name="run">Run receipt updated in place.</param>
        /// <param name="message">New error text to append.</param>
        private static void AddRunError(VbaTestRun run, string message)
        { run.Error = string.IsNullOrEmpty(run.Error) ? message : run.Error + "\r\n" + message; }

        /// <summary>Records one result, updates run-level uncertainty, and notifies the optional progress callback.</summary>
        /// <param name="run">Run receipt receiving the result.</param>
        /// <param name="result">Terminal result for one selected descriptor.</param>
        /// <param name="progress">Optional callback invoked after the result is stored.</param>
        private static void Publish(VbaTestRun run, VbaTestResult result, Action<VbaTestResult> progress)
        {
            run.Results.Add(result);
            if (result.Outcome == VbaTestOutcome.OutcomeUnknown) run.OutcomeUnknown = true;
            progress?.Invoke(result);
        }
    }
}
