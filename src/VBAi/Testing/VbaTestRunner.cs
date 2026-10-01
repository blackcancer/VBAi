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
        void Validate(VbaTestCatalog catalog);
        Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase);
    }

    /// <summary>A refused preflight is distinct from a possibly executed native call.</summary>
    internal sealed class VbaTestInvocationException : Exception
    {
        internal bool Uncertain { get; }
        internal VbaTestInvocationException(string message, bool uncertain, Exception inner = null) : base(message, inner)
        { Uncertain = uncertain; }
    }

    /// <summary>Serial test and fixture scheduling with cooperative cancellation.</summary>
    internal sealed class VbaTestRunner
    {
        private readonly IVbaTestExecutionHost host;
        private bool active;
        internal VbaTestRunner(IVbaTestExecutionHost host) { this.host = host ?? throw new ArgumentNullException(nameof(host)); }

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
                        Publish(run, new VbaTestResult { Test = test, Outcome = !string.IsNullOrEmpty(module.Diagnostic) || !string.IsNullOrEmpty(test.Diagnostic)
                            ? VbaTestOutcome.Blocked : VbaTestOutcome.Skipped, Message = FirstMessage(module.Diagnostic, test.Diagnostic, test.IgnoreReason), Phase = "Discovery" }, progress);
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
                            foreach (var test in runnable) Publish(run, new VbaTestResult { Test = test, Outcome = VbaTestOutcome.Blocked,
                                Message = setup.Message, ErrorNumber = setup.ErrorNumber, Phase = setup.Phase }, progress);
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
                                    result = new VbaTestResult { Test = test, Outcome = VbaTestOutcome.Error, Message = initialized.Message,
                                        ErrorNumber = initialized.ErrorNumber, Phase = initialized.Phase };
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
                            Publish(run, new VbaTestResult { Test = test, Outcome = VbaTestOutcome.Blocked, Phase = "TestCleanup",
                                Message = "A preceding test cleanup failed." }, progress);
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
                        Publish(run, new VbaTestResult { Test = test,
                            Outcome = cancellation.IsCancellationRequested ? VbaTestOutcome.Cancelled : VbaTestOutcome.Blocked,
                            Message = run.Error ?? "Stopped before execution.", Phase = "Scheduling" }, progress);
                }
                finally { active = false; }
            }
            return run;
        }

        private Task<VbaTestResult> Fixture(VbaTestCatalog catalog, VbaTestDescriptor fixture, string phase)
        { return fixture == null ? Task.FromResult<VbaTestResult>(null) : Invoke(catalog, fixture, phase); }

        private VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
        {
            var owned = host as IVbaTestContinuationHost;
            return owned == null ? VbaTestOwnerAwaitable<T>.Unowned(task) : owned.AwaitOwner(task);
        }

        private void Validate(VbaTestCatalog catalog)
        {
            try { host.Validate(catalog); }
            catch (VbaTestInvocationException) { throw; }
            catch (Exception error) { throw new VbaTestInvocationException(error.Message, false, error); }
        }

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

        private static VbaTestResult Refused(VbaTestDescriptor test, Exception error, string phase, bool cancelled = false)
        {
            bool uncertain = error is VbaTestInvocationException invocation && invocation.Uncertain;
            return new VbaTestResult { Test = test, Outcome = uncertain ? VbaTestOutcome.OutcomeUnknown
                : cancelled ? VbaTestOutcome.Cancelled : VbaTestOutcome.Blocked,
                Message = error.Message, Phase = phase };
        }

        private static void AppendFailure(VbaTestResult result, VbaTestResult cleanup)
        {
            result.Message = (result.Message ?? "") + "\r\n" + cleanup.Phase + ": " + cleanup.Message;
            result.Outcome = cleanup.Outcome == VbaTestOutcome.OutcomeUnknown ? VbaTestOutcome.OutcomeUnknown : VbaTestOutcome.Error;
            result.Phase = cleanup.Phase;
            if (cleanup.ErrorNumber != 0) result.ErrorNumber = cleanup.ErrorNumber;
        }

        private static string FirstMessage(params string[] messages) => messages.FirstOrDefault(message => !string.IsNullOrEmpty(message));

        private static void AddRunError(VbaTestRun run, string message)
        { run.Error = string.IsNullOrEmpty(run.Error) ? message : run.Error + "\r\n" + message; }

        private static void Publish(VbaTestRun run, VbaTestResult result, Action<VbaTestResult> progress)
        {
            run.Results.Add(result);
            if (result.Outcome == VbaTestOutcome.OutcomeUnknown) run.OutcomeUnknown = true;
            progress?.Invoke(result);
        }
    }
}
