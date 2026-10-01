using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestRunnerTests
    {
        [TestMethod]
        public async Task RunsSelectedTestsOnceInStableOrderWithRequiredFixtures()
        {
            var catalog = Catalog("Beta", "Alpha");
            var module = catalog.Modules[0];
            AddFixtures(module);
            var host = new RecordingHost();
            var published = new List<VbaTestResult>();
            var run = await new VbaTestRunner(host).RunAsync(catalog, module.Tests, published.Add, CancellationToken.None);
            CollectionAssert.AreEqual(new[]
            {
                "ModuleInitialize:StartModule", "TestInitialize:StartTest", "Test:Alpha", "TestCleanup:EndTest",
                "TestInitialize:StartTest", "Test:Beta", "TestCleanup:EndTest", "ModuleCleanup:EndModule"
            }, host.Calls);
            Assert.AreEqual(2, run.Results.Count);
            Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Passed));
            CollectionAssert.AreEqual(run.Results, published);
            Assert.AreEqual("ProjectOne", run.Project);
            Assert.AreEqual("revision1", run.Revision);
            Assert.IsFalse(run.OutcomeUnknown);
        }

        [TestMethod]
        public async Task SingleSelectionStillInvokesFixturesButDoesNotRunOtherTests()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost();
            var run = await new VbaTestRunner(host).RunAsync(catalog, new[] { catalog.Tests.First() }, null, CancellationToken.None);
            Assert.AreEqual(1, run.Results.Count);
            Assert.AreEqual(5, host.Calls.Count);
            Assert.IsFalse(host.Calls.Contains("Test:Beta"));
        }

        [TestMethod]
        public async Task OrdinaryAssertionFailureCleansUpAndContinuesBatchWithoutRetry()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                    procedure.Procedure == "Alpha" ? VbaTestOutcome.Failed : VbaTestOutcome.Passed, "assertion",
                    procedure.Procedure == "Alpha" ? -2147221504 : 0))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual(VbaTestOutcome.Failed, run.Results[0].Outcome);
            Assert.AreEqual(-2147221504, run.Results[0].ErrorNumber);
            Assert.AreEqual(VbaTestOutcome.Passed, run.Results[1].Outcome);
            Assert.AreEqual(1, host.Calls.Count(call => call == "Test:Alpha"));
            Assert.IsTrue(host.Calls.Contains("ModuleCleanup:EndModule"));
            Assert.IsFalse(run.OutcomeUnknown);
        }

        [TestMethod]
        public async Task ModuleInitializationFailureBlocksBodiesAndRunsModuleCleanup()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                    phase == "ModuleInitialize" ? VbaTestOutcome.Error : VbaTestOutcome.Passed, "setup failed", phase == "ModuleInitialize" ? 42 : 0))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            CollectionAssert.AreEqual(new[] { "ModuleInitialize:StartModule", "ModuleCleanup:EndModule" }, host.Calls);
            Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked && result.ErrorNumber == 42
                && result.Phase == "ModuleInitialize"));
            StringAssert.Contains(run.Error, "setup failed");
            Assert.IsFalse(run.OutcomeUnknown);
        }

        [TestMethod]
        public async Task TestInitializationFailureSkipsBodyAndRunsBothCleanupPhases()
        {
            var catalog = Catalog("Alpha");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                    phase == "TestInitialize" ? VbaTestOutcome.Error : VbaTestOutcome.Passed, "fixture error", phase == "TestInitialize" ? 13 : 0))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.IsFalse(host.Calls.Contains("Test:Alpha"));
            Assert.IsTrue(host.Calls.Contains("TestCleanup:EndTest"));
            Assert.IsTrue(host.Calls.Contains("ModuleCleanup:EndModule"));
            Assert.AreEqual(VbaTestOutcome.Error, run.Results.Single().Outcome);
            Assert.AreEqual("TestInitialize", run.Results.Single().Phase);
            Assert.AreEqual(13, run.Results.Single().ErrorNumber);
        }

        [TestMethod]
        public async Task CleanupFailureRetainsBodyMessageBlocksRemainingModuleAndPreservesErrorNumber()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                    phase == "TestCleanup" ? VbaTestOutcome.Error : phase == "Test" ? VbaTestOutcome.Failed : VbaTestOutcome.Passed,
                    phase == "TestCleanup" ? "cleanup broke" : "initial assertion", phase == "TestCleanup" ? 91 : 0))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            var result = run.Results.Single(item => item.Test.Procedure == "Alpha");
            Assert.AreEqual(VbaTestOutcome.Error, result.Outcome);
            Assert.AreEqual("TestCleanup", result.Phase);
            Assert.AreEqual(91, result.ErrorNumber);
            StringAssert.Contains(result.Message, "initial assertion");
            StringAssert.Contains(result.Message, "cleanup broke");
            Assert.AreEqual(VbaTestOutcome.Blocked, run.Results.Single(item => item.Test.Procedure == "Beta").Outcome);
            Assert.IsFalse(host.Calls.Contains("Test:Beta"));
            Assert.IsTrue(host.Calls.Contains("ModuleCleanup:EndModule"));
        }

        [TestMethod]
        public async Task ModuleCleanupFailureMakesRunErroneousWithoutRewritingVerifiedBodies()
        {
            var catalog = Catalog("Alpha");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                    phase == "ModuleCleanup" ? VbaTestOutcome.Error : VbaTestOutcome.Passed, "module cleanup failed", phase == "ModuleCleanup" ? 53 : 0))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual(VbaTestOutcome.Passed, run.Results.Single().Outcome);
            StringAssert.Contains(run.Error, "Module cleanup failed");
            StringAssert.Contains(run.Error, "53");
        }

        [TestMethod]
        public async Task UnknownModuleInitializationNeverInvokesCleanupOrBody()
        {
            foreach (bool returnsUnknown in new[] { false, true })
            {
                var catalog = Catalog("Alpha", "Beta");
                AddFixtures(catalog.Modules[0]);
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) => returnsUnknown
                        ? Task.FromResult(Result(procedure, phase, VbaTestOutcome.OutcomeUnknown, "No completion"))
                        : Task.FromException<VbaTestResult>(new VbaTestInvocationException("Transport lost", true))
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                CollectionAssert.AreEqual(new[] { "ModuleInitialize:StartModule" }, host.Calls);
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked));
            }
        }

        [TestMethod]
        public async Task UnknownTestInitializationAndBodyNeverInvokeAnotherNativeCall()
        {
            foreach (string unknownPhase in new[] { "TestInitialize", "Test" })
            {
                var catalog = Catalog("Alpha", "Beta");
                AddFixtures(catalog.Modules[0]);
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) => Task.FromResult(Result(procedure, phase,
                        phase == unknownPhase ? VbaTestOutcome.OutcomeUnknown : VbaTestOutcome.Passed, "No completion"))
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                Assert.AreEqual(unknownPhase + ":" + (unknownPhase == "Test" ? "Alpha" : "StartTest"), host.Calls.Last());
                Assert.AreEqual(VbaTestOutcome.OutcomeUnknown, run.Results.First().Outcome);
                Assert.AreEqual(unknownPhase, run.Results.First().Phase);
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.AreEqual(VbaTestOutcome.Blocked, run.Results.Last().Outcome);
            }
        }

        [TestMethod]
        public async Task UnknownCleanupNeverInvokesModuleCleanupOrRetries()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => phase == "TestCleanup"
                    ? Task.FromException<VbaTestResult>(new VbaTestInvocationException("Cleanup completion lost", true))
                    : Task.FromResult(Result(procedure, phase))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual("TestCleanup:EndTest", host.Calls.Last());
            Assert.IsTrue(run.OutcomeUnknown);
            Assert.AreEqual(VbaTestOutcome.OutcomeUnknown, run.Results[0].Outcome);
            Assert.AreEqual(1, host.Calls.Count(call => call == "TestCleanup:EndTest"));
        }

        [TestMethod]
        public async Task UnknownModuleCleanupStopsTheFollowingModuleWithoutRewritingCompletedTests()
        {
            var catalog = Catalog("Alpha");
            AddFixtures(catalog.Modules[0]);
            catalog.Modules.Add(new VbaTestModule { Name = "TestsTwo", Tests = { Procedure("TestsTwo", "Beta") } });
            var host = new RecordingHost
            {
                Handler = (procedure, phase) => phase == "ModuleCleanup"
                    ? Task.FromException<VbaTestResult>(new VbaTestInvocationException("Module cleanup completion lost", true))
                    : Task.FromResult(Result(procedure, phase))
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual("ModuleCleanup:EndModule", host.Calls.Last());
            Assert.IsTrue(run.OutcomeUnknown);
            Assert.AreEqual(VbaTestOutcome.Passed, run.Results[0].Outcome);
            Assert.AreEqual(VbaTestOutcome.Blocked, run.Results[1].Outcome);
            StringAssert.Contains(run.Error, "ModuleCleanup");
        }

        [TestMethod]
        public async Task UnexpectedTransportExceptionAndUnrelatedEvidenceAreConservativelyUnknown()
        {
            foreach (bool unrelated in new[] { false, true })
            {
                var catalog = Catalog("Alpha");
                AddFixtures(catalog.Modules[0]);
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) => unrelated
                        ? Task.FromResult(Result(Procedure("Wrong", "Other"), phase))
                        : Task.FromException<VbaTestResult>(new InvalidOperationException("Unexpected adapter failure"))
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
                Assert.AreEqual(1, host.Calls.Count);
                Assert.IsTrue(run.OutcomeUnknown);
            }
        }

        [TestMethod]
        public async Task ChangedProjectRefusalStopsAllCallsIncludingCleanup()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            bool stale = false;
            var host = new RecordingHost
            {
                Validation = () => { if (stale) throw new InvalidOperationException("Project revision changed"); },
                Handler = (procedure, phase) =>
                {
                    if (phase == "Test") stale = true;
                    return Task.FromResult(Result(procedure, phase));
                }
            };
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual("Test:Alpha", host.Calls.Last());
            Assert.IsFalse(run.OutcomeUnknown);
            Assert.AreEqual(VbaTestOutcome.Error, run.Results[0].Outcome);
            Assert.AreEqual(VbaTestOutcome.Blocked, run.Results[1].Outcome);
            StringAssert.Contains(run.Error, "revision changed");
        }

        [TestMethod]
        public async Task CancellationCompletesCurrentTestAndCleanupThenCancelsQueuedTests()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            using (var cancellation = new CancellationTokenSource())
            {
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) =>
                    {
                        if (phase == "Test") cancellation.Cancel();
                        return Task.FromResult(Result(procedure, phase));
                    }
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, cancellation.Token);
                Assert.AreEqual(VbaTestOutcome.Passed, run.Results[0].Outcome);
                Assert.AreEqual(VbaTestOutcome.Cancelled, run.Results[1].Outcome);
                Assert.IsTrue(host.Calls.Contains("TestCleanup:EndTest"));
                Assert.AreEqual("ModuleCleanup:EndModule", host.Calls.Last());
                Assert.IsFalse(host.Calls.Contains("Test:Beta"));
            }
        }

        [TestMethod]
        public async Task AlreadyCancelledRunDoesNotCallHostAndKeepsEverySelectedTest()
        {
            var catalog = Catalog("Alpha", "Beta");
            var host = new RecordingHost();
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, new CancellationToken(true));
            Assert.AreEqual(0, host.Calls.Count);
            Assert.AreEqual(0, host.Validations);
            Assert.AreEqual(2, run.Results.Count);
            Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Cancelled));
        }

        [TestMethod]
        public async Task CancelledQueuedBodyStillCleansUpACompletedInitializer()
        {
            var catalog = Catalog("Alpha", "Beta");
            AddFixtures(catalog.Modules[0]);
            using (var cancellation = new CancellationTokenSource())
            {
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) =>
                    {
                        if (phase == "TestInitialize") cancellation.Cancel();
                        return phase == "Test"
                            ? Task.FromException<VbaTestResult>(new VbaTestInvocationException("Stopped before dispatch", false))
                            : Task.FromResult(Result(procedure, phase));
                    }
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, cancellation.Token);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Cancelled));
                CollectionAssert.AreEqual(new[] { "ModuleInitialize:StartModule", "TestInitialize:StartTest", "Test:Alpha",
                    "TestCleanup:EndTest", "ModuleCleanup:EndModule" }, host.Calls);
                Assert.IsFalse(run.OutcomeUnknown);
            }
        }

        [TestMethod]
        public async Task CancellationBeforeFirstDispatchDoesNotInvokeUnstartedCleanupFixtures()
        {
            var catalog = Catalog("Alpha");
            catalog.Modules[0].TestCleanup = Procedure("TestsOne", "EndTest");
            catalog.Modules[0].ModuleCleanup = Procedure("TestsOne", "EndModule");
            using (var cancellation = new CancellationTokenSource())
            {
                var host = new RecordingHost
                {
                    Handler = (procedure, phase) =>
                    {
                        cancellation.Cancel();
                        return Task.FromException<VbaTestResult>(new VbaTestInvocationException("Stopped before dispatch", false));
                    }
                };
                var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, cancellation.Token);
                CollectionAssert.AreEqual(new[] { "Test:Alpha" }, host.Calls);
                Assert.AreEqual(VbaTestOutcome.Cancelled, run.Results.Single().Outcome);
                Assert.IsFalse(run.OutcomeUnknown);
            }
        }

        [TestMethod]
        public async Task IgnoredAndUnsupportedTestsAreVisibleWithoutInvocationsOrLostReasons()
        {
            var catalog = Catalog("Ignored", "Unsupported");
            catalog.Modules[0].Diagnostic = "";
            catalog.Modules[0].Tests[0].IgnoreReason = "External service unavailable";
            catalog.Modules[0].Tests[1].Diagnostic = "Private test";
            var host = new RecordingHost();
            var run = await new VbaTestRunner(host).RunAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None);
            Assert.AreEqual(0, host.Calls.Count);
            Assert.AreEqual(VbaTestOutcome.Skipped, run.Results[0].Outcome);
            Assert.AreEqual("External service unavailable", run.Results[0].Message);
            Assert.AreEqual(VbaTestOutcome.Blocked, run.Results[1].Outcome);
            Assert.AreEqual("Private test", run.Results[1].Message);
        }

        [TestMethod]
        public async Task RejectsOverlappingRunsAndFrozenSelectionDoesNotAddNewTests()
        {
            var catalog = Catalog("Alpha", "Beta");
            var pending = new TaskCompletionSource<VbaTestResult>();
            var host = new RecordingHost { Handler = (procedure, phase) => pending.Task };
            var runner = new VbaTestRunner(host);
            var selection = new List<VbaTestDescriptor> { catalog.Tests.First() };
            var running = runner.RunAsync(catalog, selection, null, CancellationToken.None);
            selection.Add(catalog.Tests.Last());
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => runner.RunAsync(catalog, selection, null, CancellationToken.None));
            pending.SetResult(Result(catalog.Tests.First(), "Test"));
            var run = await running;
            Assert.AreEqual(1, run.Results.Count);
            Assert.AreEqual(1, host.Calls.Count);
            var repeated = await runner.RunAsync(catalog, new[] { catalog.Tests.First() }, null, CancellationToken.None);
            Assert.AreEqual(1, repeated.Results.Count);
        }

        [TestMethod]
        public async Task RejectsUnknownAndDuplicateSelectionsBeforeInvokingHost()
        {
            var catalog = Catalog("Alpha");
            var host = new RecordingHost();
            var runner = new VbaTestRunner(host);
            await Assert.ThrowsExceptionAsync<ArgumentException>(() => runner.RunAsync(catalog,
                new[] { Procedure("Other", "Unknown") }, null, CancellationToken.None));
            await Assert.ThrowsExceptionAsync<ArgumentException>(() => runner.RunAsync(catalog,
                new[] { catalog.Tests.First(), catalog.Tests.First() }, null, CancellationToken.None));
            Assert.AreEqual(0, host.Calls.Count);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void OwnedRunnerPreservesFixturePublicationAndCleanupThreadAfterNativeContextLoss(bool nativeFailure)
        {
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(null);
            try
            {
                using (var queue = new BlockingCollection<Action>())
                {
                    var catalog = Catalog("Alpha"); AddFixtures(catalog.Modules[0]);
                    var host = new OwnedAsyncHost(queue) { FailTest = nativeFailure };
                    var runner = new VbaTestRunner(host);
                    var task = runner.RunAsync(catalog, catalog.Tests.ToArray(), _ => host.RequireOwner(), CancellationToken.None);
                    VbaTestOwnerAwaitableTests.Pump(task, queue);
                    var run = task.GetAwaiter().GetResult();
                    Assert.AreEqual(nativeFailure, run.OutcomeUnknown);
                    Assert.AreEqual(nativeFailure ? VbaTestOutcome.OutcomeUnknown : VbaTestOutcome.Passed, run.Results.Single().Outcome);
                    CollectionAssert.AreEqual(nativeFailure
                        ? new[] { "ModuleInitialize", "TestInitialize", "Test" }
                        : new[] { "ModuleInitialize", "TestInitialize", "Test", "TestCleanup", "ModuleCleanup" }, host.Phases);
                    Assert.IsTrue(host.Workers.All(thread => thread != Thread.CurrentThread.ManagedThreadId));
                    Assert.IsNull(SynchronizationContext.Current);
                    // The runner's final active-state reset must also allow a later independently authorized run.
                    host.FailTest = false;
                    task = runner.RunAsync(catalog, catalog.Tests.ToArray(), _ => host.RequireOwner(), CancellationToken.None);
                    VbaTestOwnerAwaitableTests.Pump(task, queue);
                    Assert.AreEqual(VbaTestOutcome.Passed, task.GetAwaiter().GetResult().Results.Single().Outcome);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        private sealed class OwnedAsyncHost : IVbaTestExecutionHost, IVbaTestContinuationHost
        {
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            private readonly BlockingCollection<Action> queue;
            internal bool FailTest;
            internal readonly List<string> Phases = new List<string>();
            internal readonly ConcurrentBag<int> Workers = new ConcurrentBag<int>();
            internal OwnedAsyncHost(BlockingCollection<Action> queue) { this.queue = queue; }
            internal void RequireOwner() => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, "Native state, publication and cleanup must remain on the owner.");
            public void Validate(VbaTestCatalog catalog) { RequireOwner(); }
            public VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task) => new VbaTestOwnerAwaitable<T>(task, queue.Add, RequireOwner);
            public Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
            {
                RequireOwner(); Phases.Add(phase);
                SynchronizationContext.SetSynchronizationContext(null);
                var completion = new TaskCompletionSource<VbaTestResult>(TaskCreationOptions.RunContinuationsAsynchronously);
                // The pump reaches this action only after InvokeAsync has returned an incomplete task.
                queue.Add(() => Task.Run(() =>
                {
                    Workers.Add(Thread.CurrentThread.ManagedThreadId);
                    if (FailTest && phase == "Test") completion.SetException(new VbaTestInvocationException("Native completion became uncertain.", true));
                    else completion.SetResult(Result(procedure, phase));
                }).GetAwaiter().GetResult());
                return completion.Task;
            }
        }

        private static VbaTestCatalog Catalog(params string[] names)
        {
            var catalog = new VbaTestCatalog { Project = new VbaTestProjectSnapshot { Name = "ProjectOne", Revision = "revision1" } };
            catalog.Modules.Add(new VbaTestModule { Name = "TestsOne", Tests = names.Select(name => Procedure("TestsOne", name)).ToList() });
            return catalog;
        }

        private static VbaTestDescriptor Procedure(string module, string name) =>
            new VbaTestDescriptor { Id = module + "." + name, Module = module, Procedure = name, Kind = "Sub" };

        private static void AddFixtures(VbaTestModule module)
        {
            module.ModuleInitialize = Procedure(module.Name, "StartModule");
            module.TestInitialize = Procedure(module.Name, "StartTest");
            module.TestCleanup = Procedure(module.Name, "EndTest");
            module.ModuleCleanup = Procedure(module.Name, "EndModule");
        }

        private static VbaTestResult Result(VbaTestDescriptor procedure, string phase, VbaTestOutcome outcome = VbaTestOutcome.Passed,
            string message = "", int error = 0) =>
            new VbaTestResult { Test = procedure, Phase = phase, Outcome = outcome, Message = message, ErrorNumber = error };

        private sealed class RecordingHost : IVbaTestExecutionHost
        {
            internal readonly List<string> Calls = new List<string>();
            internal int Validations;
            internal Action Validation;
            internal Func<VbaTestDescriptor, string, Task<VbaTestResult>> Handler;
            public void Validate(VbaTestCatalog catalog) { Validations++; Validation?.Invoke(); }
            public Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
            {
                Calls.Add(phase + ":" + procedure.Procedure);
                return Handler == null ? Task.FromResult(Result(procedure, phase)) : Handler(procedure, phase);
            }
        }
        [TestMethod]
        public async Task EverySelectionPreflightRejectsBeforeCallingTheHost()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new VbaTestRunner(null));
            var host = new RecordingHost(); var runner = new VbaTestRunner(host); var catalog = Catalog("Alpha");
            await Assert.ThrowsExceptionAsync<ArgumentException>(() => runner.RunAsync(null, catalog.Modules[0].Tests, null, CancellationToken.None));
            await Assert.ThrowsExceptionAsync<ArgumentException>(() => runner.RunAsync(new VbaTestCatalog(), catalog.Modules[0].Tests, null, CancellationToken.None));
            foreach (var selection in new IReadOnlyList<VbaTestDescriptor>[] { null, new VbaTestDescriptor[0], new VbaTestDescriptor[10001], new VbaTestDescriptor[] { null }, new[] { Procedure("Foreign", "Test") } })
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => runner.RunAsync(catalog, selection, null, CancellationToken.None));
            Assert.AreEqual(0, host.Validations); Assert.AreEqual(0, host.Calls.Count);
        }

        [TestMethod]
        public async Task InFlightRunnerRejectsConcurrentUseAndReleasesAfterVerifiedCompletion()
        {
            var completion = new TaskCompletionSource<VbaTestResult>(); var catalog = Catalog("Alpha");
            var host = new RecordingHost { Handler = (procedure,phase) => completion.Task }; var runner = new VbaTestRunner(host);
            var running = runner.RunAsync(catalog, catalog.Modules[0].Tests, null, CancellationToken.None);
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => runner.RunAsync(catalog, catalog.Modules[0].Tests, null, CancellationToken.None));
            completion.SetResult(Result(catalog.Modules[0].Tests[0], "Test"));
            Assert.AreEqual(VbaTestOutcome.Passed, (await running).Results.Single().Outcome);
            Assert.AreEqual(1, host.Calls.Count);
        }

        [TestMethod]
        public async Task MalformedCompletionEvidenceIsUncertainEvenWithoutMessageOrDescriptor()
        {
            foreach (var invalid in new VbaTestResult[] { null, new VbaTestResult(), new VbaTestResult { Message = "reason" } })
            {
                var catalog = Catalog("Alpha"); var host = new RecordingHost { Handler = (procedure,phase) => Task.FromResult(invalid) };
                var run = await new VbaTestRunner(host).RunAsync(catalog,catalog.Modules[0].Tests,null,CancellationToken.None);
                Assert.IsTrue(run.OutcomeUnknown); Assert.AreEqual(1,host.Calls.Count);
                StringAssert.Contains(run.Error, invalid?.Message ?? "verified completion evidence");
            }
        }

        [TestMethod]
        public async Task NativeValidationRefusalIsPreservedAndCleanupErrorWithoutNumberAppendsToSetupError()
        {
            var catalog = Catalog("Alpha");
            var refused = new RecordingHost { Validation = () => { throw new VbaTestInvocationException("native preflight", false); } };
            var run = await new VbaTestRunner(refused).RunAsync(catalog,catalog.Modules[0].Tests,null,CancellationToken.None);
            Assert.IsFalse(run.OutcomeUnknown); Assert.AreEqual(0,refused.Calls.Count); StringAssert.Contains(run.Error,"native preflight");
            AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost { Handler = (procedure,phase) => Task.FromResult(Result(procedure,phase,
                phase == "ModuleInitialize" || phase == "ModuleCleanup" ? VbaTestOutcome.Error : VbaTestOutcome.Passed,"fixture refused")) };
            run = await new VbaTestRunner(host).RunAsync(catalog,catalog.Modules[0].Tests,null,CancellationToken.None);
            StringAssert.Contains(run.Error,"Module initialization failed"); StringAssert.Contains(run.Error,"Module cleanup failed");
            Assert.IsFalse(run.Error.Contains("VBA error"));
        }

        [TestMethod]
        public async Task SuccessfulInconclusiveEvidenceAndEmptyBodyMessageSurviveCleanupFailure()
        {
            var catalog = Catalog("Alpha"); AddFixtures(catalog.Modules[0]);
            var host = new RecordingHost { Handler = (procedure,phase) => Task.FromResult(Result(procedure,phase,
                phase == "Test" ? VbaTestOutcome.Inconclusive : phase == "TestCleanup" ? VbaTestOutcome.Error : VbaTestOutcome.Passed,
                phase == "Test" ? null : "cleanup reason")) };
            var run = await new VbaTestRunner(host).RunAsync(catalog,catalog.Modules[0].Tests,null,CancellationToken.None);
            Assert.AreEqual(VbaTestOutcome.Error, run.Results.Single().Outcome); Assert.AreEqual(0,run.Results.Single().ErrorNumber);
            StringAssert.Contains(run.Results.Single().Message,"TestCleanup: cleanup reason");
        }
        [TestMethod]
        public async Task ProgressFailureRemainsExplicitWithoutChangingTheVerifiedTestOutcome()
        {
            var catalog=Catalog("Alpha");var host=new RecordingHost();
            var run=await new VbaTestRunner(host).RunAsync(catalog,catalog.Modules[0].Tests,_ => {throw new InvalidOperationException("publication failed");},CancellationToken.None);
            Assert.AreEqual(VbaTestOutcome.Passed,run.Results.Single().Outcome); Assert.IsFalse(run.OutcomeUnknown);
            StringAssert.Contains(run.Error,"publication failed"); Assert.AreEqual(1,host.Calls.Count);
        }

        [TestMethod]
        public async Task OwnerContinuationFailureStopsSchedulingAndNeverInvokesCleanup()
        {
            var catalog=Catalog("Alpha");var host=new RefusingContinuationHost();
            var run=await new VbaTestRunner(host).RunAsync(catalog,catalog.Modules[0].Tests,null,CancellationToken.None);
            Assert.AreEqual(VbaTestOutcome.Blocked,run.Results.Single().Outcome); Assert.IsFalse(run.OutcomeUnknown);
            StringAssert.Contains(run.Error,"owner unavailable"); Assert.AreEqual(1,host.Calls);
        }

        private sealed class RefusingContinuationHost : IVbaTestExecutionHost, IVbaTestContinuationHost
        {
            private int awaits;
            internal int Calls;
            public void Validate(VbaTestCatalog catalog) { }
            public Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog,VbaTestDescriptor procedure,string phase) { Calls++;return Task.FromResult(Result(procedure,phase)); }
            public VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
            { if(++awaits==4) throw new InvalidOperationException("owner unavailable");return VbaTestOwnerAwaitable<T>.Unowned(task); }
        }
    }
}
