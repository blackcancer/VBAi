using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{

    /// <summary>Captures the identity, revision, references, and module source for one discovered VBA project.</summary>
    internal sealed class VbaTestProjectSnapshot
    {

        /// <summary>Gets or sets the stable project identity used by test discovery and execution.</summary>
        /// <value>Project identifier; may be null when the host did not provide one.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the selector that can reopen or resolve this project in its host.</summary>
        /// <value>Host selector, which may be null when only the display name is available.</value>
        public string Selector { get; set; }

        /// <summary>Gets or sets the project name captured for display and fallback identification.</summary>
        /// <value>Captured project name.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the host document path associated with the project snapshot.</summary>
        /// <value>Host path, or null when the project has no file-backed document.</value>
        public string HostPath { get; set; }

        /// <summary>Gets or sets the revision token against which test discovery and coverage are bound.</summary>
        /// <value>Captured revision token; consumers compare it before applying results.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets the fingerprint of project references used to detect a changed compile context.</summary>
        /// <value>Reference fingerprint captured with the snapshot.</value>
        public string ReferencesHash { get; set; }

        /// <summary>Gets or sets the source snapshots for the project's VBA components.</summary>
        /// <value>Module snapshots; initialized to an empty array and may be replaced by the producer.</value>
        public VbaTestModuleSnapshot[] Modules { get; set; } = new VbaTestModuleSnapshot[0];
    }

    /// <summary>Stores one component's source and identity as observed during test discovery.</summary>
    internal sealed class VbaTestModuleSnapshot
    {

        /// <summary>Gets or sets the component name used to map discovered descriptors back to source.</summary>
        /// <value>Component name.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the complete component source captured at the project revision.</summary>
        /// <value>VBA source text, or null if the host could not read it.</value>
        public string Source { get; set; }

        /// <summary>Gets or sets the source fingerprint used to detect stale component contents.</summary>
        /// <value>Fingerprint for the captured source.</value>
        public string Hash { get; set; }

        /// <summary>Gets or sets the VBIDE component type value reported by the host.</summary>
        /// <value>Numeric component type; interpretation follows the VBIDE component-type enumeration.</value>
        public int ComponentType { get; set; }
    }

    /// <summary>Identifies a discovered test or lifecycle procedure and its source location.</summary>
    internal sealed class VbaTestDescriptor
    {

        /// <summary>Gets or sets the stable descriptor key used for selection and result correlation.</summary>
        /// <value>Descriptor identifier.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the containing component name.</summary>
        /// <value>VBA module or class name.</value>
        public string Module { get; set; }

        /// <summary>Gets or sets the procedure name to invoke for this descriptor.</summary>
        /// <value>Procedure name as discovered from source.</value>
        public string Procedure { get; set; }

        /// <summary>Gets or sets the descriptor kind, distinguishing test cases from setup and cleanup hooks.</summary>
        /// <value>Discovery kind consumed by the runner.</value>
        public string Kind { get; set; }

        /// <summary>Gets or sets the one-based source line where the procedure declaration was found.</summary>
        /// <value>One-based line number in the captured module source.</value>
        public int Line { get; set; }

        /// <summary>Gets or sets category labels parsed from the test marker for filtering and display.</summary>
        /// <value>Category names; initialized to an empty array.</value>
        public string[] Categories { get; set; } = new string[0];

        /// <summary>Gets or sets the marker-provided reason when discovery excludes this test.</summary>
        /// <value>Ignore explanation, or null when the test is not explicitly ignored.</value>
        public string IgnoreReason { get; set; }

        /// <summary>Gets or sets a discovery diagnostic that prevents or qualifies execution.</summary>
        /// <value>Diagnostic text, or null when discovery found no descriptor-specific issue.</value>
        public string Diagnostic { get; set; }
    }

    /// <summary>Groups discovered test descriptors and module-level lifecycle hooks for one component.</summary>
    internal sealed class VbaTestModule
    {

        /// <summary>Gets or sets the component name represented by this discovery group.</summary>
        /// <value>Component name.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets a component-level discovery diagnostic.</summary>
        /// <value>Diagnostic text, or null when the component was inspected without a reported issue.</value>
        public string Diagnostic { get; set; }

        /// <summary>Gets or sets the runnable test descriptors discovered in this component.</summary>
        /// <value>Test descriptors; initialized to an empty list.</value>
        public List<VbaTestDescriptor> Tests { get; set; } = new List<VbaTestDescriptor>();

        /// <summary>Gets or sets the optional module initialization hook that runs before this module's tests.</summary>
        /// <value>Initialization descriptor, or null when no hook was discovered.</value>
        public VbaTestDescriptor ModuleInitialize { get; set; }

        /// <summary>Gets or sets the optional module cleanup hook run after this module's test group.</summary>
        /// <value>Cleanup descriptor, or null when no hook was discovered.</value>
        public VbaTestDescriptor ModuleCleanup { get; set; }

        /// <summary>Gets or sets the optional per-test setup hook.</summary>
        /// <value>Setup descriptor, or null when no hook was discovered.</value>
        public VbaTestDescriptor TestInitialize { get; set; }

        /// <summary>Gets or sets the optional per-test cleanup hook.</summary>
        /// <value>Cleanup descriptor, or null when no hook was discovered.</value>
        public VbaTestDescriptor TestCleanup { get; set; }
    }

    /// <summary>Combines a project snapshot, per-component discovery results, and catalog diagnostics.</summary>
    internal sealed class VbaTestCatalog
    {

        /// <summary>Gets or sets the project snapshot that bounds this catalog.</summary>
        /// <value>Snapshot used for discovery, or null before a project is available.</value>
        public VbaTestProjectSnapshot Project { get; set; }

        /// <summary>Gets or sets component groups containing tests and lifecycle hooks.</summary>
        /// <value>Discovered module groups; initialized to an empty list.</value>
        public List<VbaTestModule> Modules { get; set; } = new List<VbaTestModule>();

        /// <summary>Gets or sets diagnostics that apply to catalog discovery as a whole.</summary>
        /// <value>Diagnostic messages; initialized to an empty list.</value>
        public List<string> Diagnostics { get; set; } = new List<string>();

        /// <summary>Gets all test descriptors flattened from the discovered module groups.</summary>
        /// <value>A deferred sequence over the current <see cref="Modules"/> collection.</value>
        public IEnumerable<VbaTestDescriptor> Tests => Modules.SelectMany(module => module.Tests);
    }

    /// <summary>Records whether a discovered VBA test ran and the terminal state reported by the host.</summary>
    internal enum VbaTestOutcome
    {

        /// <summary>No execution result was recorded for the test.</summary>
        NotRun,

        /// <summary>The test completed and reported success.</summary>
        Passed,

        /// <summary>The test completed and reported an assertion or expected failure.</summary>
        Failed,

        /// <summary>Execution raised an error before a normal test verdict was produced.</summary>
        Error,

        /// <summary>The test ran but did not produce a pass or fail verdict.</summary>
        Inconclusive,

        /// <summary>The runner deliberately omitted this test.</summary>
        Skipped,

        /// <summary>The runner observed cancellation and stopped the test sequence.</summary>
        Cancelled,

        /// <summary>Execution was refused because a required precondition was not satisfied.</summary>
        Blocked,

        /// <summary>The host call ended without reliable proof whether the test executed; do not retry automatically.</summary>
        OutcomeUnknown
    }

    /// <summary>Stores one descriptor's terminal verdict and execution diagnostics.</summary>
    internal sealed class VbaTestResult
    {

        /// <summary>Gets or sets the descriptor whose execution this result describes.</summary>
        /// <value>Executed or attempted descriptor.</value>
        public VbaTestDescriptor Test { get; set; }

        /// <summary>Gets or sets the runner's verdict, including blocked and uncertain outcomes.</summary>
        /// <value>Terminal outcome classification.</value>
        public VbaTestOutcome Outcome { get; set; }

        /// <summary>Gets or sets the diagnostic message associated with the verdict.</summary>
        /// <value>Message text, or null when no additional message was captured.</value>
        public string Message { get; set; }

        /// <summary>Gets or sets the setup, test, or cleanup phase active when the result was produced.</summary>
        /// <value>Phase label used to locate lifecycle failures.</value>
        public string Phase { get; set; }

        /// <summary>Gets or sets the host error number captured for an execution error.</summary>
        /// <value>Host error number, or zero when no numeric error was captured.</value>
        public int ErrorNumber { get; set; }

        /// <summary>Gets or sets the elapsed time measured for this result.</summary>
        /// <value>Duration in <see cref="TimeSpan"/> units.</value>
        public TimeSpan Duration { get; set; }
    }

    /// <summary>Aggregates results and coverage for one execution request against a fixed project revision.</summary>
    internal sealed class VbaTestRun
    {

        /// <summary>Gets or sets the identifier used to correlate this run with progress and final receipts.</summary>
        /// <value>Run correlation identifier.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the project identity targeted by this run.</summary>
        /// <value>Project identifier or selector captured for execution.</value>
        public string Project { get; set; }

        /// <summary>Gets or sets the project revision against which discovery and execution were authorized.</summary>
        /// <value>Revision token reported with the run.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets per-descriptor execution results collected before the run terminates.</summary>
        /// <value>Results in runner order; initialized to an empty list.</value>
        public List<VbaTestResult> Results { get; set; } = new List<VbaTestResult>();

        /// <summary>Gets or sets whether a host operation may have executed without a confirmed terminal receipt.</summary>
        /// <value><see langword="true"/> means callers must preserve uncertainty and avoid automatic replay.</value>
        public bool OutcomeUnknown { get; set; }

        /// <summary>Gets or sets a run-level failure or refusal message.</summary>
        /// <value>Error text, or null when no run-level diagnostic was recorded.</value>
        public string Error { get; set; }

        /// <summary>Gets or sets the coverage report produced from the disposable instrumented project.</summary>
        /// <value>Coverage report, or null when coverage was not requested or unavailable.</value>
        public VbaCoverageReport Coverage { get; set; }
    }
}
