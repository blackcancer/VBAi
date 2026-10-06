using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{

    /// <summary>Owns the llm vbe tools state and operations.</summary>
    internal sealed partial class LlmVbeTools
    {

        /// <summary>Maintains the testing tools state for llm vbe tools.</summary>
        private static readonly HashSet<string> TestingTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "discover_vba_tests", "preview_vba_test_support", "install_vba_test_support", "run_vba_tests",
            "vba_test_run_status", "stop_vba_tests", "navigate_vba_test", "vba_test_coverage", "show_vba_test_explorer"
        };

        /// <summary>Gets the testing definitions.</summary>
        /// <value>Current testing definitions exposed by llm vbe tools.</value>
        private static object[] TestingDefinitions => new[]
        {
            Definition("discover_vba_tests", "Discover explicitly annotated VBA tests and fixtures in one authorized project without executing or editing source. Returns stable test IDs, blocking discovery diagnostics and the whole-project ExpectedProjectVersion. A discovered test is not a passing test. Standard modules use @TestModule and @TestMethod; valid tests are explicit Public parameterless Subs or Functions As Boolean.", new[] { "Project" }, "Project"),
            Definition("preview_vba_test_support", "Preview the exact project-local VBA assertion and guarded result-channel module for the current test catalogue. Returns Text and ExpectedProjectVersion; inspect and review that complete source before install_vba_test_support. This inspection does not install, execute or save anything.", new[] { "Project" }, "Project"),
            Definition("install_vba_test_support", "Install the exact Text reviewed from preview_vba_test_support into this conversation's project with its current ExpectedProjectVersion and ExpectedMode=2. Refuses collisions, stale source and modified preview text. Subject to Agent mode and VBE edit approval. Returns source backup/recovery evidence; no implicit saving, execution or automatic retry after uncertain mutation.", new[] { "Project", "ExpectedProjectVersion", "ExpectedMode", "Text" }, "Project", "ExpectedProjectVersion", "ExpectedMode", "Text"),
            Definition("run_vba_tests", "Run an explicit nonempty selection of up to 10000 distinct test IDs with current ExpectedProjectVersion and ExpectedMode=2. Requires a qualified host, Agent mode, execution approval and shared VBE access. Omit Action for installed support in the original project. Action coverage explicitly creates and retains a separate host document copy under LOCALAPPDATA/VBAi/CoverageRuns, instruments its production procedures and installs support only in that copy; inspect vba_test_coverage beforehand. Original source is not instrumented. Tests execute arbitrary VBA with host privileges, and a copy is not an external-system sandbox. Coverage is procedure-entry coverage, not statement/branch coverage or pass rate. Inspect returned run via vba_test_run_status. Never retry uncertain completion.", new[] { "Project", "ExpectedProjectVersion", "ExpectedMode", "Items" }, "Project", "ExpectedProjectVersion", "ExpectedMode", "Items", "Action"),
            Definition("vba_test_run_status", "Read one exact project-scoped run using its identifier in Query, without executing again. Action compact (default) pages complete assertion messages with global counts, pass rate, coverage summary and uncertainty. Offset is zero-based nonnegative Int32; Limit is 1..100, default 100 (0 also selects default). Follow nextOffset until null to retrieve all tests and coverage details; total counts tests, coverage has separate probe/exclusion/diagnostic totals with the same offset. Preserve historical revision, stale source, skipped/blocked tests and partial runs. Action human selects the readable page with the same global summaries and page metadata. Local exports remain complete. Results may contain private assertion messages. Pass rate and completion rate are distinct from measured procedure coverage.", new[] { "Project", "Query" }, "Project", "Query", "Action", "Offset", "Limit"),
            Definition("stop_vba_tests", "Request cooperative stop for the exact project/run identifier in Query. Scheduling stops after the current VBA call and cleanup finish. Does not terminate VBA, reset VBE, kill the host or prove that a blocked call stopped. Requires Agent mode, execution policy and shared VBE context access.", new[] { "Project", "Query" }, "Project", "Query"),
            Definition("navigate_vba_test", "Navigate to exactly one discovered test ID in Items after validating the current whole-project ExpectedProjectVersion. Opens source location without modifying or executing VBA. Stale identities are refused.", new[] { "Project", "ExpectedProjectVersion", "Items" }, "Project", "ExpectedProjectVersion", "Items"),
            Definition("vba_test_coverage", "Preview eligible production procedures, mapped probes, exclusions, unsupported syntax and clone capability without editing or executing. Optional Query reads measured coverage for that exact run, preserving revision and incomplete/unavailable outcomes. Offset is zero-based nonnegative Int32; Limit is 1..100, default 100 (0 also selects default). Follow NextOffset until null; Probes/Hits, Exclusions and Diagnostics use the same offset, with independent totals and a global Total equal to the longest collection. All eligible/hit/percent summaries remain global. Percent means hit/eligible procedure entries; statements and branches are not measured. Never substitute pass rate or .NET coverage for this metric.", new[] { "Project" }, "Project", "Query", "Offset", "Limit"),
            Definition("show_vba_test_explorer", "Open the test explorer and select one exact authorized project without executing or editing VBA. Returns the owned window handle and docking state. A running batch retains its frozen project scope.", new[] { "Project" }, "Project")
        };

        /// <summary>Determines whether testing tool for llm vbe tools.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is testing tool on llm vbe tools.</returns>
        private static bool IsTestingTool(string name) => TestingTools.Contains(name);

        /// <summary>Handles prepare testing request for llm vbe tools.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        private void PrepareTestingRequest(Request request)
        {
            if (request.Command == "vba_test_run_status" || request.Command == "vba_test_coverage")
                VbaTestReports.PageLimit(request.Offset, request.Limit);
            if (request.Command == "run_vba_tests" || request.Command == "navigate_vba_test")
            {
                if (request.Items == null || request.Items.Length == 0 || request.Items.Any(string.IsNullOrWhiteSpace))
                    throw new ArgumentException("Items must contain explicit nonempty test identifiers.");
                if (request.Items.Distinct(StringComparer.OrdinalIgnoreCase).Count() != request.Items.Length)
                    throw new ArgumentException("Items must contain distinct test identifiers.");
                if (request.Command == "navigate_vba_test" && request.Items.Length != 1)
                    throw new ArgumentException("Navigation requires exactly one test identifier.");
            }
            if ((request.Command == "run_vba_tests" || request.Command == "install_vba_test_support") && request.ExpectedMode != 2)
                throw new ArgumentException("ExpectedMode must be 2 (design mode).");
            if (request.Command == "run_vba_tests" && !string.IsNullOrEmpty(request.Action) && request.Action != "coverage")
                throw new ArgumentException("Action must be coverage or omitted.");
            if (request.Command == "vba_test_run_status" && !string.IsNullOrEmpty(request.Action) && request.Action != "compact" && request.Action != "human")
                throw new ArgumentException("Action must be compact or human.");
            if (request.Command != "install_vba_test_support") return;
            if (string.IsNullOrWhiteSpace(request.Text) || request.Text.Length > 1024 * 1024)
                throw new ArgumentException("Reviewed support Text must be nonempty and at most 1 MiB.");
            // Approval must concern the current, concrete source rather than an unseen generated module.
            Response preview = Execute(new Request { Command = "preview_vba_test_support", Project = request.Project });
            if (!preview.Ok) throw new InvalidOperationException(preview.Error);
            var data = Fields(preview.Data);
            object revision, text;
            if (data == null || !data.TryGetValue("ExpectedProjectVersion", out revision) ||
                !string.Equals(Convert.ToString(revision), request.ExpectedProjectVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since the reviewed test-support preview.");
            if (!data.TryGetValue("Text", out text) || !string.Equals(text as string, request.Text, StringComparison.Ordinal))
                throw new InvalidOperationException("Text must exactly match the reviewed test-support preview.");
        }

        /// <summary>Handles revalidate testing dispatch for llm vbe tools.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="arguments">Text that supplies the arguments value. Use the format required by the calling operation.</param>
        /// <param name="approved">Indicates whether approved is enabled.</param>
        /// <param name="expectedBinding">Text that supplies the expected binding value. Use the format required by the calling operation.</param>
        private void RevalidateTestingDispatch(string name, string arguments, bool approved, string expectedBinding)
        {
            GuardMode(name);
            if (!SameProject(expectedBinding, BoundProject))
                throw new InvalidOperationException("The conversation project binding changed before the test operation.");
            GuardProject(name, arguments);
            GuardLegacyEditorMutation(name);
            if (!ReadOnlyTools.Contains(name) && settings.VbeEditApproval != "Automatic" &&
                !(settings.VbeEditApproval == "AskEachTime" && approved))
                throw new InvalidOperationException("VBE edit policy changed before the test operation.");
        }

        /// <summary>Executes testing request for llm vbe tools.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="arguments">Text that supplies the arguments value. Use the format required by the calling operation.</param>
        /// <param name="approved">Indicates whether approved is enabled.</param>
        /// <param name="expectedBinding">Text that supplies the expected binding value. Use the format required by the calling operation.</param>
        /// <returns>response produced by the operation for execute testing request on llm vbe tools.</returns>
        private Response ExecuteTestingRequest(Request request, string name, string arguments, bool approved, string expectedBinding)
        {
            if (name != "run_vba_tests" || session == null) return Execute(request);
            // The session captures this per-run closure before returning from StartRun.
            // Every deferred native call must recheck the conversation's current authority.
            Action previousGuard = session.TestExecutionGuard;
            session.TestExecutionGuard = () => RevalidateTestingDispatch(name, arguments, approved, expectedBinding);
            try { return Execute(request); }
            finally { session.TestExecutionGuard = previousGuard; }
        }
    }
}
