using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Project-scoped VBA test discovery, explicit execution and result inspection.</summary>
    internal sealed partial class TestExplorerWindow : Form
    {

        /// <summary>Project-scoped discovery, execution, coverage, and source-navigation adapter.</summary>
        private IVbaTestExplorerService service;

        /// <summary>Designer-created size retained for restoring a collapsed native VBE tool window.</summary>
        private readonly System.Drawing.Size initialNativeSize;

        /// <summary>Latest discovered catalog for the displayed project; cleared before each new discovery attempt.</summary>
        private VbaTestCatalog catalog;

        /// <summary>Most recently accepted run result keyed by exact test descriptor ID.</summary>
        private readonly Dictionary<string, VbaTestResult> results = new Dictionary<string, VbaTestResult>(StringComparer.Ordinal);

        /// <summary>Catalog revision that produced each stored result, used to mark stale results.</summary>
        private readonly Dictionary<string, string> resultRevisions = new Dictionary<string, string>(StringComparer.Ordinal);

        /// <summary>Cancellation source for the current explicit test or coverage run.</summary>
        private CancellationTokenSource cancellation;

        /// <summary>Mutually coordinated UI guards for execution, tree rebuild, project reads, deferred close, and coverage review.</summary>
        private bool running, rebuilding, readingProjects, closeAfterRun, reviewingCoverage;

        /// <summary>Current host-specific reason that prevents selected tests from executing.</summary>
        private string unavailableReason;

        /// <summary>Exact ID of the project currently displayed; changing it clears prior reports and coverage.</summary>
        private string displayedProjectId;

        /// <summary>Revision used to determine whether the displayed catalog or accepted results are stale.</summary>
        private string displayedRevision;

        /// <summary>Host-specific reason that prevents measured procedure coverage from running.</summary>
        private string coverageUnavailableReason;

        /// <summary>Last accepted coverage report, displayed only while its project revision remains current.</summary>
        private VbaCoverageReport lastCoverage;

        /// <summary>Stable project/module/test/group identity retained while the tree is rebuilt.</summary>
        private TreeSelection treeSelection;

        /// <summary>Save-dialog route for human-readable or compact JSON report export.</summary>
        internal Func<IWin32Window, bool, string> ChooseReportExportPath = (owner, compact) => SelectReportExportPath(owner, compact);

        /// <summary>Clipboard route used when the user copies the selected report representation.</summary>
        internal Action<string> WriteReportClipboard = Clipboard.SetText;

        /// <summary>Creates the temporary WinForms dispatcher used to resume a run on the owning UI thread.</summary>
        internal Func<Control> CreateRunDispatcher = () => new Control();

        /// <summary>Explicit user-consent prompt shown before measured coverage execution; defaults to No.</summary>
        internal Func<IWin32Window, string, bool> ConfirmCoverage = (owner, text) => MessageBox.Show(owner, text,
            UiText.Get("Review measured VBA procedure coverage"), MessageBoxButtons.YesNo,
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes;

        /// <summary>Managed thread ID captured at construction and required for UI/run continuation operations.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Serializes creation and release of the run continuation dispatcher.</summary>
        private readonly object runDispatchGate = new object();

        /// <summary>Temporary control whose handle marshals run continuations onto the form's owner thread.</summary>
        private Control runContinuationDispatcher;

        /// <summary>Creates controls without opening a project or executing a test.</summary>
        public TestExplorerWindow()
        {
            InitializeComponent();
            initialNativeSize = Size;
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
            // UiText.Apply also attaches the shared theme for the form lifetime.
            outcomeFilter.Items.Add(UiText.Get("All outcomes"));
            foreach (VbaTestOutcome outcome in Enum.GetValues(typeof(VbaTestOutcome))) outcomeFilter.Items.Add(UiText.Get(outcome.ToString()));
            outcomeFilter.SelectedIndex = 0;
            grouping.Items.AddRange(new object[] { UiText.Get("Project / module"), UiText.Get("Outcome"), UiText.Get("Category") });
            grouping.SelectedIndex = 0;
            if (!UiTheme.IsDesignPreview(this))
            {
                UiTheme.Changed += TestThemeChanged;
                Disposed += (sender, args) => UiTheme.Changed -= TestThemeChanged;
            }
            Disposed += RunWindowDisposed;
            UpdateButtons();
        }

        /// <summary>Repairs only an unusable native site while preserving a readable restored VBE layout.</summary>
        /// <param name="container">Hosted chat control used to inspect the native site client size.</param>
        /// <param name="nativeWindow">VBE tool window whose native size/position may be repaired when its docked site collapsed.</param>
        /// <param name="owner">Native owner used to select the current monitor's working area.</param>
        internal void EnsureUsableNativePlacement(ChatToolWindow container, object nativeWindow, IWin32Window owner)
        {
            System.Drawing.Size siteSize;
            if (container == null || !container.TryGetNativeSiteSize(out siteSize) ||
                (siteSize.Width >= MinimumSize.Width && siteSize.Height >= MinimumSize.Height)) return;

            dynamic window = nativeWindow;
            dynamic frame = window.LinkedWindowFrame;
            // Native frame bounds can be large even when its docked client site has collapsed.
            if (frame != null) frame.LinkedWindows.Remove(window);
            var area = Screen.FromHandle(owner.Handle).WorkingArea;
            int width = Math.Min(initialNativeSize.Width, area.Width);
            int height = Math.Min(initialNativeSize.Height, area.Height);
            window.Width = width;
            window.Height = height;
            window.Left = area.Left + Math.Max(0, (area.Width - width) / 2);
            window.Top = area.Top + Math.Max(0, (area.Height - height) / 2);
        }

        /// <summary>Attaches the session service and immediately refreshes the available project catalog.</summary>
        /// <param name="explorerService">Service bound to the current host/VBE session.</param>
        internal void Configure(IVbaTestExplorerService explorerService)
        {
            if (running) throw new InvalidOperationException("A test run is active.");
            service = explorerService ?? throw new ArgumentNullException(nameof(explorerService));
            RefreshProjects();
            UpdateFreshnessTimer();
        }

        /// <summary>Reads accessible projects, preserves the selected project ID when possible, and rediscovers its catalog.</summary>
        internal void RefreshProjects()
        {
            if (service == null || running || reviewingCoverage) return;
            try
            {
                string selectedId = (projectList.SelectedItem as VbaTestProjectSnapshot)?.Id;
                var projects = service.ReadProjects();
                readingProjects = true;
                try
                {
                    projectList.Items.Clear();
                    foreach (var project in projects) projectList.Items.Add(project);
                    int selected = projects.ToList().FindIndex(project => project.Id == selectedId);
                    projectList.SelectedIndex = selected >= 0 ? selected : projects.Count > 0 ? 0 : -1;
                }
                finally { readingProjects = false; }
                DiscoverSelectedProject();
            }
            catch (Exception ex) { catalog = null; status.Text = ex.Message; RebuildTree(); }
        }

        /// <summary>Selects an exact currently available project when no run or coverage review is active.</summary>
        /// <param name="projectId">Stable project ID returned by the attached explorer service.</param>
        internal void SelectProject(string projectId)
        {
            if (string.IsNullOrEmpty(projectId)) throw new ArgumentException("An exact test project ID is required.", nameof(projectId));
            if (running || reviewingCoverage)
            {
                if (catalog?.Project.Id == projectId) return;
                throw new InvalidOperationException("The test explorer is busy with another project.");
            }
            RefreshProjects();
            int index = projectList.Items.Cast<VbaTestProjectSnapshot>().ToList().FindIndex(project => project.Id == projectId);
            if (index < 0) throw new InvalidOperationException("The selected test project is no longer available.");
            projectList.SelectedIndex = index;
        }

        /// <summary>Clears prior executable state, discovers the selected project's current catalog, then rebuilds its tree.</summary>
        private void DiscoverSelectedProject()
        {
            if (running || readingProjects || service == null) return;
            var project = projectList.SelectedItem as VbaTestProjectSnapshot;
            if (displayedProjectId != project?.Id) { humanReport.Clear(); compactReport.Clear(); lastCoverage = null; }
            displayedProjectId = project?.Id;
            // Never keep an earlier project executable after discovery of a new selection fails.
            catalog = null;
            unavailableReason = null;
            coverageUnavailableReason = null;
            try
            {
                if (project != null)
                {
                    catalog = service.Discover(project.Id);
                    unavailableReason = service.ExecutionUnavailableReason(catalog);
                    coverageUnavailableReason = ReadCoverageUnavailableReason(catalog);
                    displayedRevision = catalog.Project.Revision;
                }
                status.Text = catalog == null ? UiText.Get("No accessible VBA project.") :
                    string.Join(Environment.NewLine, catalog.Diagnostics.Concat(string.IsNullOrEmpty(unavailableReason) ? new string[0] : new[] { unavailableReason }).Select(UiText.Get));
            }
            catch (Exception ex) { catalog = null; unavailableReason = ex.Message; coverageUnavailableReason = ex.Message; status.Text = ex.Message; }
            RebuildTree();
        }

        /// <summary>Reads the optional measured-coverage capability reason from the attached host adapter.</summary>
        /// <param name="discovered">Catalog whose project and revision are being considered.</param>
        /// <returns>Host-provided unavailability reason, or the localized unsupported-host explanation.</returns>
        private string ReadCoverageUnavailableReason(VbaTestCatalog discovered) =>
            service is IVbaTestCoverageExplorerService coverageService ? coverageService.CoverageUnavailableReason(discovered) : UiText.Get("Measured procedure coverage is unavailable in this host.");

        /// <summary>Runs freshness polling only while the form is visible, configured, and outside Designer preview.</summary>
        private void UpdateFreshnessTimer() => freshnessTimer.Enabled = Visible && service != null && !UiTheme.IsDesignPreview(this);

        /// <summary>Rediscovers the selected project and marks previous results stale when its revision changes; never starts tests automatically.</summary>
        internal void CheckFreshness()
        {
            if (!Visible || running || reviewingCoverage || readingProjects || service == null || UiTheme.IsDesignPreview(this)) return;
            var project = projectList.SelectedItem as VbaTestProjectSnapshot;
            if (project == null) return;
            try
            {
                var live = service.Discover(project.Id);
                string reason = service.ExecutionUnavailableReason(live);
                string coverageReason = ReadCoverageUnavailableReason(live);
                bool changed = catalog == null || displayedRevision != live.Project.Revision;
                bool capabilityChanged = unavailableReason != reason || coverageUnavailableReason != coverageReason;
                catalog = live;
                unavailableReason = reason;
                coverageUnavailableReason = coverageReason;
                displayedRevision = live.Project.Revision;
                if (changed)
                {
                    status.Text = UiText.Get("Source changed. Previous results are stale; no tests were run automatically.");
                    RebuildTree();
                }
                else if (capabilityChanged)
                {
                    status.Text = UiText.Get(reason);
                    UpdateCoverage();
                    UpdateButtons();
                }
            }
            catch (Exception error)
            {
                catalog = null;
                unavailableReason = error.Message;
                coverageUnavailableReason = error.Message;
                status.Text = error.Message;
                RebuildTree();
            }
        }

        /// <summary>Handles on visible changed for test explorer window.</summary>
        /// <param name="e">event args that supplies the e for this operation.</param>
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (freshnessTimer != null) UpdateFreshnessTimer();
        }

        /// <summary>Runs the non-executing catalog freshness check on the UI timer.</summary>
        /// <param name="sender">Freshness timer.</param><param name="e">Timer event data.</param>
        private void FreshnessTimer_Tick(object sender, EventArgs e) => CheckFreshness();

        /// <summary>Returns test descriptors currently present in the filtered tree.</summary>
        /// <returns>Visible test nodes in tree order, excluding category/module grouping nodes.</returns>
        internal IReadOnlyList<VbaTestDescriptor> VisibleTests()
        {
            if (catalog == null) return new VbaTestDescriptor[0];
            string term = search.Text.Trim();
            int filter = outcomeFilter.SelectedIndex - 1;
            return catalog.Tests.Where(test =>
                (term.Length == 0 || (test.Module + "." + test.Procedure + " " + string.Join(" ", test.Categories ?? new string[0]))
                    .IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) &&
                (filter < 0 || Outcome(test) == (VbaTestOutcome)filter)).ToArray();
        }

        /// <summary>Gets the stored result outcome for a descriptor, defaulting to NotRun.</summary>
        /// <param name="test">Descriptor whose last accepted result is displayed.</param>
        /// <returns>Stored outcome, or <see cref="VbaTestOutcome.NotRun"/> when no result exists.</returns>
        private VbaTestOutcome Outcome(VbaTestDescriptor test)
        {
            if (results.TryGetValue(test.Id, out var result)) return result.Outcome;
            if (!string.IsNullOrEmpty(test.Diagnostic)) return VbaTestOutcome.Blocked;
            if (test.IgnoreReason != null) return VbaTestOutcome.Skipped;
            return VbaTestOutcome.NotRun;
        }

        /// <summary>Checks whether a stored result was produced against an older project revision.</summary>
        /// <param name="test">Test descriptor whose result revision is compared with the current catalog.</param>
        /// <returns><see langword="true"/> when a result exists for a different revision.</returns>
        private bool IsStale(VbaTestDescriptor test) => resultRevisions.TryGetValue(test.Id, out var revision) && revision != catalog?.Project.Revision;

        /// <summary>Rebuilds the filtered/grouped tree while preserving stable selection and checked test IDs.</summary>
        private void RebuildTree()
        {
            var selected = ReadTreeSelection(testTree.SelectedNode) ?? treeSelection;
            if (catalog != null && selected?.ProjectId != catalog.Project.Id) selected = null;
            treeSelection = selected;
            var checkedIds = new HashSet<string>(TestNodes().Where(node => node.Checked).Select(node => ((VbaTestDescriptor)node.Tag).Id));
            rebuilding = true;
            testTree.BeginUpdate();
            try
            {
                testTree.Nodes.Clear();
                if (catalog != null)
                {
                    var root = new TreeNode(catalog.Project.Name) { Tag = catalog };
                    testTree.Nodes.Add(root);
                    var visible = VisibleTests();
                    if (grouping.SelectedIndex <= 0)
                    {
                        foreach (var module in catalog.Modules)
                        {
                            var tests = visible.Where(test => test.Module == module.Name).ToArray();
                            if (tests.Length == 0 && string.IsNullOrEmpty(module.Diagnostic)) continue;
                            var moduleNode = new TreeNode(module.Name + (string.IsNullOrEmpty(module.Diagnostic) ? "" : " — " + UiText.Get("Blocked"))) { Tag = module };
                            root.Nodes.Add(moduleNode);
                            AddTestNodes(moduleNode, tests, checkedIds);
                        }
                    }
                    else
                    {
                        var groups = grouping.SelectedIndex == 1
                            ? visible.GroupBy(Outcome).Select(group => new TestGroup { Grouping = 1, Key = group.Key.ToString(), Name = UiText.Get(group.Key.ToString()), Tests = group.ToArray() })
                            : visible.SelectMany(test => (test.Categories == null || test.Categories.Length == 0 ? new string[] { null } : test.Categories).Distinct(StringComparer.OrdinalIgnoreCase)
                                .Select(category => new { category, test })).GroupBy(pair => pair.category, StringComparer.OrdinalIgnoreCase)
                                .Select(group => new TestGroup { Grouping = 2, Key = group.Key, Name = group.Key ?? UiText.Get("Uncategorized"), Tests = group.Select(pair => pair.test).ToArray() });
                        foreach (var group in groups.OrderBy(group => group.Name, StringComparer.CurrentCultureIgnoreCase))
                        {
                            var node = new TreeNode(group.Name) { Tag = group };
                            root.Nodes.Add(node);
                            AddTestNodes(node, group.Tests, checkedIds);
                        }
                    }
                    root.Expand();
                    foreach (TreeNode module in root.Nodes) module.Expand();
                    UpdateGroupChecks();
                    // A temporarily absent outcome/filter scope stays empty, never widening to the project.
                    testTree.SelectedNode = selected == null ? root : FindSelection(root, selected);
                    if (selected == null) treeSelection = ReadTreeSelection(root);
                }
            }
            finally { testTree.EndUpdate(); rebuilding = false; }
            UpdateSummary();
            UpdateCoverage();
            UpdateDetails();
            UpdateButtons();
        }

        /// <summary>Grouping metadata attached to outcome/category tree nodes.</summary>
        private sealed class TestGroup
        {

            /// <summary>Localized node caption and stable grouping key.</summary>
            internal string Name, Key;

            /// <summary>Grouping selector value distinguishing outcome and category groups.</summary>
            internal int Grouping;

            /// <summary>Descriptors represented by this node before current search/outcome filtering.</summary>
            internal VbaTestDescriptor[] Tests;
        }

        /// <summary>Stable identity needed to restore a selected tree item after filtering or rediscovery.</summary>
        private sealed class TreeSelection
        {

            /// <summary>Project, module, test, and grouping keys; only the fields relevant to the selected node are populated.</summary>
            internal string ProjectId, Module, TestId, GroupKey;

            /// <summary>Grouping mode that produced <see cref="GroupKey"/>, or zero for project/module selection.</summary>
            internal int Grouping;
        }

        /// <summary>Extracts stable project/module/test/group identity from the selected node and its ancestors.</summary>
        /// <param name="node">Selected project, module, group, or test tree node.</param>
        /// <returns>Selection identity for a node under a discovered project, or <see langword="null"/> for no valid catalog node.</returns>
        private static TreeSelection ReadTreeSelection(TreeNode node)
        {
            if (node == null) return null;
            var root = node;
            while (root.Parent != null) root = root.Parent;
            var project = root.Tag as VbaTestCatalog;
            if (project == null) return null;
            var test = node.Tag as VbaTestDescriptor;
            var module = node.Tag as VbaTestModule;
            var group = node.Tag as TestGroup ?? node.Parent?.Tag as TestGroup;
            return new TreeSelection { ProjectId = project.Project.Id, Module = module?.Name,
                TestId = test?.Id, Grouping = group?.Grouping ?? 0, GroupKey = group?.Key };
        }

        /// <summary>Finds the closest rebuilt node matching a saved test, module, or group identity.</summary>
        /// <param name="root">Root node for the current project's rebuilt tree.</param><param name="selected">Identity captured before the rebuild.</param>
        /// <returns>Matching node, falling back from an exact group alias to the test alias, or to the project root.</returns>
        private static TreeNode FindSelection(TreeNode root, TreeSelection selected)
        {
            if (selected.TestId != null)
            {
                var aliases = root.Nodes.Cast<TreeNode>().SelectMany(group => group.Nodes.Cast<TreeNode>())
                    .Where(node => (node.Tag as VbaTestDescriptor)?.Id == selected.TestId).ToArray();
                return aliases.FirstOrDefault(node => node.Parent.Tag is TestGroup group && group.Grouping == selected.Grouping &&
                    string.Equals(group.Key, selected.GroupKey, StringComparison.OrdinalIgnoreCase)) ?? aliases.FirstOrDefault();
            }
            if (selected.Module != null) return root.Nodes.Cast<TreeNode>().FirstOrDefault(node => node.Tag is VbaTestModule module &&
                string.Equals(module.Name, selected.Module, StringComparison.OrdinalIgnoreCase));
            if (selected.Grouping != 0) return root.Nodes.Cast<TreeNode>().FirstOrDefault(node => node.Tag is TestGroup group &&
                group.Grouping == selected.Grouping && string.Equals(group.Key, selected.GroupKey, StringComparison.OrdinalIgnoreCase));
            return root;
        }

        /// <summary>Adds descriptor nodes beneath the selected module or group, restoring checked state by descriptor ID.</summary>
        /// <param name="parent">Module or group node receiving the test children.</param>
        /// <param name="tests">Visible descriptors to render under that node.</param>
        /// <param name="checkedIds">IDs checked before the tree rebuild.</param>
        private void AddTestNodes(TreeNode parent, IEnumerable<VbaTestDescriptor> tests, HashSet<string> checkedIds)
        {
            foreach (var test in tests)
            {
                var node = new TreeNode(TestCaption(test)) { Tag = test, Checked = checkedIds.Contains(test.Id), ForeColor = TestColor(test) };
                parent.Nodes.Add(node);
            }
        }

        /// <summary>Formats a test row with outcome, optional duration, grouping context, and stale-revision marker.</summary>
        /// <param name="test">Descriptor whose current result is displayed.</param>
        /// <returns>Localized tree caption; duration is shown only when a result exists.</returns>
        private string TestCaption(VbaTestDescriptor test)
        {
            string duration = results.TryGetValue(test.Id, out var result) ? "  " + result.Duration.TotalMilliseconds.ToString("0.##", CultureInfo.CurrentCulture) + " ms" : "";
            VbaTestOutcome outcome = Outcome(test);
            string symbol = outcome == VbaTestOutcome.Passed ? "✓ " : outcome == VbaTestOutcome.Failed || outcome == VbaTestOutcome.Error ? "✗ " : "";
            return symbol + (grouping.SelectedIndex > 0 ? test.Module + "." : "") + test.Procedure + "  [" + UiText.Get(outcome.ToString()) + "]" + duration + (IsStale(test) ? "  " + UiText.Get("Previous revision") : "");
        }

        /// <summary>Selects the theme color for a test outcome, with stale results muted.</summary>
        /// <param name="test">Descriptor whose outcome and result revision determine the color.</param>
        /// <returns>Muted for stale, success for passed, error for failed/error, or normal foreground otherwise.</returns>
        private System.Drawing.Color TestColor(VbaTestDescriptor test)
        {
            if (IsStale(test)) return UiTheme.Muted;
            switch (Outcome(test))
            {
                case VbaTestOutcome.Passed: return UiTheme.Success;
                case VbaTestOutcome.Failed:
                case VbaTestOutcome.Error: return UiTheme.Error;
                default: return UiTheme.Foreground;
            }
        }

        /// <summary>Rebuilds painted row colors on the owner thread after a global theme change.</summary>
        private void TestThemeChanged()
        {
            if (IsDisposed) return;
            if (!IsHandleCreated && Thread.CurrentThread.ManagedThreadId != ownerThread) return;
            if (InvokeRequired) { BeginInvoke(new Action(TestThemeChanged)); return; }
            RebuildTree();
        }

        /// <summary>Refreshes the result tree after the form becomes visible.</summary>
        /// <param name="e">Form shown event data.</param>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            RebuildTree();
        }

        /// <summary>Enumerates descriptor nodes below the root and one grouping level.</summary>
        /// <returns>All visible test leaf nodes, excluding project/module/group nodes.</returns>
        private IEnumerable<TreeNode> TestNodes() => testTree.Nodes.Cast<TreeNode>().SelectMany(root => root.Nodes.Cast<TreeNode>()).SelectMany(module => module.Nodes.Cast<TreeNode>()).Where(node => node.Tag is VbaTestDescriptor);

        /// <summary>Returns checked tests, or the selected test when no test is checked.</summary>
        /// <returns>Distinct selected descriptors in tree order; an empty list when selection is only a grouping node.</returns>
        internal IReadOnlyList<VbaTestDescriptor> SelectedTests()
        {
            var checkedTests = TestNodes().Where(node => node.Checked).Select(node => (VbaTestDescriptor)node.Tag).GroupBy(test => test.Id).Select(group => group.First()).ToArray();
            if (checkedTests.Length > 0) return checkedTests;
            var selected = testTree.SelectedNode?.Tag as VbaTestDescriptor;
            return selected == null ? new VbaTestDescriptor[0] : new[] { selected };
        }

        /// <summary>Resolves the current project, module, group, or test node to its visible test scope.</summary>
        /// <returns>Visible tests in the selected group/module, or all visible project tests when the project root is selected.</returns>
        internal IReadOnlyList<VbaTestDescriptor> ScopeTests()
        {
            var tests = VisibleTests();
            if (testTree.SelectedNode == null && treeSelection != null) return new VbaTestDescriptor[0];
            var module = testTree.SelectedNode?.Tag as VbaTestModule;
            var selected = testTree.SelectedNode?.Tag as VbaTestDescriptor;
            var group = testTree.SelectedNode?.Tag as TestGroup ?? testTree.SelectedNode?.Parent?.Tag as TestGroup;
            if (group != null)
            {
                var ids = new HashSet<string>(group.Tests.Select(test => test.Id), StringComparer.Ordinal);
                return tests.Where(test => ids.Contains(test.Id)).ToArray();
            }
            string moduleName = module?.Name ?? selected?.Module;
            return moduleName == null ? tests : tests.Where(test => test.Module == moduleName).ToArray();
        }

        /// <summary>Executes the explicit checked selection, falling back to the selected test.</summary>
        /// <returns>Task for the run; no test runs when the resolved selection is empty.</returns>
        internal Task RunSelectedAsync() => RunTestsAsync(SelectedTests());

        /// <summary>Executes the visible tests in the selected tree scope.</summary>
        /// <returns>Task for the run; an empty scope completes without invoking the service.</returns>
        internal Task RunScopeAsync() => RunTestsAsync(ScopeTests());

        /// <summary>Reruns visible failed/error tests whose result revision still matches the current catalog.</summary>
        /// <returns>Task for the explicitly requested rerun.</returns>
        internal Task RerunFailedAsync() => RunTestsAsync(VisibleTests().Where(test => !IsStale(test) && (Outcome(test) == VbaTestOutcome.Failed || Outcome(test) == VbaTestOutcome.Error)).ToArray());

        /// <summary>Shows the review and consent prompt before measuring coverage for the explicit selection.</summary>
        /// <returns>Task for coverage review and, only after consent, the host run.</returns>
        internal Task RunSelectedCoverageAsync() => ReviewCoverageAndRunAsync(SelectedTests());

        /// <summary>Shows the review and consent prompt before measuring coverage for the selected tree scope.</summary>
        /// <returns>Task for coverage review and the explicitly approved host run.</returns>
        internal Task RunScopeCoverageAsync() => ReviewCoverageAndRunAsync(ScopeTests());

        /// <summary>Freezes the selection/revision, presents side-effect and storage details, and requires affirmative user consent.</summary>
        /// <param name="selected">Exact test descriptors to review and potentially measure.</param>
        /// <returns>Completed task when refused or unavailable; otherwise the owner-dispatched coverage run.</returns>
        private Task ReviewCoverageAndRunAsync(IReadOnlyList<VbaTestDescriptor> selected)
        {
            if (IsDisposed || Disposing || running || catalog == null || selected.Count == 0 || !(service is IVbaTestCoverageExplorerService) || !string.IsNullOrEmpty(coverageUnavailableReason)) return Task.CompletedTask;
            var frozen = selected.GroupBy(test => test.Id).Select(group => group.First()).ToArray();
            var reviewCatalog = catalog;
            string text = UiText.Get("This run measures procedure-entry coverage, not statement or branch coverage.") + Environment.NewLine + Environment.NewLine +
                UiText.Get("Project") + ": " + catalog.Project.Name + Environment.NewLine + catalog.Project.HostPath + Environment.NewLine +
                UiText.Get("Revision: ") + catalog.Project.Revision + Environment.NewLine +
                UiText.Get("Selected tests") + ": " + frozen.Length + Environment.NewLine + string.Join(Environment.NewLine, frozen.Select(test => test.Module + "." + test.Procedure)) + Environment.NewLine + Environment.NewLine +
                UiText.Get("A separate host document copy is instrumented and executed in the host application. Original source is never instrumented. Application-level handlers and tests may have external side effects with the host's privileges.") + Environment.NewLine +
                UiText.Get("Copies and measurement files are retained in:") + Environment.NewLine + @"%LOCALAPPDATA%\VBAi\CoverageRuns" + Environment.NewLine + Environment.NewLine + UiText.Get("Run this reviewed coverage selection?");
            bool accepted;
            reviewingCoverage = true;
            try { accepted = ConfirmCoverage(this, text); }
            finally { reviewingCoverage = false; }
            if (!accepted || IsDisposed || Disposing) return Task.CompletedTask;
            // The reviewed identity/revision remains authoritative; service guards revalidate it.
            if (!ReferenceEquals(catalog, reviewCatalog)) { status.Text = UiText.Get("The reviewed project changed. Review coverage again."); return Task.CompletedTask; }
            // Return the owner-dispatched run directly; no forwarding await may recapture a disposed UI context.
            return RunTestsAsync(frozen, true);
        }

        /// <summary>Runs a frozen, deduplicated test selection on the owning host service and records results against its catalog revision.</summary>
        /// <param name="requested">Descriptors explicitly selected from the current catalog.</param>
        /// <param name="measuredCoverage"><see langword="true"/> only after the coverage review was accepted.</param>
        /// <returns>Task for the run; unknown outcomes are surfaced without automatic retry.</returns>
        private async Task RunTestsAsync(IReadOnlyList<VbaTestDescriptor> requested, bool measuredCoverage = false)
        {
            if (IsDisposed || Disposing || running || service == null || catalog == null || requested.Count == 0 ||
                !string.IsNullOrEmpty(measuredCoverage ? coverageUnavailableReason : unavailableReason)) return;
            var frozen = requested.GroupBy(test => test.Id).Select(group => group.First()).ToArray();
            var runCatalog = catalog;
            InitializeRunContinuationDispatcher();
            running = true;
            cancellation = new CancellationTokenSource();
            try
            {
                foreach (var test in frozen) { results.Remove(test.Id); resultRevisions.Remove(test.Id); }
                RebuildTree();
                status.Text = UiText.Get("Running tests...");
                var run = measuredCoverage
                    ? await AwaitOwner(((IVbaTestCoverageExplorerService)service).RunCoverageAsync(runCatalog, frozen, result => AcceptResult(result, runCatalog.Project.Revision), cancellation.Token))
                    : await AwaitOwner(service.RunAsync(runCatalog, frozen, result => AcceptResult(result, runCatalog.Project.Revision), cancellation.Token));
                if (IsDisposed || Disposing) return;
                foreach (var result in run.Results) AcceptResult(result, run.Revision);
                humanReport.Text = VbaTestReports.Human(run);
                compactReport.Text = VbaTestReports.Compact(run);
                lastCoverage = run.Coverage;
                UpdateCoverage();
                status.Text = !string.IsNullOrEmpty(run.Error) ? UiText.Get(run.Error) : run.OutcomeUnknown ? UiText.Get("Test outcome is unknown. No automatic retry was performed.") : UiText.Get("Test run completed.");
            }
            catch (OperationCanceledException) { if (!IsDisposed && !Disposing) status.Text = UiText.Get("Test run stopped."); }
            catch (Exception ex) { if (!IsDisposed && !Disposing) status.Text = ex.Message; }
            finally
            {
                running = false;
                cancellation.Dispose();
                cancellation = null;
                try
                {
                    if (!IsDisposed && !Disposing)
                    {
                        try { unavailableReason = service.ExecutionUnavailableReason(catalog); coverageUnavailableReason = ReadCoverageUnavailableReason(catalog); }
                        catch (Exception ex) { unavailableReason = ex.Message; }
                        UpdateButtons();
                        if (closeAfterRun) Close();
                    }
                }
                finally { ReleaseRunContinuationDispatcher(); }
            }
        }

        /// <summary>Requires run owner for test explorer window.</summary>
        private void RequireRunOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("The test explorer continuation requires its owning UI thread.");
        }

        /// <summary>Handles initialize run continuation dispatcher for test explorer window.</summary>
        private void InitializeRunContinuationDispatcher()
        {
            RequireRunOwner();
            var control = CreateRunDispatcher();
            try { var handle = control.Handle; lock (runDispatchGate) runContinuationDispatcher = control; }
            catch { control.Dispose(); throw; }
        }

        /// <summary>Releases run continuation dispatcher for test explorer window.</summary>
        private void ReleaseRunContinuationDispatcher()
        {
            RequireRunOwner();
            lock (runDispatchGate)
            {
                var control = runContinuationDispatcher;
                runContinuationDispatcher = null;
                control?.Dispose();
            }
        }

        /// <summary>Runs window disposed for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void RunWindowDisposed(object sender, EventArgs e)
        {
            RequireRunOwner();
            cancellation?.Cancel();
            // A native/add-in teardown may force Dispose despite the cooperative close guard.
            // This independent handle remains alive until the pending result/exception settles.
            if (!running) ReleaseRunContinuationDispatcher();
        }

        /// <summary>Handles await owner for test explorer window.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        /// <returns>vba test owner awaitable&lt;t&gt; produced by the operation for await owner on test explorer window.</returns>
        private VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
        {
            RequireRunOwner();
            var dispatcher = runContinuationDispatcher;
            if (dispatcher == null) throw new InvalidOperationException("The test explorer run dispatcher is unavailable.");
            return new VbaTestOwnerAwaitable<T>(task, action => dispatcher.BeginInvoke(action), RequireRunOwner);
        }

        /// <summary>Handles accept result for test explorer window.</summary>
        /// <param name="result">vba test result that supplies the result for this operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        private void AcceptResult(VbaTestResult result, string revision)
        {
            if (result?.Test == null) return;
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
            {
                lock (runDispatchGate)
                {
                    if (IsDisposed || Disposing || runContinuationDispatcher == null) return;
                    runContinuationDispatcher.BeginInvoke(new Action(() => AcceptResult(result, revision)));
                }
                return;
            }
            if (IsDisposed || Disposing) return;
            results[result.Test.Id] = result;
            resultRevisions[result.Test.Id] = revision;
            RebuildTree();
        }

        /// <summary>Updates summary for test explorer window.</summary>
        private void UpdateSummary()
        {
            var tests = catalog?.Tests.ToArray() ?? new VbaTestDescriptor[0];
            int passed = tests.Count(test => !IsStale(test) && Outcome(test) == VbaTestOutcome.Passed);
            int stale = tests.Count(IsStale);
            var counts = tests.Where(test => !IsStale(test)).GroupBy(Outcome).Select(group => UiText.Get(group.Key.ToString()) + ": " + group.Count());
            int completed = tests.Count(test => !IsStale(test) && (Outcome(test) == VbaTestOutcome.Passed || Outcome(test) == VbaTestOutcome.Failed || Outcome(test) == VbaTestOutcome.Error));
            summary.Text = UiText.Get("Tests") + ": " + tests.Length + "  ·  " + string.Join("  ·  ", counts) + (stale > 0 ? "  ·  " + UiText.Get("Previous revision") + ": " + stale : "") +
                "  ·  " + UiText.Get("Pass rate") + ": " + (completed > 0 ? (100.0 * passed / completed).ToString("0.##", CultureInfo.CurrentCulture) + "%" : "—");
        }

        /// <summary>Updates coverage for test explorer window.</summary>
        private void UpdateCoverage()
        {
            if (lastCoverage == null)
            {
                coverage.Text = UiText.Get("VBA procedure coverage: not measured.") + (string.IsNullOrEmpty(coverageUnavailableReason) ? "" : " " + UiText.Get(coverageUnavailableReason));
                return;
            }
            bool stale = catalog == null || lastCoverage.Revision != catalog.Project.Revision;
            string measure = !lastCoverage.Available ? UiText.Get("Unavailable") : lastCoverage.Percent.HasValue && lastCoverage.Eligible.HasValue && lastCoverage.Hit.HasValue
                ? lastCoverage.Percent.Value.ToString("0.##", CultureInfo.CurrentCulture) + "% (" + lastCoverage.Hit + "/" + lastCoverage.Eligible + ")"
                : UiText.Get("Unknown");
            coverage.Text = UiText.Get("VBA procedure coverage") + ": " + measure + "  ·  " +
                UiText.Get(lastCoverage.Complete ? "Complete measurement" : "Partial measurement") +
                (stale ? "  ·  " + UiText.Get("Previous revision") : "") + "  ·  " + UiText.Get("Statement and branch coverage are not measured.") +
                (lastCoverage.Diagnostics.Count == 0 ? "" : Environment.NewLine + string.Join(Environment.NewLine, lastCoverage.Diagnostics.Select(UiText.Get)));
        }

        /// <summary>Updates details for test explorer window.</summary>
        private void UpdateDetails()
        {
            var test = testTree.SelectedNode?.Tag as VbaTestDescriptor;
            if (test == null)
            {
                var module = testTree.SelectedNode?.Tag as VbaTestModule;
                details.Text = module?.Diagnostic ?? (catalog == null ? "" : catalog.Project.Name + Environment.NewLine + catalog.Project.HostPath + Environment.NewLine + string.Join(Environment.NewLine, catalog.Diagnostics));
                return;
            }
            results.TryGetValue(test.Id, out var result);
            details.Text = test.Module + "." + test.Procedure + Environment.NewLine + UiText.Get("Source line") + ": " + test.Line + Environment.NewLine +
                UiText.Get("Categories") + ": " + string.Join(", ", test.Categories ?? new string[0]) + Environment.NewLine +
                UiText.Get("Outcome") + ": " + UiText.Get(Outcome(test).ToString()) + Environment.NewLine +
                (result == null ? "" : UiText.Get("Duration") + ": " + result.Duration.TotalMilliseconds.ToString("0.##", CultureInfo.CurrentCulture) + " ms" + Environment.NewLine +
                    UiText.Get("Phase") + ": " + result.Phase + Environment.NewLine + UiText.Get("Run revision") + ": " + resultRevisions[test.Id] + Environment.NewLine + result.Message + Environment.NewLine) +
                (result == null || result.ErrorNumber == 0 ? "" : UiText.Get("VBA error") + ": " + result.ErrorNumber + Environment.NewLine) +
                (IsStale(test) ? UiText.Get("Result belongs to a previous source revision.") + Environment.NewLine : "") + test.Diagnostic + Environment.NewLine + test.IgnoreReason;
        }

        /// <summary>Updates buttons for test explorer window.</summary>
        private void UpdateButtons()
        {
            bool available = service != null && catalog != null && !running;
            bool canRun = available && string.IsNullOrEmpty(unavailableReason);
            refresh.Enabled = service != null && !running;
            projectList.Enabled = !running;
            search.Enabled = !running;
            outcomeFilter.Enabled = !running;
            grouping.Enabled = !running;
            testTree.Enabled = !running;
            runSelected.Enabled = canRun && SelectedTests().Count > 0;
            runScope.Enabled = canRun && ScopeTests().Count > 0;
            bool canMeasure = available && service is IVbaTestCoverageExplorerService && string.IsNullOrEmpty(coverageUnavailableReason);
            runSelectedCoverage.Enabled = canMeasure && SelectedTests().Count > 0;
            runScopeCoverage.Enabled = canMeasure && ScopeTests().Count > 0;
            rerunFailed.Enabled = canRun && VisibleTests().Any(test => !IsStale(test) && (Outcome(test) == VbaTestOutcome.Failed || Outcome(test) == VbaTestOutcome.Error));
            source.Enabled = available && testTree.SelectedNode?.Tag is VbaTestDescriptor;
            installSupport.Enabled = available;
            stop.Enabled = running && cancellation != null && !cancellation.IsCancellationRequested;
            copyReport.Enabled = !running && ReportText().Length > 0;
            exportReport.Enabled = copyReport.Enabled;
        }

        /// <summary>Handles refresh click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void Refresh_Click(object sender, EventArgs e) => RefreshProjects();

        /// <summary>Handles project list selected index changed for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void ProjectList_SelectedIndexChanged(object sender, EventArgs e) => DiscoverSelectedProject();

        /// <summary>Handles project list format for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">list control convert event args that supplies the e for this operation.</param>
        private void ProjectList_Format(object sender, ListControlConvertEventArgs e)
        {
            if (e.ListItem is VbaTestProjectSnapshot project)
                e.Value = project.Name + (string.IsNullOrEmpty(project.HostPath) ? "" : " — " + project.HostPath);
        }

        /// <summary>Handles filter changed for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void Filter_Changed(object sender, EventArgs e) { if (!running) RebuildTree(); }

        /// <summary>Handles test tree after select for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">tree view event args that supplies the e for this operation.</param>
        private void TestTree_AfterSelect(object sender, TreeViewEventArgs e) { if (!rebuilding) { treeSelection = ReadTreeSelection(e.Node); UpdateDetails(); UpdateButtons(); } }

        /// <summary>Handles test tree after check for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">tree view event args that supplies the e for this operation.</param>
        private void TestTree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (rebuilding) return;
            rebuilding = true;
            try
            {
                var affected = new HashSet<string>(DescendantsAndSelf(e.Node).Select(node => node.Tag as VbaTestDescriptor)
                    .Where(test => test != null).Select(test => test.Id), StringComparer.Ordinal);
                foreach (var node in TestNodes().Where(node => affected.Contains(((VbaTestDescriptor)node.Tag).Id))) node.Checked = e.Node.Checked;
                UpdateGroupChecks();
            }
            finally { rebuilding = false; }
            UpdateButtons();
        }

        /// <summary>Handles descendants and self for test explorer window.</summary>
        /// <param name="node">tree node that supplies the node for this operation.</param>
        /// <returns>i enumerable&lt;tree node&gt; produced by the operation for descendants and self on test explorer window.</returns>
        private static IEnumerable<TreeNode> DescendantsAndSelf(TreeNode node)
        {
            yield return node;
            foreach (TreeNode child in node.Nodes)
                foreach (var descendant in DescendantsAndSelf(child)) yield return descendant;
        }

        /// <summary>Updates group checks for test explorer window.</summary>
        private void UpdateGroupChecks()
        {
            foreach (TreeNode root in testTree.Nodes)
            {
                foreach (TreeNode group in root.Nodes)
                    group.Checked = group.Nodes.Count > 0 && group.Nodes.Cast<TreeNode>().All(node => node.Checked);
                root.Checked = root.Nodes.Count > 0 && root.Nodes.Cast<TreeNode>().All(node => node.Checked);
            }
        }

        /// <summary>Runs selected click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private async void RunSelected_Click(object sender, EventArgs e) => await RunSelectedAsync();

        /// <summary>Runs scope click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private async void RunScope_Click(object sender, EventArgs e) => await RunScopeAsync();

        /// <summary>Handles rerun failed click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private async void RerunFailed_Click(object sender, EventArgs e) => await RerunFailedAsync();

        /// <summary>Runs selected coverage click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private async void RunSelectedCoverage_Click(object sender, EventArgs e) => await RunSelectedCoverageAsync();

        /// <summary>Runs scope coverage click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private async void RunScopeCoverage_Click(object sender, EventArgs e) => await RunScopeCoverageAsync();

        /// <summary>Stops click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void Stop_Click(object sender, EventArgs e)
        {
            cancellation?.Cancel();
            status.Text = UiText.Get("Stopping after the current test; waiting for its verified result.");
            UpdateButtons();
        }

        /// <summary>Handles source click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void Source_Click(object sender, EventArgs e)
        {
            if (running || catalog == null || !(testTree.SelectedNode?.Tag is VbaTestDescriptor test)) return;
            try { service.Navigate(catalog, test); } catch (Exception ex) { status.Text = ex.Message; }
        }

        /// <summary>Handles test tree node mouse double click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">tree node mouse click event args that supplies the e for this operation.</param>
        private void TestTree_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e) => Source_Click(sender, e);

        /// <summary>Handles report text for test explorer window.</summary>
        /// <returns>Text produced by the operation for report text on test explorer window.</returns>
        private string ReportText() => resultTabs.SelectedTab == compactTab ? compactReport.Text : resultTabs.SelectedTab == humanTab ? humanReport.Text : "";

        /// <summary>Handles result tabs selected index changed for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void ResultTabs_SelectedIndexChanged(object sender, EventArgs e) => UpdateButtons();

        /// <summary>Handles copy report click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void CopyReport_Click(object sender, EventArgs e)
        {
            if (running || ReportText().Length == 0) return;
            try { WriteReportClipboard(ReportText()); } catch (Exception ex) { status.Text = ex.Message; }
        }

        /// <summary>Handles export report click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void ExportReport_Click(object sender, EventArgs e)
        {
            if (running) return;
            string text = ReportText();
            if (text.Length == 0) return;
            bool compact = resultTabs.SelectedTab == compactTab;
            try
            {
                string path = ChooseReportExportPath(this, compact);
                if (path != null) File.WriteAllText(path, text, new System.Text.UTF8Encoding(false));
            }
            catch (Exception ex) { if (!IsDisposed && !Disposing) status.Text = ex.Message; }
        }

        /// <summary>Handles select report export path for test explorer window.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <param name="compact">Indicates whether compact is enabled.</param>
        /// <param name="show">func&lt;save file dialog, i win32 window, dialog result&gt; that supplies the show for this operation.</param>
        /// <returns>Text produced by the operation for select report export path on test explorer window.</returns>
        internal static string SelectReportExportPath(IWin32Window owner, bool compact, Func<SaveFileDialog, IWin32Window, DialogResult> show = null)
        {
            using (var dialog = new SaveFileDialog { FileName = compact ? "vba-test-results.json" : "vba-test-results.txt", Filter = compact ? "JSON (*.json)|*.json" : "Text (*.txt)|*.txt" })
                return (show == null ? dialog.ShowDialog(owner) : show(dialog, owner)) == DialogResult.OK ? dialog.FileName : null;
        }

        /// <summary>Handles install support click for test explorer window.</summary>
        /// <param name="sender">object that supplies the sender for this operation.</param>
        /// <param name="e">event args that supplies the e for this operation.</param>
        private void InstallSupport_Click(object sender, EventArgs e)
        {
            if (running || service == null || catalog == null) return;
            try { service.InstallSupport(catalog); RefreshProjects(); } catch (Exception ex) { status.Text = ex.Message; }
        }

        /// <summary>Handles on form closing for test explorer window.</summary>
        /// <param name="e">form closing event args that supplies the e for this operation.</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running) { e.Cancel = true; closeAfterRun = true; Stop_Click(this, EventArgs.Empty); }
            else if (reviewingCoverage) e.Cancel = true;
            base.OnFormClosing(e);
        }
    }
}
