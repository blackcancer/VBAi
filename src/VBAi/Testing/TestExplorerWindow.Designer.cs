using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Designer-owned layout; no project access occurs during construction.</summary>
    internal sealed partial class TestExplorerWindow
    {

        /// <summary>Owns designer-created child components disposed with the form.</summary>
        private IContainer components;

        /// <summary>Places filters, test tree, details, and reports in the fixed form layout.</summary>
        private TableLayoutPanel layout;

        /// <summary>Hosts refresh, run, coverage, source, support, stop, and report commands.</summary>
        private FlowLayoutPanel commands;

        /// <summary>Contains project, search, outcome, and grouping selectors above the tree.</summary>
        private TableLayoutPanel filters;

        /// <summary>Selects the project catalog currently shown in the explorer.</summary>
        private UiComboBox projectList;

        /// <summary>Filters visible descriptors by their displayed names and metadata.</summary>
        private UiTextBox search;

        /// <summary>Limits visible test results to the selected outcome category.</summary>
        private UiComboBox outcomeFilter;

        /// <summary>Chooses how test descriptors are grouped in the tree.</summary>
        private UiComboBox grouping;

        /// <summary>Displays module groups, test descriptors, and checkable selection state.</summary>
        private TreeView testTree;

        /// <summary>Separates the selectable test tree from the detail and report panes.</summary>
        private SplitContainer split;

        /// <summary>Shows the selected descriptor or result details without executing VBA.</summary>
        private UiTextBox details;

        /// <summary>Switches between descriptor details and human or compact run reports.</summary>
        private ThemedTabControl resultTabs;

        /// <summary>Tab containing selected test metadata and diagnostics.</summary>
        private TabPage detailsTab;

        /// <summary>Tab containing the human-readable execution report.</summary>
        private TabPage humanTab;

        /// <summary>Tab containing the compact machine-friendly execution report.</summary>
        private TabPage compactTab;

        /// <summary>Read-only control displaying the human-readable report for copying.</summary>
        private UiTextBox humanReport;

        /// <summary>Read-only control displaying the compact report representation.</summary>
        private UiTextBox compactReport;

        /// <summary>Copies the currently selected report representation to the clipboard.</summary>
        private UiActionButton copyReport;

        /// <summary>Exports the selected report representation to a user-chosen file.</summary>
        private UiActionButton exportReport;

        /// <summary>Displays aggregate counts for the current catalog or run.</summary>
        private Label summary;

        /// <summary>Displays current project freshness and execution status.</summary>
        private Label status;

        /// <summary>Displays procedure-entry coverage counts and completeness.</summary>
        private Label coverage;

        /// <summary>Reloads the current project snapshot and test discovery catalog.</summary>
        private UiActionButton refresh;

        /// <summary>Runs the checked descriptors in the selected project.</summary>
        private UiActionButton runSelected;

        /// <summary>Runs all currently runnable tests in the selected project scope.</summary>
        private UiActionButton runScope;

        /// <summary>Runs checked descriptors with an explicit procedure-entry coverage review.</summary>
        private UiActionButton runSelectedCoverage;

        /// <summary>Runs the project test scope with an explicit procedure-entry coverage review.</summary>
        private UiActionButton runScopeCoverage;

        /// <summary>Periodically detects whether source changed since the displayed catalog was captured.</summary>
        private Timer freshnessTimer;

        /// <summary>Selects and runs tests whose prior result was failed or errored.</summary>
        private UiActionButton rerunFailed;

        /// <summary>Requests cooperative cancellation between native test calls.</summary>
        private UiActionButton stop;

        /// <summary>Navigates to the source declaration for the selected descriptor.</summary>
        private UiActionButton source;

        /// <summary>Installs or updates the project-local VBA test support module.</summary>
        private UiActionButton installSupport;

        /// <summary>Builds the designer-owned control hierarchy without accessing project services.</summary>
        private void InitializeComponent()
        {
            components = new Container();
            layout = new TableLayoutPanel();
            layout.Name = "layout";
            commands = new FlowLayoutPanel();
            commands.Name = "commands";
            filters = new TableLayoutPanel();
            filters.Name = "filters";
            projectList = new UiComboBox();
            projectList.Name = "projectList";
            search = new UiTextBox();
            search.Name = "search";
            outcomeFilter = new UiComboBox();
            outcomeFilter.Name = "outcomeFilter";
            grouping = new UiComboBox();
            grouping.Name = "grouping";
            testTree = new TreeView();
            testTree.Name = "testTree";
            split = new SplitContainer();
            split.Name = "split";
            details = new UiTextBox();
            resultTabs = new ThemedTabControl();
            resultTabs.Name = "resultTabs";
            detailsTab = new TabPage();
            detailsTab.Name = "detailsTab";
            humanTab = new TabPage();
            humanTab.Name = "humanTab";
            compactTab = new TabPage();
            compactTab.Name = "compactTab";
            humanReport = new UiTextBox();
            humanReport.Name = "humanReport";
            compactReport = new UiTextBox();
            compactReport.Name = "compactReport";
            copyReport = new UiActionButton();
            copyReport.Name = "copyReport";
            exportReport = new UiActionButton();
            exportReport.Name = "exportReport";
            summary = new Label();
            summary.Name = "summary";
            status = new Label();
            status.Name = "status";
            coverage = new Label();
            coverage.Name = "coverage";
            refresh = new UiActionButton();
            refresh.Name = "refresh";
            runSelected = new UiActionButton();
            runSelected.Name = "runSelected";
            runScope = new UiActionButton();
            runScope.Name = "runScope";
            runSelectedCoverage = new UiActionButton();
            runSelectedCoverage.Name = "runSelectedCoverage";
            runScopeCoverage = new UiActionButton();
            runScopeCoverage.Name = "runScopeCoverage";
            freshnessTimer = new Timer(components);
            rerunFailed = new UiActionButton();
            rerunFailed.Name = "rerunFailed";
            stop = new UiActionButton();
            stop.Name = "stop";
            source = new UiActionButton();
            source.Name = "source";
            installSupport = new UiActionButton();
            installSupport.Name = "installSupport";
            ((ISupportInitialize)split).BeginInit();
            split.Panel1.SuspendLayout();
            split.Panel2.SuspendLayout();
            split.SuspendLayout();
            layout.SuspendLayout();
            commands.SuspendLayout();
            filters.SuspendLayout();
            SuspendLayout();
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1120, 680);
            MinimumSize = new Size(820, 480);
            Font = new Font("Segoe UI", 9F);
            Name = "TestExplorerWindow";
            Text = "VBAi Test Explorer";
            StartPosition = FormStartPosition.CenterParent;
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(12);
            layout.ColumnCount = 1;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowCount = 6;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            commands.Dock = DockStyle.Fill;
            commands.AutoSize = true;
            commands.WrapContents = true;
            commands.Margin = new Padding(0, 0, 0, 8);
            refresh.Text = "Refresh tests";
            refresh.Symbol = UiSymbol.Refresh;
            refresh.AutoSize = true;
            refresh.Click += Refresh_Click;
            runSelected.Text = "Run selected tests";
            runSelected.Symbol = UiSymbol.Play;
            runSelected.Primary = true;
            runSelected.AutoSize = true;
            runSelected.Click += RunSelected_Click;
            runScope.Text = "Run visible scope";
            runScope.Symbol = UiSymbol.Play;
            runScope.AutoSize = true;
            runScope.Click += RunScope_Click;
            runSelectedCoverage.Text = "Run selected with coverage...";
            runSelectedCoverage.Symbol = UiSymbol.Inspect;
            runSelectedCoverage.AutoSize = true;
            runSelectedCoverage.Click += RunSelectedCoverage_Click;
            runScopeCoverage.Text = "Run visible scope with coverage...";
            runScopeCoverage.Symbol = UiSymbol.Inspect;
            runScopeCoverage.AutoSize = true;
            runScopeCoverage.Click += RunScopeCoverage_Click;
            freshnessTimer.Interval = 2000;
            freshnessTimer.Tick += FreshnessTimer_Tick;
            rerunFailed.Text = "Rerun failed tests";
            rerunFailed.Symbol = UiSymbol.Refresh;
            rerunFailed.AutoSize = true;
            rerunFailed.Click += RerunFailed_Click;
            stop.Text = "Stop after current test";
            stop.Symbol = UiSymbol.Stop;
            stop.AutoSize = true;
            stop.Click += Stop_Click;
            source.Text = "Open test source";
            source.Symbol = UiSymbol.Code;
            source.AutoSize = true;
            source.Click += Source_Click;
            installSupport.Text = "Install test support...";
            installSupport.Symbol = UiSymbol.Add;
            installSupport.AutoSize = true;
            installSupport.Click += InstallSupport_Click;
            copyReport.Text = "Copy result report";
            copyReport.Symbol = UiSymbol.Copy;
            copyReport.AutoSize = true;
            copyReport.Click += CopyReport_Click;
            exportReport.Text = "Export result report...";
            exportReport.Symbol = UiSymbol.Save;
            exportReport.AutoSize = true;
            exportReport.Click += ExportReport_Click;
            commands.Controls.AddRange(new Control[] { refresh, runSelected, runScope, runSelectedCoverage, runScopeCoverage, rerunFailed, stop, source, installSupport, copyReport, exportReport });
            filters.Dock = DockStyle.Fill;
            filters.AutoSize = true;
            filters.Margin = new Padding(0, 0, 0, 10);
            filters.ColumnCount = 6;
            filters.RowCount = 2;
            filters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            filters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            var projectLabel = new Label { Text = "Project", AutoSize = true, Anchor = AnchorStyles.Left };
            var searchLabel = new Label { Text = "Name or category", AutoSize = true, Anchor = AnchorStyles.Left };
            var outcomeLabel = new Label { Text = "Outcome", AutoSize = true, Anchor = AnchorStyles.Left };
            var groupLabel = new Label { Text = "Group by", AutoSize = true, Anchor = AnchorStyles.Left };
            projectList.Dock = DockStyle.Fill;
            projectList.AccessibleName = "Test project selection";
            projectList.DropDownStyle = ComboBoxStyle.DropDownList;
            projectList.DisplayMember = "Name";
            projectList.FormattingEnabled = true;
            projectList.Format += ProjectList_Format;
            projectList.SelectedIndexChanged += ProjectList_SelectedIndexChanged;
            search.Dock = DockStyle.Fill;
            search.AccessibleName = "Test search";
            search.TextChanged += Filter_Changed;
            outcomeFilter.Dock = DockStyle.Fill;
            outcomeFilter.AccessibleName = "Test outcome filter";
            outcomeFilter.DropDownStyle = ComboBoxStyle.DropDownList;
            outcomeFilter.SelectedIndexChanged += Filter_Changed;
            grouping.Dock = DockStyle.Fill;
            grouping.AccessibleName = "Test grouping";
            grouping.DropDownStyle = ComboBoxStyle.DropDownList;
            grouping.SelectedIndexChanged += Filter_Changed;
            filters.Controls.Add(projectLabel, 0, 0);
            filters.Controls.Add(projectList, 1, 0);
            filters.Controls.Add(searchLabel, 2, 0);
            filters.Controls.Add(search, 3, 0);
            filters.Controls.Add(outcomeLabel, 4, 0);
            filters.Controls.Add(outcomeFilter, 5, 0);
            filters.Controls.Add(groupLabel, 0, 1);
            filters.Controls.Add(grouping, 1, 1);
            split.Dock = DockStyle.Fill;
            split.Orientation = Orientation.Vertical;
            split.Size = new Size(1096, 470);
            split.SplitterDistance = 610;
            split.Panel1MinSize = 240;
            split.Panel2MinSize = 220;
            split.Margin = new Padding(0);
            testTree.Dock = DockStyle.Fill;
            testTree.AccessibleName = "Test hierarchy";
            testTree.RightToLeft = RightToLeft.No;
            testTree.CheckBoxes = true;
            testTree.HideSelection = false;
            testTree.BorderStyle = BorderStyle.FixedSingle;
            testTree.AfterSelect += TestTree_AfterSelect;
            testTree.AfterCheck += TestTree_AfterCheck;
            testTree.NodeMouseDoubleClick += TestTree_NodeMouseDoubleClick;
            details.Dock = DockStyle.Fill;
            details.AccessibleName = "Test result details";
            details.Name = "details";
            details.RightToLeft = RightToLeft.No;
            details.Multiline = true;
            details.ReadOnly = true;
            details.ScrollBars = ScrollBars.Vertical;
            details.Margin = new Padding(8, 0, 0, 0);
            resultTabs.Dock = DockStyle.Fill;
            detailsTab.Text = "Test details";
            humanTab.Text = "Readable report";
            compactTab.Text = "LLM report (JSON)";
            detailsTab.Padding = new Padding(8);
            humanTab.Padding = new Padding(8);
            compactTab.Padding = new Padding(8);
            humanReport.Dock = DockStyle.Fill;
            humanReport.AccessibleName = "Human-readable test report";
            humanReport.RightToLeft = RightToLeft.No;
            humanReport.Multiline = true;
            humanReport.ReadOnly = true;
            humanReport.ScrollBars = ScrollBars.Both;
            humanReport.WordWrap = false;
            compactReport.Dock = DockStyle.Fill;
            compactReport.AccessibleName = "Compact test report for LLM";
            compactReport.RightToLeft = RightToLeft.No;
            compactReport.Multiline = true;
            compactReport.ReadOnly = true;
            compactReport.ScrollBars = ScrollBars.Both;
            compactReport.WordWrap = false;
            detailsTab.Controls.Add(details);
            humanTab.Controls.Add(humanReport);
            compactTab.Controls.Add(compactReport);
            resultTabs.TabPages.AddRange(new[] { detailsTab, humanTab, compactTab });
            resultTabs.SelectedIndexChanged += ResultTabs_SelectedIndexChanged;
            split.Panel1.Controls.Add(testTree);
            split.Panel2.Controls.Add(resultTabs);
            summary.AutoSize = true;
            summary.Dock = DockStyle.Fill;
            summary.Padding = new Padding(0, 8, 0, 0);
            status.AutoSize = true;
            status.Dock = DockStyle.Fill;
            status.Padding = new Padding(0, 6, 0, 0);
            coverage.AutoSize = true;
            coverage.Dock = DockStyle.Fill;
            coverage.Padding = new Padding(0, 6, 0, 0);
            coverage.Text = "VBA procedure coverage: not measured.";
            layout.Controls.Add(commands, 0, 0);
            layout.Controls.Add(filters, 0, 1);
            layout.Controls.Add(split, 0, 2);
            layout.Controls.Add(summary, 0, 3);
            layout.Controls.Add(status, 0, 4);
            layout.Controls.Add(coverage, 0, 5);
            Controls.Add(layout);
            split.Panel1.ResumeLayout(false);
            split.Panel2.ResumeLayout(false);
            split.Panel2.PerformLayout();
            ((ISupportInitialize)split).EndInit();
            split.ResumeLayout(false);
            commands.ResumeLayout(false);
            commands.PerformLayout();
            filters.ResumeLayout(false);
            filters.PerformLayout();
            layout.ResumeLayout(false);
            layout.PerformLayout();
            ResumeLayout(false);
        }

        /// <summary>Releases the component container when the form is disposed.</summary>
        /// <param name="disposing">True when managed components should also be disposed.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }
    }
}
