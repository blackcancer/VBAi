using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class TestExplorerWindowTests
    {
        [STATestMethod]
        public void InitialNativePlacementRepairsOnlyCollapsedOwnedSitesAndKeepsDesignerDimensions()
        {
            for (int index = 0; index < 4; index++)
            using (var host = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual })
            using (var container = new ChatToolWindow())
            using (var explorer = new TestExplorerWindow())
            {
                var designerSize = explorer.Size;
                var frame = index % 2 == 0 ? null : new ExplorerNativeFrameDouble();
                var native = new ExplorerNativeWindowDouble(host) { LinkedWindowFrame = frame };
                host.Controls.Add(container);
                var minimum = explorer.MinimumSize;
                var tinySizes = new[] { new Size(200, 100), new Size(Math.Max(1920, minimum.Width), 6),
                    new Size(minimum.Width - 1, minimum.Height + 120), new Size(minimum.Width + 120, minimum.Height - 1) };
                host.ClientSize = tinySizes[index];
                host.Show();
                container.Attach(explorer);
                Application.DoEvents();
                var owner = index % 2 == 0 ? (IWin32Window)host : new ExplorerMonitorFallbackOwner();
                var area = Screen.FromHandle(owner.Handle).WorkingArea;
                explorer.EnsureUsableNativePlacement(container, native, owner);
                Application.DoEvents();
                Assert.AreEqual(Math.Min(designerSize.Width, area.Width), host.Width);
                Assert.AreEqual(Math.Min(designerSize.Height, area.Height), host.Height);
                Assert.IsTrue(area.Contains(host.Bounds), "Recovered native bounds must remain on the owner's monitor.");
                Assert.IsFalse(explorer.TopLevel, "Recovery retains the native tool container.");
                Assert.IsTrue(explorer.Visible);
                if (frame != null) {
                    Assert.AreEqual(1, frame.LinkedWindows.Removed.Count);
                    Assert.AreSame(native, frame.LinkedWindows.Removed[0]);
                    Assert.IsNull(native.LinkedWindowFrame);
                }
                var recovered = host.Bounds;
                explorer.EnsureUsableNativePlacement(container, native, owner);
                Assert.AreEqual(recovered, host.Bounds, "Reopening a readable pane must not reset its placement.");
            }
        }

        [STATestMethod]
        public void NativePlacementPreservesReadableFloatingAndDockedUserBounds()
        {
            foreach (bool docked in new[] { false, true })
            using (var host = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual })
            using (var container = new ChatToolWindow())
            using (var explorer = new TestExplorerWindow())
            {
                host.ClientSize = new Size(explorer.MinimumSize.Width + 120, explorer.MinimumSize.Height + 80);
                var frame = docked ? new ExplorerNativeFrameDouble() : null;
                var native = new ExplorerNativeWindowDouble(host) { LinkedWindowFrame = frame };
                host.Controls.Add(container);
                host.Show();
                container.Attach(explorer);
                Application.DoEvents();
                var userBounds = host.Bounds;
                explorer.EnsureUsableNativePlacement(container, native, host);
                Assert.AreEqual(userBounds, host.Bounds);
                Assert.AreSame(frame, native.LinkedWindowFrame);
                if (frame != null) Assert.AreEqual(0, frame.LinkedWindows.Removed.Count);
                host.ClientSize = new Size(explorer.MinimumSize.Width + 160, explorer.MinimumSize.Height + 130);
                Application.DoEvents();
                userBounds = host.Bounds;
                explorer.EnsureUsableNativePlacement(container, native, host);
                Assert.AreEqual(userBounds, host.Bounds, "A later user resize is authoritative.");
            }
        }

        [STATestMethod]
        public void NativePlacementDoesNotMutateWhenTheOwnedSiteCannotBeRead()
        {
            using (var host = new Form { ClientSize = new Size(200, 100), ShowInTaskbar = false })
            using (var container = new ChatToolWindow())
            using (var explorer = new TestExplorerWindow())
            {
                var native = new ExplorerNativeWindowDouble(host);
                var bounds = host.Bounds;
                explorer.EnsureUsableNativePlacement(null, native, host);
                Assert.AreEqual(bounds, host.Bounds);
                explorer.EnsureUsableNativePlacement(container, native, host);
                Assert.AreEqual(bounds, host.Bounds, "No handle must not be treated as a collapsed native pane.");
                host.Controls.Add(container);
                host.Show();
                container.Attach(explorer);
                Application.DoEvents();
                container.ParentReader = unused => IntPtr.Zero;
                bounds = host.Bounds;
                explorer.EnsureUsableNativePlacement(container, native, host);
                Assert.AreEqual(bounds, host.Bounds, "An unverified parent must never be repositioned.");
            }
        }

        public sealed class ExplorerNativeWindowDouble
        {
            private readonly Form form;
            public ExplorerNativeWindowDouble(Form form) { this.form = form; }
            public ExplorerNativeFrameDouble LinkedWindowFrame { get; set; }
            public int Width { get => form.Width; set => form.Width = value; }
            public int Height { get => form.Height; set => form.Height = value; }
            public int Left { get => form.Left; set => form.Left = value; }
            public int Top { get => form.Top; set => form.Top = value; }
        }

        public sealed class ExplorerNativeFrameDouble
        {
            public ExplorerLinkedWindowsDouble LinkedWindows { get; } = new ExplorerLinkedWindowsDouble();
        }

        public sealed class ExplorerLinkedWindowsDouble
        {
            public List<ExplorerNativeWindowDouble> Removed { get; } = new List<ExplorerNativeWindowDouble>();
            public void Remove(ExplorerNativeWindowDouble window) { Removed.Add(window); window.LinkedWindowFrame = null; }
        }

        private sealed class ExplorerMonitorFallbackOwner : IWin32Window
        {
            public IntPtr Handle => IntPtr.Zero;
        }

        [STATestMethod]
        public void DiscoveryDoesNotExecuteAndFilteredCheckedTestsAreNotDispatched()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                Assert.AreEqual(0, service.Runs);
                var tree = Field<TreeView>(window, "testTree");
                tree.Nodes[0].Nodes[0].Nodes[0].Checked = true;
                Field<TextBox>(window, "search").Text = "Other";
                Assert.AreEqual(0, window.SelectedTests().Count);
                window.RunSelectedAsync().GetAwaiter().GetResult();
                Assert.AreEqual(0, service.Runs);
                window.RunScopeAsync().GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "second" }, service.Selected.Select(test => test.Id).ToArray());
            }
        }

        [STATestMethod]
        public void ResultsRetainMessageDurationAndRevisionWhileRefreshMarksStale()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.RunScopeAsync().GetAwaiter().GetResult();
                var tree = Field<TreeView>(window, "testTree");
                var handle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                string details = Field<TextBox>(window, "details").Text;
                StringAssert.Contains(details, "expected 4, actual 3");
                StringAssert.Contains(details, "12");
                StringAssert.Contains(details, "r1");
                Assert.IsTrue(Field<TextBox>(window, "humanReport").Text.Length > 0);
                Assert.IsTrue(Field<TextBox>(window, "compactReport").Text.Length > 0);
                service.Catalog.Project.Revision = "r2";
                window.RefreshProjects();
                Assert.IsFalse(Field<Button>(window, "rerunFailed").Enabled);
                int previousRuns = service.Runs;
                window.RerunFailedAsync().GetAwaiter().GetResult();
                Assert.AreEqual(previousRuns, service.Runs, "An explicit rerun must refuse stale results as well as disabling its button.");
            }
        }

        [STATestMethod]
        public void CategoryAndOutcomeFiltersSelectOnlyVisibleFailedTests()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.RunScopeAsync().GetAwaiter().GetResult();
                Field<TextBox>(window, "search").Text = "Arithmetic";
                Field<ComboBox>(window, "outcomeFilter").SelectedIndex = 1 + (int)VbaTestOutcome.Failed;
                Assert.AreEqual(1, window.VisibleTests().Count);
                window.RerunFailedAsync().GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "first" }, service.Selected.Select(test => test.Id).ToArray());
            }
        }

        [STATestMethod]
        public void RunningStateFreezesActionsAndCloseRequestsCooperativeStop()
        {
            var service = new ExplorerDouble { Pending = new TaskCompletionSource<VbaTestRun>() };
            var clock = Stopwatch.StartNew();
            var observations = new StringBuilder();
            Action<string> observe = phase =>
            {
                int workers, ports, minimumWorkers, minimumPorts;
                ThreadPool.GetAvailableThreads(out workers, out ports);
                ThreadPool.GetMinThreads(out minimumWorkers, out minimumPorts);
                string line = clock.ElapsedMilliseconds + " ms | " + phase + " | thread=" + Thread.CurrentThread.ManagedThreadId +
                    " apartment=" + Thread.CurrentThread.GetApartmentState() + " context=" +
                    (SynchronizationContext.Current?.GetType().FullName ?? "null") + " scheduler=" + TaskScheduler.Current.Id +
                    " workers=" + workers + " minWorkers=" + minimumWorkers + " ports=" + ports + " minPorts=" + minimumPorts;
                lock (observations) observations.AppendLine(line);
            };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                Task run = null;
                bool sentinelReceived = false;
                window.FormClosing += (_, args) => observe("FormClosing Cancel=" + args.Cancel);
                window.Disposed += (_, __) => observe("Form disposed");
                Field<Label>(window, "status").TextChanged += (_, __) => observe("Status=" + Field<Label>(window, "status").Text);
                Field<TextBox>(window, "humanReport").TextChanged += (_, __) => observe("Human report populated");
                Action<string> snapshot = phase => observe(phase + " | Pending=" + service.Pending.Task.Status + " Run=" + run?.Status +
                    " Running=" + Field<bool>(window, "running") + " CloseAfterRun=" + Field<bool>(window, "closeAfterRun") +
                    " Disposed=" + window.IsDisposed + " HandleCreated=" + window.IsHandleCreated + " InvokeRequired=" + window.InvokeRequired +
                    " Sentinel=" + sentinelReceived + " Queue=" + DescribeCallbackQueue(Field<Control>(window, "runContinuationDispatcher") ?? window));
                try
                {
                    observe("Owner before run hwnd=" + handle.ToInt64());
                    run = window.RunScopeAsync();
                    snapshot("Run suspended");
                    Assert.IsFalse(Field<Button>(window, "refresh").Enabled);
                    Assert.IsFalse(Field<Button>(window, "installSupport").Enabled);
                    Assert.IsFalse(Field<ComboBox>(window, "projectList").Enabled);
                    var args = new FormClosingEventArgs(CloseReason.UserClosing, false);
                    typeof(TestExplorerWindow).GetMethod("OnFormClosing", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(window, new object[] { args });
                    Assert.IsTrue(args.Cancel);
                    Assert.IsTrue(service.Cancellation.IsCancellationRequested);
                    Assert.AreEqual(1, service.Runs);
                    snapshot("Before Pending.SetResult");
                    service.Pending.SetResult(new VbaTestRun { Revision = "r1" });
                    snapshot("After Pending.SetResult");
                    PumpOwner(run);
                    snapshot("Pump completed");
                }
                catch (AssertFailedException)
                {
                    snapshot("Assertion failed at original deadline");
                    if (run != null && !run.IsCompleted && !window.IsDisposed)
                    {
                        // Publish only after failure: a second marshal message must not rescue the original five-second assertion.
                        // One diagnostic drain distinguishes a broken owner pump from an already queued but unwoken continuation.
                        (Field<Control>(window, "runContinuationDispatcher") ?? window).BeginInvoke(new Action(() => { sentinelReceived = true; snapshot("Post-failure owner sentinel received"); }));
                        snapshot("Post-failure owner sentinel posted");
                        Application.DoEvents();
                        snapshot("Post-failure single diagnostic drain");
                    }
                    throw;
                }
                finally
                {
                    snapshot("Final observation");
                    lock (observations) Console.WriteLine(observations.ToString());
                }
            }
        }

        [STATestMethod]
        public void CategoryGroupingNeverDispatchesTheSameTestTwice()
        {
            var service = new ExplorerDouble();
            service.Catalog.Modules[0].Tests[0].Categories = new[] { "Arithmetic", "Regression" };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                Field<ComboBox>(window, "grouping").SelectedIndex = 2;
                var tree = Field<TreeView>(window, "testTree");
                var handle = tree.Handle;
                tree.Nodes[0].Checked = true;
                Assert.AreEqual(2, window.SelectedTests().Count);
                window.RunSelectedAsync().GetAwaiter().GetResult();
                Assert.AreEqual(2, service.Selected.Count);
                CollectionAssert.AreEquivalent(new[] { "first", "second" }, service.Selected.Select(test => test.Id).ToArray());
                Field<ComboBox>(window, "grouping").SelectedIndex = 1;
                tree.SelectedNode = tree.Nodes[0].Nodes.Cast<TreeNode>().First(node => node.Nodes.Cast<TreeNode>().Any(test => ((VbaTestDescriptor)test.Tag).Id == "first"));
                Assert.AreEqual(1, window.ScopeTests().Count);
                Assert.AreEqual("first", window.ScopeTests()[0].Id);
            }
        }

        [STATestMethod]
        public void ModuleScopeSurvivesExecutionRefreshAndThemeWithoutIncludingAnotherModule()
        {
            var service = new ExplorerDouble();
            service.Catalog.Modules.Add(new VbaTestModule { Name = "Elsewhere", Tests = new List<VbaTestDescriptor>
                { new VbaTestDescriptor { Id = "third", Module = "Elsewhere", Procedure = "AnotherTest", Kind = "Sub" } } });
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                var tree = Field<TreeView>(window, "testTree");
                var treeHandle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => ((VbaTestModule)node.Tag).Name == "Tests");
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.AreEqual("Tests", ((VbaTestModule)tree.SelectedNode.Tag).Name);
                window.RefreshProjects(); RaiseThemeChanged();
                Assert.AreEqual("Tests", ((VbaTestModule)tree.SelectedNode.Tag).Name);
                window.RunScopeAsync().GetAwaiter().GetResult();
                CollectionAssert.AreEquivalent(new[] { "first", "second" }, service.Selected.Select(test => test.Id).ToArray());
                Assert.IsFalse(service.Selected.Any(test => test.Id == "third"));
            }
        }

        [STATestMethod]
        public void CategoryAndOutcomeScopesReturnAfterTransientAbsenceAndSurviveTheme()
        {
            foreach (int grouping in new[] { 1, 2 })
            using (var window = new TestExplorerWindow())
            {
                var service = new ExplorerDouble();
                window.Configure(service);
                var handle = window.Handle;
                window.RunScopeAsync().GetAwaiter().GetResult();
                Field<ComboBox>(window, "grouping").SelectedIndex = grouping;
                var tree = Field<TreeView>(window, "testTree");
                var treeHandle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => node.Nodes.Cast<TreeNode>().Any(test => ((VbaTestDescriptor)test.Tag).Id == "first"));
                string scope = tree.SelectedNode.Text;
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.AreEqual(scope, tree.SelectedNode.Text);
                window.RefreshProjects(); RaiseThemeChanged();
                Assert.AreEqual(scope, tree.SelectedNode.Text);
                CollectionAssert.AreEqual(new[] { "first" }, window.ScopeTests().Select(test => test.Id).ToArray());
                window.RunScopeAsync().GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "first" }, service.Selected.Select(test => test.Id).ToArray());
            }
        }

        [STATestMethod]
        public void AnOutcomeScopeThatDisappearsRemainsEmptyUntilTheUserSelectsAnotherScope()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                window.RunScopeAsync().GetAwaiter().GetResult();
                Field<ComboBox>(window, "grouping").SelectedIndex = 1;
                var tree = Field<TreeView>(window, "testTree");
                var treeHandle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => node.Text == UiText.Get("Failed"));
                service.PassAll = true;
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.IsNull(tree.SelectedNode);
                Assert.AreEqual(0, window.ScopeTests().Count);
                RaiseThemeChanged(); window.RefreshProjects();
                Assert.AreEqual(0, window.ScopeTests().Count);
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
                int runs = service.Runs;
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.AreEqual(runs, service.Runs);
                tree.SelectedNode = tree.Nodes[0];
                Assert.AreEqual(2, window.ScopeTests().Count);
            }
        }

        [STATestMethod]
        public void ASelectedTestAliasKeepsItsCategoryAfterRunAndTheme()
        {
            var service = new ExplorerDouble();
            service.Catalog.Modules[0].Tests[0].Categories = new[] { "Arithmetic", "Regression" };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                Field<ComboBox>(window, "grouping").SelectedIndex = 2;
                var tree = Field<TreeView>(window, "testTree");
                var treeHandle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => node.Text == "Regression").Nodes[0];
                window.RunSelectedAsync().GetAwaiter().GetResult();
                RaiseThemeChanged();
                Assert.AreEqual("first", ((VbaTestDescriptor)tree.SelectedNode.Tag).Id);
                Assert.AreEqual("Regression", tree.SelectedNode.Parent.Text);
            }
        }

        [STATestMethod]
        public void CategoryAliasChecksStaySynchronizedForTestGroupAndProjectChanges()
        {
            var service = new ExplorerDouble();
            service.Catalog.Modules[0].Tests[0].Categories = new[] { "Arithmetic", "Regression" };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                Field<ComboBox>(window, "grouping").SelectedIndex = 2;
                var tree = Field<TreeView>(window, "testTree");
                var treeHandle = tree.Handle;
                Func<TreeNode[]> aliases = () => tree.Nodes[0].Nodes.Cast<TreeNode>().SelectMany(node => node.Nodes.Cast<TreeNode>())
                    .Where(node => ((VbaTestDescriptor)node.Tag).Id == "first").ToArray();
                aliases()[0].Checked = true;
                Assert.IsTrue(aliases().All(node => node.Checked));
                RaiseThemeChanged();
                tree.Nodes[0].Checked = true;
                aliases()[1].Parent.Collapse();
                aliases()[0].Checked = false;
                Assert.IsTrue(aliases().All(node => !node.Checked));
                window.RunSelectedAsync().GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "second" }, service.Selected.Select(test => test.Id).ToArray());
                tree.Nodes[0].Checked = true; tree.Nodes[0].Checked = false;
                Assert.AreEqual(0, window.SelectedTests().Count);
                tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => node.Text == "Arithmetic").Checked = true;
                Assert.IsTrue(aliases().All(node => node.Checked));
                tree.Nodes[0].Nodes.Cast<TreeNode>().Single(node => node.Text == "Regression").Checked = false;
                Assert.IsTrue(aliases().All(node => !node.Checked));
                Assert.AreEqual(0, window.SelectedTests().Count);
            }
        }

        [STATestMethod]
        public void HumanAndJsonExportsRetainTheirSnapshotAcrossModalProjectAndFormatChanges()
        {
            foreach (bool compact in new[] { false, true })
            {
                string path = Path.Combine(Path.GetTempPath(), "VBAi-export-test-" + Guid.NewGuid().ToString("N") + (compact ? ".json" : ".txt"));
                try
                {
                    var service = new ExplorerDouble();
                    using (var window = new TestExplorerWindow())
                    {
                        window.Configure(service);
                        window.RunScopeAsync().GetAwaiter().GetResult();
                        var tabs = Field<TabControl>(window, "resultTabs");
                        tabs.SelectedTab = Field<TabPage>(window, compact ? "compactTab" : "humanTab");
                        string snapshot = Field<TextBox>(window, compact ? "compactReport" : "humanReport").Text;
                        window.ChooseReportExportPath = (owner, format) =>
                        {
                            Assert.AreSame(window, owner); Assert.AreEqual(compact, format);
                            service.Catalog.Project.Id = "other-project";
                            service.Catalog.Project.Name = "Other project";
                            service.Catalog.Project.Revision = "r2";
                            window.RefreshProjects();
                            tabs.SelectedTab = Field<TabPage>(window, compact ? "humanTab" : "compactTab");
                            Assert.AreEqual("", Field<TextBox>(window, "humanReport").Text);
                            Assert.AreEqual("", Field<TextBox>(window, "compactReport").Text);
                            return path;
                        };
                        typeof(TestExplorerWindow).GetMethod("ExportReport_Click", BindingFlags.Instance | BindingFlags.NonPublic)
                            .Invoke(window, new object[] { window, EventArgs.Empty });
                        Assert.AreEqual(snapshot, File.ReadAllText(path, Encoding.UTF8));
                        Assert.IsTrue(snapshot.Length > 0);
                        if (compact)
                        {
                            var report = (Dictionary<string, object>)new System.Web.Script.Serialization.JavaScriptSerializer().DeserializeObject(snapshot);
                            Assert.AreEqual("Disposable", report["project"]); Assert.AreEqual("r1", report["revision"]);
                            Assert.AreEqual(2, ((object[])report["tests"]).Length);
                        }
                        else { StringAssert.Contains(snapshot, "Disposable"); StringAssert.Contains(snapshot, "r1"); }
                    }
                }
                finally { if (File.Exists(path)) File.Delete(path); }
            }
        }

        [STATestMethod]
        public void PassedAndFailedGlyphsUseGlobalColorsAcrossThemesAndHighContrast()
        {
            using (var theme = new ThemeScope())
            using (var window = new TestExplorerWindow())
            {
                window.Configure(new ExplorerDouble());
                window.RunScopeAsync().GetAwaiter().GetResult();
                foreach (bool contrast in new[] { false, true })
                foreach (var choice in new[] { ThemeChoice.Light, ThemeChoice.Dark })
                {
                    UiTheme.HighContrast = () => contrast;
                    ThemeScope.SetChoice(choice);
                    RaiseThemeChanged();
                    var tests = Field<TreeView>(window, "testTree").Nodes[0].Nodes[0].Nodes.Cast<TreeNode>().ToArray();
                    StringAssert.StartsWith(tests[0].Text, "✗ ");
                    StringAssert.StartsWith(tests[1].Text, "✓ ");
                    StringAssert.Contains(tests[0].Text, UiText.Get(VbaTestOutcome.Failed.ToString()));
                    StringAssert.Contains(tests[1].Text, UiText.Get(VbaTestOutcome.Passed.ToString()));
                    Assert.AreEqual(UiTheme.Error, tests[0].ForeColor);
                    Assert.AreEqual(UiTheme.Success, tests[1].ForeColor);
                    if (contrast)
                    {
                        Assert.AreEqual(SystemColors.WindowText, tests[0].ForeColor);
                        Assert.AreEqual(SystemColors.WindowText, tests[1].ForeColor);
                    }
                }
            }
        }

        [STATestMethod]
        public void RenderOwnedSimulatedWindowInBothGlobalThemes()
        {
            using (var theme = new ThemeScope())
            using (var window = new TestExplorerWindow())
            {
                var service = new ExplorerDouble();
                service.Catalog.Project.HostPath = @"C:\Disposable\VBAi-test-fixture.xlsm";
                window.Configure(service);
                window.Show();
                Application.DoEvents();
                window.RunScopeAsync().GetAwaiter().GetResult();
                var tree = Field<TreeView>(window, "testTree");
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-explorer-ui");
                Directory.CreateDirectory(directory);
                foreach (var choice in new[] { ThemeChoice.Light, ThemeChoice.Dark })
                {
                    ThemeScope.SetChoice(choice);
                    RaiseThemeChanged();
                    window.Refresh();
                    Application.DoEvents();
                    Assert.AreEqual(UiTheme.Background, window.BackColor);
                    Assert.IsTrue(Field<TextBox>(window, "details").ClientSize.Width > 200);
                    var actions = Field<FlowLayoutPanel>(window, "commands");
                    foreach (Control control in actions.Controls) Assert.IsTrue(actions.ClientRectangle.Contains(control.Bounds), "Clipped action: " + control.Text);
                    using (var bitmap = new Bitmap(window.Width, window.Height))
                    {
                        window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                        bitmap.Save(Path.Combine(directory, "explorer-" + choice.ToString().ToLowerInvariant() + "-" + window.DeviceDpi + "dpi.png"), ImageFormat.Png);
                    }
                }
                window.Close();
            }
        }

        [STATestMethod]
        public void FreshnessPollMarksExternalEditsStaleAndClearsClosedProjectWithoutExecution()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.Show();
                window.RunScopeAsync().GetAwaiter().GetResult();
                service.Catalog.Project.Revision = "external-change";
                window.CheckFreshness();
                var tests = Field<TreeView>(window, "testTree").Nodes[0].Nodes[0].Nodes.Cast<TreeNode>().ToArray();
                StringAssert.Contains(tests[1].Text, UiText.Get("Previous revision"));
                Assert.AreEqual(UiTheme.Muted, tests[1].ForeColor);
                Assert.AreEqual(1, service.Runs);
                service.DiscoveryError = new InvalidOperationException("Project closed.");
                window.CheckFreshness();
                Assert.AreEqual(0, window.VisibleTests().Count);
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
                Assert.IsFalse(Field<Button>(window, "runScopeCoverage").Enabled);
                Assert.AreEqual(1, service.Runs);
                window.Close();
            }
        }

        [STATestMethod]
        public void CoverageRequiresReviewedSelectionAndShowsProcedureMetricAndPartialState()
        {
            var service = new CoverageDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                Field<TreeView>(window, "testTree").Nodes[0].Nodes[0].Nodes[0].Checked = true;
                string review = null;
                window.ConfirmCoverage = (owner, text) => { review = text; return false; };
                window.RunSelectedCoverageAsync().GetAwaiter().GetResult();
                Assert.AreEqual(0, service.CoverageRuns);
                StringAssert.Contains(review, "Tests.Adds");
                StringAssert.Contains(review, @"%LOCALAPPDATA%\VBAi\CoverageRuns");
                StringAssert.Contains(review, "r1");
                window.ConfirmCoverage = (owner, text) => true;
                window.RunSelectedCoverageAsync().GetAwaiter().GetResult();
                Assert.AreEqual(1, service.CoverageRuns);
                Assert.AreEqual(1, service.Selected.Count);
                Assert.AreEqual("first", service.Selected[0].Id);
                StringAssert.Contains(Field<Label>(window, "coverage").Text, "50%");
                StringAssert.Contains(Field<Label>(window, "coverage").Text, "(1/2)");
                StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Partial measurement"));
                StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Statement and branch coverage are not measured."));
            }
        }

        [STATestMethod]
        public void CoverageUnavailableAndUnknownAreExplicitWithoutFalsePercentage()
        {
            var service = new CoverageDouble { Report = new VbaCoverageReport { Available = false, Complete = false, Revision = "r1" } };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.ConfirmCoverage = (owner, text) => true;
                window.RunScopeCoverageAsync().GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Unavailable"));
                Assert.IsFalse(Field<Label>(window, "coverage").Text.Contains("0%"));
                service.Report = new VbaCoverageReport { Available = true, Complete = false, Revision = "r1" };
                window.RunScopeCoverageAsync().GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Unknown"));
                Assert.IsFalse(Field<Label>(window, "coverage").Text.Contains("0%"));
            }
        }

        [STATestMethod]
        public void ActiveCoverageRetainsProgressAndClosingStopsOnlyAfterPendingResult()
        {
            var service = new CoverageDouble { Pending = new TaskCompletionSource<VbaTestRun>() };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.Show();
                window.ConfirmCoverage = (owner, text) => true;
                Task run = window.RunScopeCoverageAsync();
                int discoveries = service.DiscoveryCalls;
                window.CheckFreshness();
                Assert.AreEqual(discoveries, service.DiscoveryCalls);
                Assert.IsFalse(Field<Button>(window, "runScopeCoverage").Enabled);
                Assert.IsFalse(Field<Button>(window, "refresh").Enabled);
                var args = new FormClosingEventArgs(CloseReason.UserClosing, false);
                typeof(TestExplorerWindow).GetMethod("OnFormClosing", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(window, new object[] { args });
                Assert.IsTrue(args.Cancel);
                Assert.IsTrue(service.Cancellation.IsCancellationRequested);
                Assert.IsFalse(window.IsDisposed);
                var result = new VbaTestResult { Test = service.Catalog.Modules[0].Tests[0], Outcome = VbaTestOutcome.Passed, Message = "verified late completion" };
                service.SavedProgress(result);
                StringAssert.StartsWith(Field<TreeView>(window, "testTree").Nodes[0].Nodes[0].Nodes[0].Text, "✓ ");
                service.Pending.SetResult(new VbaTestRun { Revision = "r1", Results = new List<VbaTestResult> { result } });
                PumpOwner(run);
                Assert.IsTrue(window.IsDisposed);
                Assert.AreEqual(1, service.CoverageRuns);
                Assert.AreEqual(1, service.Runs);
            }
        }

        [STATestMethod]
        public void CoverageCompletionAfterNativeContextLossUpdatesReportsAndClosesOnOwner()
            => CheckCoverageContinuationAfterContextLoss(false);

        [STATestMethod]
        public void CoverageExceptionAfterNativeContextLossUpdatesStatusAndClosesOnOwner()
            => CheckCoverageContinuationAfterContextLoss(true);

        [STATestMethod]
        public void ForcedDisposalKeepsPendingSuccessOnOwnerWithoutTouchingDisposedUiOrService()
            => CheckForcedDisposal(false);

        [STATestMethod]
        public void ForcedDisposalKeepsPendingExceptionOnOwnerWithoutUnhandledPublication()
            => CheckForcedDisposal(true);

        [STATestMethod]
        public void EarlyDisposalTestAssertionStillSettlesPendingAndReleasesOwnerHandle()
        {
            var previous = SynchronizationContext.Current;
            int owner = Thread.CurrentThread.ManagedThreadId;
            try
            {
                foreach (bool coverage in new[] { false, true })
                using (var window = new TestExplorerWindow())
                {
                    var service = new DisposalDouble { ClearContext = true, Pending = new TaskCompletionSource<VbaTestRun>() };
                    window.Configure(service);
                    window.ConfirmCoverage = (_, __) => true;
                    Task run = null;
                    int releasedOn = 0;
                    var error = Assert.ThrowsException<AssertFailedException>(() =>
                    {
                        try
                        {
                            run = coverage ? window.RunScopeCoverageAsync() : window.RunScopeAsync();
                            Field<Control>(window, "runContinuationDispatcher").Disposed += (_, __) => releasedOn = Thread.CurrentThread.ManagedThreadId;
                            Assert.Fail("Deliberate assertion before pending completion.");
                        }
                        finally { SettleDisposalTest(window, service, run); }
                    });
                    StringAssert.Contains(error.Message, "Deliberate assertion before pending completion.");
                    Assert.IsTrue(service.Pending.Task.IsCanceled);
                    Assert.IsTrue(run.IsCompleted); Assert.AreEqual(owner, releasedOn);
                    Assert.IsNull(Field<Control>(window, "runContinuationDispatcher"));
                    Assert.AreEqual(0, service.WrongOwnerCalls);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        private static void CheckForcedDisposal(bool fail)
        {
            var previous = SynchronizationContext.Current;
            int owner = Thread.CurrentThread.ManagedThreadId;
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            var unhandled = new List<Exception>();
            ThreadExceptionEventHandler uiError = (_, args) => { lock (unhandled) unhandled.Add(args.Exception); };
            UnhandledExceptionEventHandler processError = (_, args) => { lock (unhandled) unhandled.Add(args.ExceptionObject as Exception); };
            Application.ThreadException += uiError;
            AppDomain.CurrentDomain.UnhandledException += processError;
            try
            {
                foreach (bool clearContext in new[] { false, true })
                foreach (bool coverage in new[] { false, true })
                using (var window = new TestExplorerWindow())
                {
                    SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                    var service = new DisposalDouble { ClearContext = clearContext, Pending = new TaskCompletionSource<VbaTestRun>() };
                    window.Configure(service);
                    window.ConfirmCoverage = (_, __) => true;
                    Task run = null;
                    try
                    {
                        run = coverage ? window.RunScopeCoverageAsync() : window.RunScopeAsync();
                        Assert.AreEqual(owner, service.DispatchThread);
                        if (clearContext) Assert.IsNull(service.ContextBeforeReturn, "The fake native boundary must clear context before returning its pending task.");
                        else Assert.IsInstanceOfType(service.ContextBeforeReturn, typeof(WindowsFormsSynchronizationContext));
                        var dispatcher = Field<Control>(window, "runContinuationDispatcher");
                        Assert.IsNotNull(dispatcher); Assert.IsTrue(dispatcher.IsHandleCreated);
                        Assert.IsFalse(run.IsCompleted);
                        int releasedOn = 0;
                        ApartmentState releasedApartment = ApartmentState.Unknown;
                        dispatcher.Disposed += (_, __) => { releasedOn = Thread.CurrentThread.ManagedThreadId; releasedApartment = Thread.CurrentThread.GetApartmentState(); };
                        bool disposedUi = false;
                        int lateUiChanges = 0;
                        EventHandler uiChanged = (_, __) => { if (disposedUi) lateUiChanges++; };
                        Field<Label>(window, "status").TextChanged += uiChanged;
                        Field<Label>(window, "coverage").TextChanged += uiChanged;
                        Field<TextBox>(window, "humanReport").TextChanged += uiChanged;
                        Field<TextBox>(window, "compactReport").TextChanged += uiChanged;
                        var result = new VbaTestResult { Test = service.Selected[0], Outcome = VbaTestOutcome.Passed, Message = "Late native result" };
                        Exception progressError = null;
                        var progress = new Thread(() => { try { service.SavedProgress(result); } catch (Exception error) { progressError = error; } });
                        progress.Start(); progress.Join();
                        Assert.IsNull(progressError);
                        Assert.IsTrue(CallbackCount(dispatcher) > 0, "Worker progress must be queued before forced disposal.");
                        int readsBeforeDisposal = service.AvailabilityCalls;
                        window.Dispose();
                        disposedUi = true;
                        service.Disposed = true; // Models AddIn teardown disposing the service while its native result is still pending.
                        Assert.IsTrue(window.IsDisposed); Assert.IsTrue(service.Cancellation.IsCancellationRequested);
                        Assert.IsFalse(dispatcher.IsDisposed, "Only the run continuation handle must survive forced UI disposal.");
                        Assert.IsFalse(run.IsCompleted);
                        Exception completionError = null;
                        int completionThread = 0;
                        var completion = new Thread(() =>
                        {
                            try
                            {
                                completionThread = Thread.CurrentThread.ManagedThreadId;
                                service.SavedProgress(result); // A late progress callback must be ignored after disposal.
                                if (fail) service.Pending.SetException(new InvalidOperationException("Late native failure"));
                                else service.Pending.SetResult(new VbaTestRun { Revision = "r1", Results = new List<VbaTestResult> { result } });
                            }
                            catch (Exception error) { completionError = error; }
                        });
                        // A plain worker with no context and a synchronous TCS publishes BeginInvoke before Join returns.
                        // Drain the already posted WinForms queue explicitly on the owner; no polling, timer or deadline is used.
                        completion.Start(); completion.Join();
                        Assert.AreNotEqual(owner, completionThread); Assert.IsNull(completionError);
                        Assert.IsTrue(CallbackCount(dispatcher) >= 2, "Both prior progress and the native continuation must be posted to the surviving owner handle.");
                        typeof(Control).GetMethod("InvokeMarshaledCallbacks", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(dispatcher, null);
                        Assert.IsTrue(run.IsCompleted, "The deterministic owner drain did not settle the disposed window's run.");
                        run.GetAwaiter().GetResult();
                        Assert.IsTrue(dispatcher.IsDisposed); Assert.AreEqual(owner, releasedOn);
                        Assert.AreEqual(ApartmentState.STA, releasedApartment);
                        Assert.AreEqual(0, lateUiChanges, "Completion changed a disposed UI control.");
                        Assert.IsNull(Field<Control>(window, "runContinuationDispatcher"));
                        Assert.IsFalse(Field<bool>(window, "running"));
                        Assert.IsNull(Field<CancellationTokenSource>(window, "cancellation"));
                        Assert.AreEqual(0, Field<Dictionary<string, VbaTestResult>>(window, "results").Count, "A disposed UI accepted a queued or late native result.");
                        Assert.AreEqual(readsBeforeDisposal, service.AvailabilityCalls, "Completion queried a service that AddIn teardown had already disposed.");
                        Assert.AreEqual(0, service.WrongOwnerCalls);
                        lock (unhandled) Assert.AreEqual(0, unhandled.Count, "An owner/worker continuation escaped as an unhandled exception.");
                    }
                    finally { SettleDisposalTest(window, service, run); }
                }
            }
            finally
            {
                AppDomain.CurrentDomain.UnhandledException -= processError;
                Application.ThreadException -= uiError;
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        private static void SettleDisposalTest(TestExplorerWindow window, DisposalDouble service, Task run)
        {
            // An assertion can fail while the independent dispatcher still owns a pending result.
            // Settle it from a context-free worker and drain on the owner even on that failure path.
            window.Dispose();
            service.Disposed = true;
            if (!service.Pending.Task.IsCompleted)
            {
                Exception publicationError = null;
                var worker = new Thread(() => { try { service.Pending.TrySetCanceled(); } catch (Exception error) { publicationError = error; } });
                worker.Start(); worker.Join();
                Assert.IsNull(publicationError, "Cleanup could not publish the pending fake result.");
            }
            var dispatcher = Field<Control>(window, "runContinuationDispatcher");
            if (dispatcher != null && !dispatcher.IsDisposed)
                typeof(Control).GetMethod("InvokeMarshaledCallbacks", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(dispatcher, null);
            if (run != null)
            {
                Assert.IsTrue(run.IsCompleted, "Cleanup did not settle the pending run on its owner.");
                run.GetAwaiter().GetResult();
            }
            Assert.IsNull(Field<Control>(window, "runContinuationDispatcher"), "Cleanup left an independent dispatcher alive.");
        }

        private static int CallbackCount(Control control)
        {
            var queue = (System.Collections.ICollection)typeof(Control).GetField("threadCallbackList", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(control);
            if (queue == null) return 0;
            lock (queue.SyncRoot) return queue.Count;
        }

        private static void CheckCoverageContinuationAfterContextLoss(bool fail)
        {
            var previous = SynchronizationContext.Current;
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            try
            {
                var service = new ContextClearingCoverageDouble();
                using (var window = new TestExplorerWindow())
                {
                    window.Configure(service); window.Show();
                    window.ConfirmCoverage = (owner, text) => true;
                    var human = Field<TextBox>(window, "humanReport");
                    var compact = Field<TextBox>(window, "compactReport");
                    var status = Field<Label>(window, "status");
                    var coverage = Field<Label>(window, "coverage");
                    var updateThreads = new List<int>();
                    EventHandler updated = (_, __) => { lock (updateThreads) updateThreads.Add(Thread.CurrentThread.ManagedThreadId); };
                    human.TextChanged += updated; compact.TextChanged += updated; status.TextChanged += updated; coverage.TextChanged += updated;
                    string closedHuman = null, closedCompact = null, closedStatus = null, closedCoverage = null;
                    int finalCloseThread = 0;
                    window.FormClosing += (_, args) =>
                    {
                        if (args.Cancel) return;
                        finalCloseThread = Thread.CurrentThread.ManagedThreadId;
                        closedHuman = human.Text; closedCompact = compact.Text; closedStatus = status.Text; closedCoverage = coverage.Text;
                    };
                    int availabilityBefore = service.AvailabilityCalls;
                    var running = window.RunScopeCoverageAsync();
                    Assert.IsTrue(service.ContextWasCleared, "The fake native boundary must remove the context synchronously before returning its task.");
                    Assert.IsFalse(running.IsCompleted);
                    Assert.IsFalse(Field<Button>(window, "runScopeCoverage").Enabled);
                    window.Close();
                    Assert.IsFalse(window.IsDisposed); Assert.IsTrue(service.Cancellation.IsCancellationRequested);
                    int completionThread = Task.Run(() =>
                    {
                        if (fail) service.Completion.SetException(new InvalidOperationException("Asynchronous native coverage failure."));
                        else
                        {
                            var result = new VbaTestResult { Test = service.Selected[0], Outcome = VbaTestOutcome.Passed, Message = "Verified asynchronous coverage completion." };
                            service.Completion.SetResult(new VbaTestRun { Id = "async-run", Revision = "r1", Results = new List<VbaTestResult> { result },
                                Coverage = new VbaCoverageReport { Available = true, Complete = false, DenominatorKnown = true, Revision = "r1", Hit = 1, Eligible = 2, Percent = 50 } });
                        }
                        return Thread.CurrentThread.ManagedThreadId;
                    }).GetAwaiter().GetResult();
                    Assert.AreNotEqual(ownerThread, completionThread);
                    PumpOwner(running);
                    Assert.IsTrue(window.IsDisposed); Assert.AreEqual(ownerThread, finalCloseThread);
                    Assert.AreEqual(0, service.WrongOwnerCalls);
                    Assert.IsTrue(service.AvailabilityCalls > availabilityBefore, "The final capability refresh must exercise the strict owning-thread service.");
                    lock (updateThreads) { Assert.IsTrue(updateThreads.Count > 0); Assert.IsTrue(updateThreads.All(thread => thread == ownerThread)); }
                    if (fail)
                    {
                        StringAssert.Contains(closedStatus, "Asynchronous native coverage failure.");
                        Assert.AreEqual("", closedHuman); Assert.AreEqual("", closedCompact);
                    }
                    else
                    {
                        StringAssert.Contains(closedHuman, "Verified asynchronous coverage completion.");
                        StringAssert.Contains(closedCompact, "async-run");
                        StringAssert.Contains(closedCoverage, "50%");
                        StringAssert.Contains(closedStatus, UiText.Get("Test run completed."));
                    }
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [STATestMethod]
        public void ExactProjectSelectionRejectsMissingOrAnotherProjectDuringRun()
        {
            var service = new ExplorerDouble { Pending = new TaskCompletionSource<VbaTestRun>() };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var handle = window.Handle;
                window.SelectProject("project");
                Assert.ThrowsException<InvalidOperationException>(() => window.SelectProject("missing"));
                var run = window.RunScopeAsync();
                window.SelectProject("project");
                Assert.ThrowsException<InvalidOperationException>(() => window.SelectProject("another"));
                service.Pending.SetResult(new VbaTestRun { Revision = "r1" });
                PumpOwner(run);
            }
        }

        [STATestMethod]
        public void UnavailableCapabilityBlocksDispatchButKeepsDiscoveryUsable()
        {
            var service = new ExplorerDouble { Unavailable = "Host cannot return verified results." };
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                Assert.AreEqual(2, window.VisibleTests().Count);
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.AreEqual(0, service.Runs);
                StringAssert.Contains(Field<Label>(window, "status").Text, service.Unavailable);
            }
        }

        [STATestMethod]
        public void CommandHandlersRespectAvailabilityAndDisplayNavigationOrInstallationFailures()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                Invoke(window, "Refresh_Click", window, EventArgs.Empty);
                Invoke(window, "InstallSupport_Click", window, EventArgs.Empty);
                Invoke(window, "Source_Click", window, EventArgs.Empty);
                Assert.ThrowsException<ArgumentNullException>(() => window.Configure(null));
                Assert.ThrowsException<ArgumentException>(() => window.SelectProject(""));
                window.Configure(service);
                Invoke(window, "Source_Click", window, EventArgs.Empty);
                var tree = Field<TreeView>(window, "testTree");
                var handle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                Invoke(window, "Source_Click", window, EventArgs.Empty);
                Assert.AreEqual(1, service.Navigations);
                service.NavigationError = new InvalidOperationException("Navigation failed");
                Invoke(window, "TestTree_NodeMouseDoubleClick", tree, new TreeNodeMouseClickEventArgs(tree.SelectedNode, MouseButtons.Left, 2, 0, 0));
                StringAssert.Contains(Field<Label>(window, "status").Text, "Navigation failed");
                Invoke(window, "InstallSupport_Click", window, EventArgs.Empty);
                Assert.AreEqual(1, service.Installations);
                service.InstallError = new InvalidOperationException("Support rejected");
                Invoke(window, "InstallSupport_Click", window, EventArgs.Empty);
                StringAssert.Contains(Field<Label>(window, "status").Text, "Support rejected");
                Invoke(window, "Refresh_Click", window, EventArgs.Empty);
                service.ReadError = new InvalidOperationException("Projects closed");
                window.RefreshProjects();
                Assert.AreEqual(0, window.VisibleTests().Count);
                StringAssert.Contains(Field<Label>(window, "status").Text, "Projects closed");
                service.ReadError = null;
                service.DiscoveryError = new InvalidOperationException("Discovery failed");
                window.RefreshProjects();
                Assert.IsNull(Field<object>(window, "catalog"));
                StringAssert.Contains(Field<Label>(window, "status").Text, "Discovery failed");
            }
        }

        [STATestMethod]
        public void ClipboardAndExportEffectsUseCurrentReportAndDoNotEscapeAsUiErrors()
        {
            using (var window = new TestExplorerWindow())
            {
                Invoke(window, "CopyReport_Click", window, EventArgs.Empty);
                Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                window.Configure(new ExplorerDouble());
                window.RunScopeAsync().GetAwaiter().GetResult();
                var tabHandle = Field<TabControl>(window, "resultTabs").Handle;
                Field<TabControl>(window, "resultTabs").SelectedTab = Field<TabPage>(window, "humanTab");
                Assert.IsTrue(Field<Button>(window, "copyReport").Enabled);
                string copied = null;
                window.WriteReportClipboard = value => copied = value;
                Invoke(window, "CopyReport_Click", window, EventArgs.Empty);
                Assert.AreEqual(Field<TextBox>(window, "humanReport").Text, copied);
                window.WriteReportClipboard = _ => { throw new InvalidOperationException("Clipboard busy"); };
                Invoke(window, "CopyReport_Click", window, EventArgs.Empty);
                StringAssert.Contains(Field<Label>(window, "status").Text, "Clipboard busy");
                window.ChooseReportExportPath = (_, __) => null;
                Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                window.ChooseReportExportPath = (_, __) => { throw new IOException("Export denied"); };
                Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                StringAssert.Contains(Field<Label>(window, "status").Text, "Export denied");
                Field<TabControl>(window, "resultTabs").SelectedIndex = -1;
                Assert.AreEqual("", Invoke(window, "ReportText"));
                Invoke(window, "CopyReport_Click", window, EventArgs.Empty);
                window.ChooseReportExportPath = (_, __) => { window.Dispose(); throw new IOException("Disposed during export"); };
                Field<TabControl>(window, "resultTabs").SelectedTab = Field<TabPage>(window, "humanTab");
                Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                Assert.IsTrue(window.IsDisposed);
            }
        }

        [STATestMethod]
        public void ButtonsDispatchTheirSingleAndBatchSelectionsAndIgnoreActionsDuringAnActiveRun()
        {
            var service = new CoverageDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                window.ConfirmCoverage = (_, __) => true;
                var tree = Field<TreeView>(window, "testTree");
                var handle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                Invoke(window, "RunSelected_Click", window, EventArgs.Empty);
                Assert.AreEqual(1, service.Runs);
                Invoke(window, "RunScope_Click", window, EventArgs.Empty);
                Assert.AreEqual(2, service.Runs);
                Invoke(window, "RerunFailed_Click", window, EventArgs.Empty);
                Assert.AreEqual(3, service.Runs);
                Invoke(window, "RunSelectedCoverage_Click", window, EventArgs.Empty);
                Invoke(window, "RunScopeCoverage_Click", window, EventArgs.Empty);
                Assert.AreEqual(2, service.CoverageRuns);
                service.Pending = new TaskCompletionSource<VbaTestRun>();
                var run = window.RunScopeAsync();
                Assert.ThrowsException<InvalidOperationException>(() => window.Configure(service));
                int reads = service.DiscoveryCalls;
                window.RefreshProjects();
                Invoke(window, "Source_Click", window, EventArgs.Empty);
                Invoke(window, "InstallSupport_Click", window, EventArgs.Empty);
                Invoke(window, "CopyReport_Click", window, EventArgs.Empty);
                Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                Invoke(window, "Filter_Changed", window, EventArgs.Empty);
                window.RunScopeCoverageAsync().GetAwaiter().GetResult();
                window.RunScopeAsync().GetAwaiter().GetResult();
                Assert.AreEqual(reads, service.DiscoveryCalls);
                Invoke(window, "Stop_Click", window, EventArgs.Empty);
                Assert.IsTrue(service.Cancellation.IsCancellationRequested);
                service.Pending.SetCanceled();
                PumpOwner(run);
                StringAssert.Contains(Field<Label>(window, "status").Text, UiText.Get("Test run stopped."));
                Invoke(window, "Stop_Click", window, EventArgs.Empty);
            }
        }

        [STATestMethod]
        public void FreshnessCapabilityChangesAndEmptyProjectListsNeverTriggerExecution()
        {
            var service = new CoverageDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service); window.Show();
                window.CheckFreshness();
                service.Unavailable = "Execution trust changed";
                window.CheckFreshness();
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
                StringAssert.Contains(Field<Label>(window, "status").Text, service.Unavailable);
                service.Unavailable = null;
                service.CoverageUnavailable = "Coverage trust changed";
                Invoke(window, "FreshnessTimer_Tick", window, EventArgs.Empty);
                Assert.IsFalse(Field<Button>(window, "runScopeCoverage").Enabled);
                service.NoProjects = true;
                window.RefreshProjects();
                Assert.AreEqual(-1, Field<ComboBox>(window, "projectList").SelectedIndex);
                window.CheckFreshness();
                Assert.AreEqual(0, window.ScopeTests().Count);
                Assert.AreEqual(0, service.Runs);
            }
        }

        [STATestMethod]
        public void CoverageReviewFreezesRefreshAndRejectsChangedOrDisposedCatalogs()
        {
            foreach (bool dispose in new[] { false, true })
            {
                var service = new CoverageDouble();
                using (var window = new TestExplorerWindow())
                {
                    window.Configure(service);
                    window.ConfirmCoverage = (_, __) =>
                    {
                        int reads = service.DiscoveryCalls;
                        window.RefreshProjects(); window.SelectProject("project");
                        Assert.ThrowsException<InvalidOperationException>(() => window.SelectProject("other"));
                        var closing = new FormClosingEventArgs(CloseReason.UserClosing, false);
                        Invoke(window, "OnFormClosing", closing);
                        Assert.IsTrue(closing.Cancel);
                        Assert.AreEqual(reads, service.DiscoveryCalls);
                        if (dispose) window.Dispose();
                        else SetField(window, "catalog", new VbaTestCatalog { Project = service.Catalog.Project });
                        return true;
                    };
                    window.RunScopeCoverageAsync().GetAwaiter().GetResult();
                    Assert.AreEqual(0, service.CoverageRuns);
                    if (!dispose) StringAssert.Contains(Field<Label>(window, "status").Text, UiText.Get("The reviewed project changed. Review coverage again."));
                }
            }
        }

        [STATestMethod]
        public void BlockedIgnoredAndErrorTestsRetainDiagnosticsAndCoverageCompleteness()
        {
            var service = new ExplorerDouble();
            var first = service.Catalog.Modules[0].Tests[0];
            var second = service.Catalog.Modules[0].Tests[1];
            first.Diagnostic = "Invalid signature"; first.Categories = null;
            second.IgnoreReason = "Manual only";
            service.Catalog.Modules.Add(new VbaTestModule { Name = "Blocked", Diagnostic = "Private module" });
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                var tree = Field<TreeView>(window, "testTree"); var handle = tree.Handle;
                StringAssert.Contains(tree.Nodes[0].Nodes[0].Nodes[0].Text, UiText.Get("Blocked"));
                Assert.AreEqual(UiTheme.Foreground, tree.Nodes[0].Nodes[0].Nodes[1].ForeColor);
                tree.SelectedNode = tree.Nodes[0].Nodes[1];
                Assert.AreEqual("Private module", Field<TextBox>(window, "details").Text);
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                Field<TextBox>(window, "search").Text = "Adds";
                Assert.AreEqual(1, window.VisibleTests().Count);
                Invoke(window, "AcceptResult", new VbaTestResult { Test = first, Outcome = VbaTestOutcome.Error, ErrorNumber = 5, Message = "Runtime failure" }, "old");
                StringAssert.Contains(Field<TextBox>(window, "details").Text, UiText.Get("VBA error") + ": 5");
                StringAssert.Contains(Field<TextBox>(window, "details").Text, UiText.Get("Result belongs to a previous source revision."));
                Invoke(window, "AcceptResult", null, "r1");
                Invoke(window, "AcceptResult", new VbaTestResult(), "r1");
                foreach (var report in new[] {
                    new VbaCoverageReport { Available = true, Complete = true, Revision = "old", Percent = 100, Eligible = 1, Hit = 1, Diagnostics = new List<string> { "Coverage details" } },
                    new VbaCoverageReport { Available = true, Revision = "r1", Percent = 100, Eligible = null, Hit = 1 },
                    new VbaCoverageReport { Available = true, Revision = "r1", Percent = 100, Eligible = 1, Hit = null } })
                {
                    SetField(window, "lastCoverage", report); Invoke(window, "UpdateCoverage");
                    if (report.Complete) StringAssert.Contains(Field<Label>(window, "coverage").Text, "Coverage details");
                    else StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Unknown"));
                }
                SetField(window, "catalog", null); Invoke(window, "UpdateCoverage");
                StringAssert.Contains(Field<Label>(window, "coverage").Text, UiText.Get("Previous revision"));
                SetField(window, "lastCoverage", null);
            }
        }

        [STATestMethod]
        public void RunFailuresAndUnknownOutcomesAreDisplayedWithoutRetry()
        {
            foreach (bool unknown in new[] { false, true })
            using (var window = new TestExplorerWindow())
            {
                var service = new ExplorerDouble { Pending = new TaskCompletionSource<VbaTestRun>() };
                window.Configure(service);
                service.Pending.SetResult(new VbaTestRun { Revision = "r1", Error = unknown ? null : "Run guard failed", OutcomeUnknown = unknown });
                window.RunScopeAsync().GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "status").Text, unknown ? UiText.Get("Test outcome is unknown. No automatic retry was performed.") : "Run guard failed");
                Assert.AreEqual(1, service.Runs);
            }
            using (var window = new TestExplorerWindow())
            {
                var service = new ExplorerDouble { Pending = new TaskCompletionSource<VbaTestRun>() };
                window.Configure(service);
                service.Pending.SetException(new InvalidOperationException("Native call failed"));
                service.AvailabilityError = new InvalidOperationException("Host disconnected");
                window.RunScopeAsync().GetAwaiter().GetResult();
                StringAssert.Contains(Field<Label>(window, "status").Text, "Native call failed");
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
            }
        }

        [STATestMethod]
        public void DispatcherCreationFailuresDisposeOwnedControlAndRejectWrongThreadUse()
        {
            using (var window = new TestExplorerWindow())
            {
                var control = new BrokenDispatcherControl();
                window.CreateRunDispatcher = () => control;
                var error = Assert.ThrowsException<TargetInvocationException>(() => Invoke(window, "InitializeRunContinuationDispatcher"));
                Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
                Assert.IsTrue(control.IsDisposed);
                var missing = typeof(TestExplorerWindow).GetMethod("AwaitOwner", BindingFlags.Instance | BindingFlags.NonPublic).MakeGenericMethod(typeof(VbaTestRun));
                Assert.IsInstanceOfType(Assert.ThrowsException<TargetInvocationException>(() => missing.Invoke(window, new object[] { Task.FromResult(new VbaTestRun()) })).InnerException, typeof(InvalidOperationException));
                Exception wrongThread = null;
                var worker = new Thread(() => { try { Invoke(window, "RequireRunOwner"); } catch (TargetInvocationException e) { wrongThread = e.InnerException; } });
                worker.Start(); worker.Join();
                Assert.IsInstanceOfType(wrongThread, typeof(InvalidOperationException));
            }
        }

        [STATestMethod]
        public void SaveDialogPreservesHumanOrJsonDefaultsAndReturnsCancellationWithoutWriting()
        {
            using (var owner = new Form())
            foreach (bool compact in new[] { false, true })
            foreach (var choice in new[] { DialogResult.Cancel, DialogResult.OK })
            {
                string value = TestExplorerWindow.SelectReportExportPath(owner, compact, (dialog, suppliedOwner) =>
                {
                    Assert.AreSame(owner, suppliedOwner);
                    Assert.AreEqual(compact ? "vba-test-results.json" : "vba-test-results.txt", dialog.FileName);
                    StringAssert.StartsWith(dialog.Filter, compact ? "JSON" : "Text");
                    dialog.FileName = "reviewed-result";
                    return choice;
                });
                Assert.AreEqual(choice == DialogResult.OK ? "reviewed-result" : null, value);
            }
        }

        [STATestMethod]
        public void NativeSaveDialogCancellationReturnsNoPath()
        {
            bool closedOwnedDialog = false;
            using (var owner = new Form())
            using (var timer = new System.Windows.Forms.Timer { Interval = 50 })
            {
                timer.Tick += (_, __) =>
                {
                    EnumThreadWindows(GetCurrentThreadId(), (handle, state) =>
                    {
                        var name = new StringBuilder(64);
                        GetClassName(handle, name, name.Capacity);
                        if (name.ToString() != "#32770") return true;
                        timer.Stop(); closedOwnedDialog = true;
                        PostMessage(handle, 0x0010, IntPtr.Zero, IntPtr.Zero);
                        return false;
                    }, IntPtr.Zero);
                };
                timer.Start();
                Assert.IsNull(TestExplorerWindow.SelectReportExportPath(owner, false));
                Assert.IsTrue(closedOwnedDialog, "Only a native dialog belonging to this test UI thread may be dismissed.");
            }
        }

        private delegate bool EnumWindowCallback(IntPtr window, IntPtr state);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint thread, EnumWindowCallback callback, IntPtr state);
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder value, int capacity);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr window, uint message, IntPtr wparam, IntPtr lparam);
        [STATestMethod]
        public void ThemeCallbacksMarshalToOwnerAndIgnoreDisposedOrUncreatedForeignThreadWindows()
        {
            using (var window = new TestExplorerWindow())
            {
                window.Configure(new ExplorerDouble());
                Exception workerError = null;
                var worker = new Thread(() => { try { Invoke(window, "TestThemeChanged"); } catch (Exception error) { workerError = error; } });
                worker.Start(); worker.Join(); Assert.IsNull(workerError);
                var handle = window.Handle;
                worker = new Thread(() => { try { Invoke(window, "TestThemeChanged"); } catch (Exception error) { workerError = error; } });
                worker.Start(); worker.Join(); Assert.IsNull(workerError);
                Application.DoEvents();
                Assert.AreEqual(2, window.VisibleTests().Count);
                window.Dispose(); Invoke(window, "TestThemeChanged");
            }
        }

        [STATestMethod]
        public void SourceFormattingAndNoLongerPresentTreeScopesHaveSafeFallbacks()
        {
            var service = new ExplorerDouble();
            using (var window = new TestExplorerWindow())
            {
                window.Configure(service);
                foreach (string path in new[] { "", @"C:\Disposable\Tests.xlsm" })
                {
                    service.Catalog.Project.HostPath = path;
                    var args = new ListControlConvertEventArgs("", typeof(string), service.Catalog.Project);
                    Invoke(window, "ProjectList_Format", window, args);
                    StringAssert.Contains((string)args.Value, "Disposable");
                    if (path.Length > 0) StringAssert.Contains((string)args.Value, path);
                }
                var unrecognized = new ListControlConvertEventArgs("unchanged", typeof(string), new object());
                Invoke(window, "ProjectList_Format", window, unrecognized);
                Assert.AreEqual("unchanged", unrecognized.Value);
                var tree = Field<TreeView>(window, "testTree"); var handle = tree.Handle;
                tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[0];
                Assert.AreEqual(2, window.ScopeTests().Count);
                Field<ComboBox>(window, "grouping").SelectedIndex = 2;
                Assert.AreEqual("first", ((VbaTestDescriptor)tree.SelectedNode.Tag).Id);
                Field<ComboBox>(window, "grouping").SelectedIndex = 0;
                tree.SelectedNode = tree.Nodes[0].Nodes[0];
                service.Catalog.Modules[0].Name = "Renamed";
                window.RefreshProjects();
                Assert.AreEqual(0, window.ScopeTests().Count);
                service.Catalog.Project.Id = "replacement";
                window.RefreshProjects();
                Assert.AreEqual(2, window.ScopeTests().Count);
                var detached = new TreeNode("Unrecognized");
                var selectionMethod = typeof(TestExplorerWindow).GetMethod("ReadTreeSelection", BindingFlags.Static | BindingFlags.NonPublic);
                Assert.IsNull(selectionMethod.Invoke(null, new object[] { detached }));
                var emptyModule = new TreeNode("Empty") { Tag = new VbaTestModule() };
                var root = tree.Nodes[0]; root.Nodes.Add(emptyModule);
                Invoke(window, "UpdateGroupChecks"); Assert.IsFalse(emptyModule.Checked); Assert.IsFalse(root.Checked);
                root.Nodes.Clear(); Invoke(window, "UpdateGroupChecks"); Assert.IsFalse(root.Checked);
            }
        }
        [STATestMethod]
        public void UnselectedAndUnavailableCatalogStatesKeepScopeAndNavigationEmpty()
        {
            using (var window = new TestExplorerWindow())
            {
                Assert.AreEqual(0, window.ScopeTests().Count);
                var service = new ExplorerDouble(); window.Configure(service);
                var tree = Field<TreeView>(window, "testTree"); var handle = tree.Handle;
                tree.SelectedNode = null;
                Invoke(window, "Source_Click", window, EventArgs.Empty);
                Assert.AreEqual(0, service.Navigations);
                var test = service.Catalog.Modules[0].Tests[0];
                Field<Dictionary<string, string>>(window, "resultRevisions")[test.Id] = "r1";
                SetField(window, "catalog", null);
                Assert.AreEqual(true, Invoke(window, "IsStale", test));
                SetField(window, "reviewingCoverage", true);
                Assert.ThrowsException<TargetInvocationException>(() => Invoke(window, "SelectProject", "project"));
                SetField(window, "reviewingCoverage", false);
                service.DiscoveryError = new InvalidOperationException("Temporarily inaccessible");
                window.RefreshProjects(); window.Show();
                service.DiscoveryError = null;
                window.CheckFreshness();
                Assert.AreEqual(2, window.VisibleTests().Count);
                Assert.AreEqual(0, service.Runs);
            }
        }

        [STATestMethod]
        public void SelectionResolverIgnoresUnrecognizedNodesAndRetainsOnlyMatchingModuleAndGroupKeys()
        {
            var type = typeof(TestExplorerWindow);
            var selectionType = type.GetNestedType("TreeSelection", BindingFlags.NonPublic);
            var groupType = type.GetNestedType("TestGroup", BindingFlags.NonPublic);
            var find = type.GetMethod("FindSelection", BindingFlags.Static | BindingFlags.NonPublic);
            var root = new TreeNode("Disposable");
            var module = new TreeNode("Tests") { Tag = new VbaTestModule { Name = "Tests" } };
            var wrongGroup = Activator.CreateInstance(groupType, true);
            SetField(wrongGroup, "Grouping", 1); SetField(wrongGroup, "Key", "Failed");
            var category = Activator.CreateInstance(groupType, true);
            SetField(category, "Grouping", 2); SetField(category, "Key", "Regression");
            var outcomeNode = new TreeNode("Failed") { Tag = wrongGroup };
            var categoryNode = new TreeNode("Regression") { Tag = category };
            root.Nodes.Add(new TreeNode("Unrecognized")); root.Nodes.Add(module); root.Nodes.Add(outcomeNode); root.Nodes.Add(categoryNode);
            module.Nodes.Add(new TreeNode("Unrecognized child"));
            module.Nodes.Add(new TreeNode("Adds") { Tag = new VbaTestDescriptor { Id = "first" } });
            var selected = Activator.CreateInstance(selectionType, true);
            SetField(selected, "TestId", "first");
            Assert.AreSame(module.Nodes[1], find.Invoke(null, new[] { root, selected }));
            SetField(selected, "TestId", null); SetField(selected, "Module", "Tests");
            Assert.AreSame(module, find.Invoke(null, new[] { root, selected }));
            SetField(selected, "Module", null); SetField(selected, "Grouping", 2); SetField(selected, "GroupKey", "Absent");
            Assert.IsNull(find.Invoke(null, new[] { root, selected }));
            SetField(selected, "GroupKey", "Regression");
            Assert.AreSame(categoryNode, find.Invoke(null, new[] { root, selected }));
        }

        [STATestMethod]
        public void ExportFailureDuringDisposingDoesNotMutateTheTearingDownUi()
        {
            using (var window = new TestExplorerWindow())
            {
                window.Configure(new ExplorerDouble()); window.RunScopeAsync().GetAwaiter().GetResult();
                Field<TabControl>(window, "resultTabs").SelectedTab = Field<TabPage>(window, "humanTab");
                string status = Field<Label>(window, "status").Text;
                // Simulate the WinForms reentrant disposal phase without destroying user-owned native handles.
                var setState = typeof(Control).GetMethod("SetState", BindingFlags.Instance | BindingFlags.NonPublic);
                window.ChooseReportExportPath = (_, __) => { setState.Invoke(window, new object[] { 0x1000, true }); Assert.IsTrue(window.Disposing); throw new IOException("Disposal export failure"); };
                try
                {
                    Invoke(window, "ExportReport_Click", window, EventArgs.Empty);
                    Assert.AreEqual(status, Field<Label>(window, "status").Text);
                }
                finally { setState.Invoke(window, new object[] { 0x1000, false }); }
            }
        }
        private sealed class BrokenDispatcherControl : Control
        {
            protected override void CreateHandle() => throw new InvalidOperationException("Dispatcher handle creation failed");
        }

        private static object Invoke(object target, string method, params object[] arguments) => target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, arguments);
        private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static T Field<T>(object target, string name) => (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static string DescribeCallbackQueue(Control control)
        {
            // Read-only diagnostic of the .NET Framework WinForms marshal queue; never dequeue callbacks.
            try
            {
                var field = typeof(Control).GetField("threadCallbackList", BindingFlags.Instance | BindingFlags.NonPublic);
                var queue = field?.GetValue(control) as System.Collections.ICollection;
                if (queue == null) return field == null ? "unavailable" : "null";
                lock (queue.SyncRoot)
                {
                    var methods = new List<string>();
                    foreach (object entry in queue)
                    {
                        var callback = entry.GetType().GetField("method", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)?.GetValue(entry) as Delegate;
                        methods.Add(callback == null ? entry.GetType().Name : callback.Method.DeclaringType?.FullName + "." + callback.Method.Name);
                    }
                    return queue.Count + " [" + string.Join(", ", methods) + "]";
                }
            }
            catch (Exception error) { return "read failed: " + error.GetType().Name; }
        }
        private static void PumpOwner(Task task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(task.IsCompleted, "The UI owner did not receive its queued test continuation within five seconds.");
            task.GetAwaiter().GetResult();
        }
        private static void RaiseThemeChanged() => ((Action)typeof(UiTheme).GetField("Changed", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null))?.Invoke();

        private class ExplorerDouble : IVbaTestExplorerService
        {
            internal readonly VbaTestCatalog Catalog = new VbaTestCatalog
            {
                Project = new VbaTestProjectSnapshot { Id = "project", Name = "Disposable", Revision = "r1" },
                Modules = new List<VbaTestModule>
                {
                    new VbaTestModule { Name = "Tests", Tests = new List<VbaTestDescriptor>
                    {
                        new VbaTestDescriptor { Id = "first", Module = "Tests", Procedure = "Adds", Line = 4, Kind = "Function", Categories = new[] { "Arithmetic" } },
                        new VbaTestDescriptor { Id = "second", Module = "Tests", Procedure = "Other", Line = 10, Kind = "Sub" }
                    } }
                }
            };
            internal int Runs;
            internal bool PassAll;
            internal string Unavailable;
            internal IReadOnlyList<VbaTestDescriptor> Selected;
            internal CancellationToken Cancellation;
            internal TaskCompletionSource<VbaTestRun> Pending;
            internal Exception DiscoveryError;
            internal int DiscoveryCalls;
            internal Action<VbaTestResult> SavedProgress;
            internal int Navigations, Installations;
            internal Exception NavigationError, InstallError, ReadError, AvailabilityError;
            internal bool NoProjects;
            public IReadOnlyList<VbaTestProjectSnapshot> ReadProjects() { if (ReadError != null) throw ReadError; return NoProjects ? new VbaTestProjectSnapshot[0] : new[] { Catalog.Project }; }
            public VbaTestCatalog Discover(string projectId) { DiscoveryCalls++; if (DiscoveryError != null) throw DiscoveryError; return Catalog; }
            public virtual string ExecutionUnavailableReason(VbaTestCatalog catalog) { if (AvailabilityError != null) throw AvailabilityError; return Unavailable; }
            public virtual Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests, Action<VbaTestResult> onResult, CancellationToken cancellation)
            {
                Runs++;
                Selected = tests;
                Cancellation = cancellation;
                SavedProgress = onResult;
                if (Pending != null) return Pending.Task;
                var run = new VbaTestRun { Id = "run", Project = catalog.Project.Name, Revision = catalog.Project.Revision };
                foreach (var test in tests)
                {
                    var result = new VbaTestResult { Test = test, Outcome = !PassAll && test.Id == "first" ? VbaTestOutcome.Failed : VbaTestOutcome.Passed,
                        Message = !PassAll && test.Id == "first" ? "expected 4, actual 3" : "Completed", Duration = TimeSpan.FromMilliseconds(12), Phase = "test" };
                    run.Results.Add(result);
                    onResult(result);
                }
                return Task.FromResult(run);
            }
            public void Navigate(VbaTestCatalog catalog, VbaTestDescriptor test) { Navigations++; if (NavigationError != null) throw NavigationError; }
            public void InstallSupport(VbaTestCatalog catalog) { Installations++; if (InstallError != null) throw InstallError; }
        }

        private sealed class CoverageDouble : ExplorerDouble, IVbaTestCoverageExplorerService
        {
            internal int CoverageRuns;
            internal string CoverageUnavailable;
            internal VbaCoverageReport Report = new VbaCoverageReport { Available = true, Complete = false, DenominatorKnown = true, Revision = "r1", Eligible = 2, Hit = 1, Percent = 50 };
            public string CoverageUnavailableReason(VbaTestCatalog catalog) => CoverageUnavailable;
            public async Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests, Action<VbaTestResult> onResult, CancellationToken cancellation)
            {
                CoverageRuns++;
                var run = await RunAsync(catalog, tests, onResult, cancellation);
                run.Coverage = Report;
                return run;
            }
        }

        private sealed class ContextClearingCoverageDouble : ExplorerDouble, IVbaTestCoverageExplorerService
        {
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            internal int AvailabilityCalls, WrongOwnerCalls;
            internal bool ContextWasCleared;
            internal readonly TaskCompletionSource<VbaTestRun> Completion = new TaskCompletionSource<VbaTestRun>(TaskCreationOptions.RunContinuationsAsynchronously);
            private void RequireOwner()
            {
                if (owner == Thread.CurrentThread.ManagedThreadId) return;
                Interlocked.Increment(ref WrongOwnerCalls);
                throw new InvalidOperationException("Coverage UI cleanup called the service outside its owner thread.");
            }
            public override string ExecutionUnavailableReason(VbaTestCatalog catalog)
            { RequireOwner(); AvailabilityCalls++; return null; }
            public string CoverageUnavailableReason(VbaTestCatalog catalog) { RequireOwner(); return null; }
            public Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
                Action<VbaTestResult> onResult, CancellationToken cancellation)
            {
                RequireOwner(); Selected = tests; Cancellation = cancellation;
                SynchronizationContext.SetSynchronizationContext(null);
                ContextWasCleared = SynchronizationContext.Current == null;
                return Completion.Task;
            }
        }

        private sealed class DisposalDouble : ExplorerDouble, IVbaTestCoverageExplorerService
        {
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            internal bool Disposed, ClearContext;
            internal int AvailabilityCalls, WrongOwnerCalls;
            internal int DispatchThread;
            internal SynchronizationContext ContextBeforeReturn;
            public override string ExecutionUnavailableReason(VbaTestCatalog catalog)
            {
                if (Thread.CurrentThread.ManagedThreadId != owner) { WrongOwnerCalls++; throw new InvalidOperationException("Wrong service owner"); }
                AvailabilityCalls++;
                if (Disposed) throw new ObjectDisposedException("Native service");
                return null;
            }
            public string CoverageUnavailableReason(VbaTestCatalog catalog) => ExecutionUnavailableReason(catalog);
            public override Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
                Action<VbaTestResult> onResult, CancellationToken cancellation)
            {
                var task = base.RunAsync(catalog, tests, onResult, cancellation);
                if (ClearContext) SynchronizationContext.SetSynchronizationContext(null);
                DispatchThread = Thread.CurrentThread.ManagedThreadId;
                ContextBeforeReturn = SynchronizationContext.Current;
                return task;
            }
            public Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
                Action<VbaTestResult> onResult, CancellationToken cancellation) => RunAsync(catalog, tests, onResult, cancellation);
        }
    }
}
