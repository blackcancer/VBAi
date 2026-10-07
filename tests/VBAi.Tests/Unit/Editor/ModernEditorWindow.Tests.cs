using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.ComponentModel;
using System.Linq;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorWindowTests
    {
        private const string ClosedProjectError = "The VBA project is closed. Your draft is preserved.";
        private const string SynchronizedStatus = "Synchronized with VBA. Save the macro in its host application.";

        /// <summary>A closed background native project retains its draft and diagnostic without replacing live-module status.</summary>
        [STATestMethod]
        public void InactiveClosedProjectKeepsDraftAndFailureWithoutContaminatingLiveStatus()
        {
            var previousLog = LoadLog.AppendText;
            var diagnostics = new System.Collections.Generic.List<string>();
            try
            {
                LoadLog.AppendText = (path, text) => diagnostics.Add(text);
                using (var f = new ModernEditorDebugFixture())
                using (var live = new EditorFixture())
                {
                    string native = f.Native.Original.CodeModule.Raw;
                    string draft = f.Document.Text + "\n' preserved closed-project edit";
                    f.Document.Edit(draft);
                    var active = ModernEditorDebugFixture.Wait(f.Window.OpenModule(live));
                    f.Native.Vbe.VBProjects.Clear(); f.Ready(true);
                    ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true));
                    Assert.AreSame(active, f.Window.Current);
                    Assert.AreEqual(UiText.Get(SynchronizedStatus), f.Get<System.Windows.Forms.Label>("status").Text);
                    Assert.IsTrue(f.Window.Documents.Contains(f.Document));
                    Assert.AreEqual(draft, f.Document.Text);
                    Assert.AreEqual(draft, f.Window.Drafts.Recover(f.Document.RecoveryKey).Text);
                    Assert.AreEqual(native, f.Native.Original.CodeModule.Raw, "Closed native identity must still refuse writes.");
                    Assert.IsTrue(diagnostics.Any(line => line.Contains("Monaco: InvalidOperationException")), "Background failures must remain diagnosable.");

                    ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Native.Adapter));
                    Assert.AreEqual(UiText.Get(ClosedProjectError), f.Get<System.Windows.Forms.Label>("status").Text);
                    ModernEditorDebugFixture.Wait(f.Window.OpenModule(live));
                    Assert.AreEqual(UiText.Get(SynchronizedStatus), f.Get<System.Windows.Forms.Label>("status").Text);
                    Assert.AreEqual(0, live.Writes);
                }
            }
            finally { LoadLog.AppendText = previousLog; }
        }

        /// <summary>Successful inactive work cannot clear an active write failure; a later successful observation can clear only its own failure.</summary>
        [STATestMethod]
        public void ActiveSynchronizationFailureSurvivesInactiveCompletionAndClearsAfterItsOwnSuccess()
        {
            using (var f = new ModernEditorDebugFixture())
            using (var activeModule = new EditorFixture { Fail = true })
            {
                var active = ModernEditorDebugFixture.Wait(f.Window.OpenModule(activeModule));
                active.Edit(active.Text + "\n' pending write"); f.Ready(true);
                ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true));
                Assert.AreEqual(UiText.Get("Write refused"), f.Get<System.Windows.Forms.Label>("status").Text);
                Assert.IsTrue(active.Dirty); Assert.AreEqual(0, activeModule.Writes);
                ModernEditorDebugFixture.Wait((System.Threading.Tasks.Task)UiInvoke.Call(typeof(ModernEditorWindow), "ProcessCapturedDocumentsCore",
                    f.Window, false, false, false, null, new[] { f.Document }));
                Assert.AreEqual(UiText.Get("Write refused"), f.Get<System.Windows.Forms.Label>("status").Text,
                    "A batch that did not process the active document must retain its failure.");
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Native.Adapter));
                Assert.AreEqual(UiText.Get(SynchronizedStatus), f.Get<System.Windows.Forms.Label>("status").Text);
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(activeModule));
                Assert.AreEqual(UiText.Get("Write refused"), f.Get<System.Windows.Forms.Label>("status").Text);
                activeModule.Fail = false;
                // Read-only reconciliation succeeds while preserving the unsynchronized draft.
                ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(false));
                Assert.AreEqual(UiText.Get("Changes pending synchronization with VBA."), f.Get<System.Windows.Forms.Label>("status").Text);
                Assert.IsTrue(active.Dirty); Assert.AreEqual(0, activeModule.Writes);
            }
        }

        /// <summary>A batch completing across a real renderer await presents the currently selected document, keeping other failures attached to their tabs.</summary>
        [STATestMethod]
        public void BackgroundBatchCompletionAfterAsyncTabSwitchUsesCurrentDocumentStatus()
        {
            using (var f = new ModernEditorDebugFixture())
            using (var updating = new EditorFixture())
            using (var next = new EditorFixture())
            {
                var updatingDocument = ModernEditorDebugFixture.Wait(f.Window.OpenModule(updating));
                var nextDocument = ModernEditorDebugFixture.Wait(f.Window.OpenModule(next));
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(updating)); f.Ready(true);
                f.Native.Vbe.VBProjects.Clear(); updating.Code += "\n' external update";
                var entered = new System.Threading.Tasks.TaskCompletionSource<bool>();
                var release = new System.Threading.Tasks.TaskCompletionSource<string>();
                var renderer = f.Window.ScriptExecution;
                f.Window.ScriptExecution = (method, values) =>
                {
                    if (method == "apply" && (string)values[0] == updatingDocument.Id)
                    { entered.TrySetResult(true); return release.Task; }
                    return renderer(method, values);
                };
                var processing = (System.Threading.Tasks.Task)UiInvoke.Call(typeof(ModernEditorWindow), "ProcessCapturedDocumentsCore",
                    f.Window, false, false, false, null, new[] { f.Document, updatingDocument });
                try
                {
                    ModernEditorDebugFixture.Wait(entered.Task);
                    Assert.IsFalse(processing.IsCompleted, "The test requires a pending real renderer boundary.");
                    ModernEditorDebugFixture.Wait(f.Window.OpenModule(next));
                    Assert.AreSame(nextDocument, f.Window.Current);
                }
                finally { release.TrySetResult("2"); ModernEditorDebugFixture.Wait(processing); }
                Assert.AreEqual(UiText.Get(SynchronizedStatus), f.Get<System.Windows.Forms.Label>("status").Text);
                Assert.AreEqual(EditorDocument.Normalize(updating.Code), updatingDocument.Text);
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Native.Adapter));
                Assert.AreEqual(UiText.Get(ClosedProjectError), f.Get<System.Windows.Forms.Label>("status").Text);
                Assert.AreEqual(0, updating.Writes); Assert.AreEqual(0, next.Writes);
            }
        }

        /// <summary>A newer tool or dispatch error still invalidates queued reconciliation status rather than being replaced by a healthy tab.</summary>
        [STATestMethod]
        public void DocumentScopedSynchronizationDoesNotSuppressNewerToolDispatchErrors()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                UiInvoke.Call(typeof(ModernEditorWindow), "SetStatus", f.Window);
                UiInvoke.Call(typeof(ModernEditorWindow), "Report", f.Window, new System.InvalidOperationException("tool dispatch refused"));
                System.Windows.Forms.Application.DoEvents();
                Assert.AreEqual(UiText.Get("tool dispatch refused"), f.Get<System.Windows.Forms.Label>("status").Text);
                f.Set("lastSaveError", "native save unverified");
                UiInvoke.Call(typeof(ModernEditorWindow), "UpdateStatus", f.Window);
                Assert.AreEqual(UiText.Get("native save unverified"), f.Get<System.Windows.Forms.Label>("status").Text);
            }
        }

        /// <summary>Ignores theme changes during disposal and dispatches worker notifications onto the editor thread.</summary>
        [STATestMethod]
        public void ThemeNotificationsRespectWindowLifetimeAndUiThreadOwnership()
        {
            using (var uncreated = new ModernEditorWindow())
            { UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", uncreated); Assert.IsFalse(uncreated.IsHandleCreated); }
            using (var fixture = new ModernEditorDebugFixture())
            {
                foreach (bool closing in new[] { true, false })
                {
                    fixture.Set("closing", closing); fixture.Scripts.Clear();
                    var thread = new System.Threading.Thread(() => UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", fixture.Window));
                    thread.Start(); Assert.IsTrue(thread.Join(5000));
                    System.Windows.Forms.Application.DoEvents();
                    Assert.AreEqual(closing ? 0 : 1, fixture.Scripts.Count(script => script.Item1 == "theme"));
                }
                bool refused = false;
                fixture.Window.DispatchTheme = action => { refused = true; throw new System.InvalidOperationException("window closed during dispatch"); };
                var closingThread = new System.Threading.Thread(() => UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", fixture.Window));
                closingThread.Start(); Assert.IsTrue(closingThread.Join(5000)); Assert.IsTrue(refused);
                var panel = new System.Windows.Forms.Panel(); fixture.Window.Controls.Add(panel);
                panel.Disposed += (sender, args) => UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", fixture.Window);
                fixture.Scripts.Clear(); fixture.Window.Dispose();
                UiInvoke.Call(typeof(ModernEditorWindow), "ThemeChanged", fixture.Window);
                Assert.AreEqual(0, fixture.Scripts.Count);
            }
        }

        /// <summary>Diff and tab close keep renderer calls on the owning STA even without an ambient UI context.</summary>
        [STATestMethod]
        public void DiffAndCloseWithoutAmbientContextKeepRendererCallsOnTheOwningThread()
        {
            var previousContext = System.Threading.SynchronizationContext.Current;
            try
            {
                using (var f = new Editor.ModernEditorToolFixture())
                {
                    int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                    int initialCount = f.Window.Documents.Count();
                    Assert.AreSame(f.Document, f.Window.Current);
                    var calls = new System.Collections.Concurrent.ConcurrentQueue<System.Tuple<string, int>>();
                    f.Override = (method, values) =>
                    {
                        calls.Enqueue(System.Tuple.Create(method, System.Threading.Thread.CurrentThread.ManagedThreadId));
                        return null;
                    };
                    System.Action<System.Func<bool>> pump = complete =>
                    {
                        var timeout = System.Diagnostics.Stopwatch.StartNew();
                        while (!complete() && timeout.ElapsedMilliseconds < 10000)
                        { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                        Assert.IsTrue(complete(), "The owned editor action did not complete; status: " + UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text);
                    };

                    System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                    UiInvoke.Call(typeof(ModernEditorWindow), "DiffClick", f.Window, null, System.EventArgs.Empty);
                    pump(() => UiInvoke.Field<bool>(f.Window, "showingDiff") && !UiInvoke.Field<bool>(f.Window, "busy"));
                    Assert.IsTrue(calls.Any(call => call.Item1 == "compare"));
                    Assert.IsTrue(calls.All(call => call.Item2 == ownerThread), "Diff renderer calls must remain on the editor's owning STA.");
                    Assert.AreEqual(0, f.Module.Writes);

                    System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                    UiInvoke.Call(typeof(ModernEditorWindow), "CloseModuleClick", f.Window, null, System.EventArgs.Empty);
                    pump(() => f.Window.Documents.Count() == initialCount - 1 && !UiInvoke.Field<bool>(f.Window, "busy"));
                    Assert.IsFalse(f.Window.Documents.Contains(f.Document));
                    Assert.IsTrue(calls.Any(call => call.Item1 == "close"));
                    Assert.IsTrue(calls.All(call => call.Item2 == ownerThread), "Close renderer calls must remain on the editor's owning STA.");
                    Assert.AreEqual(0, f.Module.Writes, "Diff and tab close must not write the native module.");
                }
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previousContext); }
        }

        /// <summary>Reload and draft restoration keep renderer calls on the owner after a real asynchronous script boundary.</summary>
        [STATestMethod]
        public void ResolveReloadAndRestoreWithoutAmbientContextKeepRendererOwnershipAndArchivedDraft()
        {
            var previousContext = System.Threading.SynchronizationContext.Current;
            try
            {
                using (var f = new Editor.ModernEditorToolFixture())
                {
                    int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                    string draft = f.Document.Text + "\n' owned pending recovery";
                    f.Document.Edit(draft);
                    f.Module.Code += "\n' owned concurrent native change";
                    f.Document.Observe();
                    Assert.IsTrue(f.Document.Dirty && f.Document.Conflict);
                    var calls = new System.Collections.Concurrent.ConcurrentQueue<System.Tuple<string, int>>();
                    f.Override = (method, values) =>
                    {
                        calls.Enqueue(System.Tuple.Create(method, System.Threading.Thread.CurrentThread.ManagedThreadId));
                        return null;
                    };
                    var renderer = f.Window.ScriptExecution;
                    var ledger = new System.Collections.Concurrent.ConcurrentQueue<string>();
                    var elapsed = System.Diagnostics.Stopwatch.StartNew();
                    int invocationSequence = 0;
                    System.Action<int, string, string> trace = (invocation, method, stage) =>
                    {
                        if (method == "snapshots" || method == "apply" || method == "hideDiff")
                            ledger.Enqueue(elapsed.ElapsedMilliseconds + "ms #" + invocation + " " + method + " " + stage +
                                " thread=" + System.Threading.Thread.CurrentThread.ManagedThreadId);
                    };
                    f.Window.ScriptExecution = async (method, values) =>
                    {
                        int invocation = System.Threading.Interlocked.Increment(ref invocationSequence);
                        trace(invocation, method, "enter");
                        try
                        {
                            string result = await renderer(method, values);
                            trace(invocation, method, "renderer-return");
                            await System.Threading.Tasks.Task.Delay(10).ConfigureAwait(false);
                            trace(invocation, method, "delay-complete-return");
                            return result;
                        }
                        catch (System.Exception error) { trace(invocation, method, "error=" + error); throw; }
                    };
                    System.Action<System.Func<bool>> pump = complete =>
                    {
                        var timeout = System.Diagnostics.Stopwatch.StartNew();
                        while (!complete() && timeout.ElapsedMilliseconds < 5000)
                        { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                        Assert.IsTrue(complete(), "The owned editor recovery action did not complete; status: " + UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text +
                            "; Dirty=" + f.Document.Dirty + "; Conflict=" + f.Document.Conflict +
                            "; Diff=" + f.Base.Get<bool>("showingDiff") + "; Busy=" + f.Base.Get<bool>("busy") +
                            "; Current=" + (f.Window.Current?.Id ?? "<null>") + "; Target=" + f.Document.Id +
                            "; Closing=" + f.Base.Get<bool>("closing") + "; Disposed=" + f.Window.IsDisposed +
                            "; Timer=" + UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "timer").Enabled +
                            "; StreamTimer=" + UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "streamTimer").Enabled +
                            "; DebugTimer=" + UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "debugTimer").Enabled +
                            "; Writes=" + f.Module.Writes + "; Captures=" + f.Captures + "; Applies=" + f.Applies +
                            "; Reviewed=" + f.Base.Get<System.Collections.Generic.Dictionary<string, string>>("reviewed").ContainsKey(f.Document.Id) +
                            "; Ledger=" + string.Join(" | ", ledger));
                    };

                    f.Base.Get<System.Collections.Generic.Dictionary<string, string>>("reviewed")[f.Document.Id] = f.Document.Native;
                    f.Base.Set("showingDiff", true);
                    System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                    UiInvoke.Call(typeof(ModernEditorWindow), "ResolveClick", f.Window, null, System.EventArgs.Empty);
                    pump(() => !f.Document.Dirty && !f.Base.Get<bool>("showingDiff") && !f.Base.Get<bool>("busy"));
                    Assert.IsTrue(calls.Any(call => call.Item1 == "hideDiff"));
                    Assert.IsTrue(calls.All(call => call.Item2 == ownerThread), "Resolve renderer calls must remain on the owning STA.");
                    Assert.AreEqual(1, f.Module.Writes, "Only the explicit conflict resolution may write native source.");
                    StringAssert.Contains(f.Module.Code, "owned pending recovery");

                    draft = f.Document.Text + "\n' owned reloaded recovery";
                    f.Document.Edit(draft);
                    f.Module.Code += "\n' owned second native change";
                    f.Document.Observe();
                    Assert.IsTrue(f.Document.Dirty && f.Document.Conflict);
                    System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                    UiInvoke.Call(typeof(ModernEditorWindow), "ReloadClick", f.Window, null, System.EventArgs.Empty);
                    pump(() => !f.Document.Dirty && !UiInvoke.Field<bool>(f.Window, "busy"));
                    Assert.AreEqual(EditorDocument.Normalize(f.Module.Code), f.Document.Text);
                    Assert.AreEqual(draft, f.Window.Drafts.Recover(f.Document.RecoveryKey).Text);
                    Assert.IsTrue(calls.Any(call => call.Item1 == "apply"));
                    Assert.IsTrue(calls.All(call => call.Item2 == ownerThread), "Reload renderer calls must remain on the owning STA.");
                    Assert.AreEqual(1, f.Module.Writes);
                    int applies = f.Applies;

                    System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                    UiInvoke.Call(typeof(ModernEditorWindow), "RestoreClick", f.Window, null, System.EventArgs.Empty);
                    pump(() => f.Document.Dirty && f.Document.Text == draft && !UiInvoke.Field<bool>(f.Window, "busy"));
                    Assert.AreEqual(applies + 1, f.Applies);
                    Assert.IsTrue(calls.All(call => call.Item2 == ownerThread), "Restore renderer calls must remain on the owning STA.");
                    Assert.AreEqual(draft, f.Window.Drafts.Recover(f.Document.RecoveryKey).Text);
                    Assert.AreEqual(1, f.Module.Writes, "Reload and draft restoration must not add native writes after conflict resolution.");
                }
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previousContext); }
        }

        /// <summary>Preserves visible conflict captions and scales command height for large fonts and long translations.</summary>
        [STATestMethod]
        public void ConflictChoicesKeepCaptionsAndFitScaledToolbar()
        {
            foreach (float scale in new[] { 1f, 1.5f, 2f })
                using (var module = new EditorFixture())
                using (var window = new ModernEditorWindow())
                using (var font = new System.Drawing.Font("Segoe UI", 12f * scale))
                {
                    window.Drafts = new EditorDraftStore(module.Root);
                    var doc = ModernEditorDebugFixture.Wait(window.OpenModule(module));
                    doc.Edit("edited source"); module.Code = "changed native source"; doc.Observe();
                    typeof(ModernEditorWindow).GetField("initializing", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(window, true);
                    window.Scale(new System.Drawing.SizeF(scale, scale)); window.Font = font;
                    window.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    window.Location = new System.Drawing.Point(-10000, -10000);
                    window.Show();
                    var resolve = UiInvoke.Field<ThemedButton>(window, "resolve"); var reload = UiInvoke.Field<ThemedButton>(window, "reload");
                    foreach (var caption in new[] { "Utiliser la version modifiée", new string('W', 80) })
                    {
                        resolve.Text = caption; reload.Text = "Recharger la version VBA · " + caption;
                        UiInvoke.Call(typeof(ModernEditorWindow), "UpdateStatus", window); window.PerformLayout();
                        var toolbar = UiInvoke.Field<System.Windows.Forms.FlowLayoutPanel>(window, "toolbar");
                        foreach (var button in new[] { resolve, reload })
                        {
                            Assert.IsFalse(button.IconOnly); Assert.IsTrue(button.AutoSize); Assert.IsTrue(button.Visible);
                            Assert.AreEqual(button.Text, button.AccessibilityObject.Name);
                            Assert.IsTrue(button.Width >= button.GetPreferredSize(System.Drawing.Size.Empty).Width);
                            Assert.IsTrue(toolbar.ClientSize.Height >= button.Height + button.Margin.Vertical + toolbar.Padding.Vertical);
                        }
                    }
                }
        }

        [STATestMethod]
        public void DesignerAndConstructionNeverLaunchWebViewOrReadVba()
        {
            using (var window = new ModernEditorWindow()) { Assert.IsNull(window.Browser); Assert.IsFalse(window.Ready); }
            using (var window = (ModernEditorWindow)LicenseManager.CreateWithContext(typeof(ModernEditorWindow), new DesignContext()))
            { Assert.IsNull(window.Browser); Assert.AreEqual("VBAi editor", window.Text); }
        }
        [DataTestMethod]
        [DataRow("Clear All Breakpoints")]
        [DataRow("Effacer tous les points d'arrêt")]
        public void ToggleNeverAcceptsClearAllBreakpoints(string caption)
        { Assert.IsFalse(VbeDebug.IsAllowed("toggle_breakpoint", caption, 2)); }
        [STATestMethod]
        public void WorkspaceCannotBeClosedByANativeCloseCommand()
        {
            using (var window = new ModernEditorWindow())
            {
                window.WorkspaceHosted = true;
                var args = new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false);
                UiInvoke.Call(typeof(ModernEditorWindow), "ClosingWindow", window, window, args);
                Assert.IsTrue(args.Cancel);
                Assert.IsFalse(window.ControlBox);
                Assert.AreEqual(System.Windows.Forms.FormBorderStyle.None, window.FormBorderStyle);
                Assert.IsTrue(UiInvoke.Field<ThemedTabControl>(window, "tabs").ShowCloseButtons);
            }
        }
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern System.IntPtr SendMessage(System.IntPtr window, uint message, System.IntPtr first, System.IntPtr second);
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern System.IntPtr SetParent(System.IntPtr child, System.IntPtr parent);

        [STATestMethod]
        public void HostedWorkspaceSurvivesActualWmCloseWithoutStartingDisposal()
        {
            using (var parent = new System.Windows.Forms.Form { IsMdiContainer = true })
            using (var window = new ModernEditorWindow())
            {
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                var initializing = typeof(ModernEditorWindow).GetField("initializing", flags);
                initializing.SetValue(window, true); // Suppress WebView startup in this native-message regression.
                window.WorkspaceHosted = true;
                window.TopLevel = false;
                parent.Show();
                var workspace = parent.Controls.OfType<System.Windows.Forms.MdiClient>().Single();
                SetParent(window.Handle, workspace.Handle);
                window.Show();
                System.Windows.Forms.Application.DoEvents();
                initializing.SetValue(window, false);
                var original = window.Handle;
                System.Windows.Forms.CloseReason? observed = null;
                bool? cancelled = null;
                window.FormClosing += (sender, args) => { observed = args.CloseReason; cancelled = args.Cancel; };

                SendMessage(original, 0x10, System.IntPtr.Zero, System.IntPtr.Zero);
                System.Windows.Forms.Application.DoEvents();

                Assert.AreEqual(System.Windows.Forms.CloseReason.TaskManagerClosing, observed);
                Assert.AreEqual(true, cancelled);
                Assert.IsFalse(window.IsDisposed);
                Assert.IsTrue(window.Visible);
                Assert.AreEqual(original, window.Handle);
                Assert.IsFalse((bool)typeof(ModernEditorWindow).GetField("closing", flags).GetValue(window));
                Assert.IsNull(window.Browser);
            }
        }
        [DataTestMethod]
        [DataRow("https://editor.vbai.local/index.html", true)]
        [DataRow("https://editor.vbai.local/index.html?external", false)]
        [DataRow("https://editor.vbai.local.evil.test/index.html", false)]
        [DataRow("http://editor.vbai.local/index.html", false)]
        [DataRow("https://user@editor.vbai.local/index.html", false)]
        public void BridgeAcceptsOnlyTheBundledDocument(string source, bool allowed)
        { Assert.AreEqual(allowed, ModernEditorWindow.Trusted(source)); }
        [STATestMethod]
        public void LocalResourcePolicyRejectsAllExternalAuthorityForms()
        {
            foreach (string uri in new[] { null, "relative", "http://editor.vbai.local/file", "https://external.example/file", "https://editor.vbai.local:444/file", "https://user@editor.vbai.local/file" }) Assert.IsFalse(ModernEditorWindow.LocalResource(uri));
            Assert.IsTrue(ModernEditorWindow.LocalResource("https://editor.vbai.local/asset.js"));
        }
        [STATestMethod]
        public void ProcessDocumentsHonorsInitialAndLateOwnedLifetimeGuards()
        {
            foreach (string state in new[] { "busy", "not-ready", "closing", "late-disposed", "late-closing", "snapshot-error" })
                using (var f = new Editor.ModernEditorToolFixture())
                {
                    if (state == "busy" || state == "closing") f.Base.Set(state, true); if (state == "not-ready") f.Base.Ready(false);
                    f.Override = (method, values) =>
                    {
                        if (method == "snapshots") { if (state == "late-disposed") f.Window.Dispose(); if (state == "late-closing") f.Base.Set("closing", true); if (state == "snapshot-error") throw new System.IO.IOException("owned snapshot error"); }
                        return null;
                    };
                    ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true)); Assert.AreEqual(0, f.Module.Writes);
                    if (state != "busy") Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
            foreach (string state in new[] { "disposed", "closing", "text", "baseline" })
                using (var f = new Editor.ModernEditorToolFixture())
                using (var entered = new System.Threading.ManualResetEventSlim())
                using (var release = new System.Threading.ManualResetEventSlim())
                {
                    var worker = new EditorSyncWorker(); f.Base.Set("synchronizationWorker", worker);
                    var blocked = worker.Evaluate(() => { entered.Set(); if (!release.Wait(5000)) throw new System.TimeoutException("owned worker gate"); return 1; }); Assert.IsTrue(entered.Wait(5000));
                    f.Base.Document.Edit(f.Base.Document.Text + "\n' queued draft");
                    var processing = f.Window.ProcessDocuments(true);
                    // Observe the actual immutable plan queue before releasing the worker.
                    var queue = (System.Collections.Concurrent.BlockingCollection<System.Action>)typeof(EditorSyncWorker).GetField("work", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(worker);
                    var limit = System.Diagnostics.Stopwatch.StartNew();
                    while (queue.Count == 0 && !processing.IsCompleted) { if (limit.ElapsedMilliseconds > 5000) throw new System.TimeoutException("Owned immutable plan was not queued"); System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                    Assert.AreEqual(1, queue.Count); Assert.IsFalse(processing.IsCompleted);
                    if (state == "disposed") f.Window.Dispose(); if (state == "closing") f.Base.Set("closing", true);
                    if (state == "text") f.Base.Document.Edit(f.Base.Document.Text + "\n' concurrent plan"); if (state == "baseline") f.Base.Document.Acknowledge("new baseline", "different captured revision");
                    release.Set(); ModernEditorDebugFixture.Wait(processing); Assert.AreEqual(0, f.Module.Writes); Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
        }
        [STATestMethod]
        public void CaptureRevisionsAndSynchronizationFailuresPreserveOwnedDrafts()
        {
            foreach (string state in new[] { "missing", "stale", "same", "newer" })
                using (var f = new Editor.ModernEditorToolFixture())
                {
                    string text = f.Document.Text + "\n' renderer";
                    f.Override = (method, values) => method == "snapshots" ? f.Json.Serialize(new[] { new { id = state == "missing" ? "missing" : f.Document.Id, version = state == "stale" ? 0 : state == "newer" ? 2 : 1, text } }) : null;
                    ModernEditorDebugFixture.Wait((System.Threading.Tasks.Task)f.Private("CaptureDocuments"));
                    Assert.AreEqual(state == "same" || state == "newer", f.Document.Text == text);
                }
            foreach (string state in new[] { "write-failure", "readonly", "conflict", "invalid-apply", "zero-apply" })
                using (var f = new Editor.ModernEditorToolFixture())
                {
                    f.Document.Edit(f.Document.Text + "\n' draft"); f.Module.Fail = state == "write-failure"; f.Module.CanWrite = state != "readonly";
                    if (state == "conflict") f.Module.Code += "\n' external";
                    if (state == "invalid-apply" || state == "zero-apply") { f.Document.AcceptRemote(f.Module.Code); f.Module.Code += "\n' external"; f.Override = (method, values) => method == "apply" ? state == "zero-apply" ? "0" : "invalid" : null; }
                    ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true)); Assert.AreEqual(0, f.Module.Writes); Assert.IsFalse(f.Base.Get<bool>("busy"));
                    if (state == "write-failure") Assert.IsTrue(UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text.Contains("Write refused"));
                }
        }
        [STATestMethod]
        public void OwnedStatusGenerationsAndDisposalGuardsSuppressObsoleteResults()
        {
            using (var f = new Editor.ModernEditorToolFixture())
            {
                f.Private("SetStatus"); f.Private("SetStatus"); f.Private("SetResultStatus", "owned result"); System.Windows.Forms.Application.DoEvents(); Assert.AreEqual("owned result", UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text);
                var module = new Editor.EditorLifetimeModule(); var doc = ModernEditorDebugFixture.Wait(f.Window.OpenModule(module)); module.FailName = true; f.Private("UpdateStatus"); Assert.AreEqual(doc, f.Window.Current);
                f.Base.Set("closing", true); f.Private("SetStatus"); f.Private("UpdateStatus"); f.Private("Report", new System.IO.IOException("owned closing error")); f.Base.Set("closing", false);
                var child = new System.Windows.Forms.Panel(); f.Window.Controls.Add(child);
                child.Disposed += (sender, args) => { Assert.IsTrue(f.Window.Disposing); f.Private("SetStatus"); f.Private("UpdateStatus"); f.Private("Report", new System.IO.IOException("owned disposing error")); };
                f.Window.Dispose(); f.Private("SetStatus"); f.Private("UpdateStatus"); f.Private("Report", new System.IO.IOException("owned disposed error"));
            }
            using (var window = new ModernEditorWindow())
            { UiInvoke.Call(typeof(ModernEditorWindow), "Dispose", window, false); Assert.IsFalse(window.IsDisposed); UiInvoke.Field<System.ComponentModel.IContainer>(window, "components").Dispose(); LlmBoundaryScope.Set(window, "components", null); window.Dispose(); Assert.IsTrue(window.IsDisposed); }
        }
        [STATestMethod]
        public void OwnedEditThemeDiffAndModuleClosurePropagateScriptAndQueuedTabFailures()
        {
            foreach (string operation in new[] { "EditClick", "ThemeChanged", "DiffClick", "CloseModuleClick" })
                foreach (bool failed in new[] { false, true })
                    using (var f = new Editor.ModernEditorToolFixture())
                    using (var dispatcher = new Editor.OwnedEditorDispatcher())
                    {
                        if (failed) f.Override = (method, values) => { if (method == "hideDiff" || method == "theme" || method == "compare" || method == "close") throw new System.IO.IOException("owned script refusal"); return null; };
                        if (operation == "ThemeChanged") f.Private(operation); else f.Private(operation, null, System.EventArgs.Empty); dispatcher.Drain();
                        if (failed) Assert.IsTrue(UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text.Contains("owned script refusal"));
                        if (operation == "CloseModuleClick" && !failed) Assert.IsFalse(f.Window.Documents.Contains(f.Document));
                    }
            using (var f = new Editor.ModernEditorToolFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                f.Base.Set("selected", f.Base.Document.Id);
                f.Base.Native.Original.CodeModule.CodePane.Window.OnClose = () => { throw new System.IO.IOException("owned pane closure failure"); };
                f.Private("CloseModuleClick", null, System.EventArgs.Empty); dispatcher.Drain(); Assert.IsTrue(f.Window.Documents.Contains(f.Base.Document)); Assert.IsFalse(f.Base.Get<bool>("busy"));
            }
        }
        [STATestMethod]
        public void OwnedReloadRestoreAndResolveHandleEveryGuardReplacementAndFailure()
        {
            foreach (string operation in new[] { "ReloadClick", "RestoreClick", "ResolveClick", "CloseModuleClick" })
                foreach (string state in new[] { "busy", "missing", "no-entry", "invalid", "zero", "error", "success" })
                    using (var f = new Editor.ModernEditorToolFixture())
                    using (var dispatcher = new Editor.OwnedEditorDispatcher())
                    {
                        if (state == "busy") f.Base.Set("busy", true); if (state == "missing") f.Base.Set("selected", null);
                        if (state != "no-entry")
                        {
                            f.Base.Get<System.Collections.Generic.Dictionary<string, string>>("reviewed")[f.Document.Id] = f.Module.Code;
                            f.Base.Get<System.Collections.Generic.Dictionary<string, EditorDraft>>("recovered")[f.Document.Id] = new EditorDraft { Key = f.Document.RecoveryKey, Baseline = f.Document.Baseline, Text = f.Document.Text + "\n' recovered" };
                        }
                        if (operation == "ReloadClick" && state == "success") f.Document.Edit(f.Document.Text + "\n' local");
                        f.Override = (method, values) => { if (method == "snapshots" && state == "error") throw new System.IO.IOException("owned event failure"); if (method == "apply" && (state == "invalid" || state == "zero")) return state == "zero" ? "0" : "invalid"; return null; };
                        f.Private(operation, null, System.EventArgs.Empty); dispatcher.Drain();
                        if (state == "error") Assert.IsTrue(UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text.Contains("owned event failure"));
                        if (operation == "RestoreClick" && state == "success") Assert.IsTrue(f.Document.Text.Contains("recovered"));
                        if (state != "busy") Assert.IsFalse(f.Base.Get<bool>("busy"));
                    }
        }
        [STATestMethod]
        public void OwnedClosingWaitsForEachActiveOperationAndPreservesDraftsEvenAfterCaptureFailure()
        {
            foreach (string state in new[] { "allowed", "idle", "already-closing", "capture-error", "not-ready", "initializing", "busy", "layout", "disposed" })
                using (var f = new Editor.ModernEditorToolFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    f.Document.Edit(f.Document.Text + "\n\u0027 owned recovery draft");
                    if (state == "allowed") f.Base.Set("closeAllowed", true);
                    if (state == "idle" || state == "not-ready" || state == "initializing" || state == "busy" || state == "layout") f.Base.Ready(false);
                    if (state == "already-closing") f.Base.Set("closing", true);
                    if (state == "initializing" || state == "busy") f.Base.Set(state, true); if (state == "layout") f.Base.Set("activeStatusLayouts", 1);
                    if (state == "not-ready") f.Base.Set("busy", true);
                    if (state == "capture-error") f.Override = (method, values) => method == "snapshots" ? throw new System.IO.IOException("owned close capture error") : (string)null;
                    if (state == "disposed") f.Override = (method, values) => { if (method == "snapshots") f.Window.Dispose(); return null; };
                    var args = new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false); f.Private("ClosingWindow", null, args);
                    Assert.AreEqual(state != "allowed" && state != "idle", args.Cancel, state);
                    if (state == "initializing" || state == "busy" || state == "layout" || state == "not-ready")
                    {
                        Assert.IsFalse(f.Window.IsDisposed);
                        f.Base.Set("initializing", false); f.Base.Set("busy", false); f.Base.Set("activeStatusLayouts", 0);
                    }
                    dispatcher.Drain();
                    if (state != "already-closing") Assert.IsNotNull(f.Window.Drafts.Recover(f.Document.RecoveryKey));
                }
            using (var f = new Editor.ModernEditorToolFixture())
            {
                var child = new System.Windows.Forms.Panel(); f.Window.Controls.Add(child); f.Window.ScriptExecution = null;
                child.Disposed += (sender, args) => { Assert.IsTrue(f.Window.Disposing); Assert.IsFalse(f.Window.IsDisposed); f.Private("ClosingWindow", null, new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false)); };
                f.Window.Dispose();
            }
        }
        [STATestMethod]
        public void OwnedTabsTimersRecoveryAndNativePaneClosureExerciseRealErrorBoundaries()
        {
            using (var f = new Editor.ModernEditorToolFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                f.Private("TabChanged", null, System.EventArgs.Empty); dispatcher.Drain();
                f.Override = (method, values) => method == "select" ? throw new System.IO.IOException("owned selection error") : (string)null;
                f.Private("TabChanged", null, System.EventArgs.Empty); dispatcher.Drain(); Assert.IsTrue(UiInvoke.Field<System.Windows.Forms.Label>(f.Window, "status").Text.Contains("owned selection error"));
                UiInvoke.Field<System.Windows.Forms.TabControl>(f.Window, "tabs").SelectedTab = null; f.Private("TabChanged", null, System.EventArgs.Empty); dispatcher.Drain();
            }
            foreach (string state in new[] { "busy", "closing", "plain", "late-closing", "late-disposed", "observation-error" })
                using (var f = new Editor.ModernEditorToolFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    if (state == "busy" || state == "closing") f.Base.Set(state, true);
                    if (state == "observation-error") { f.Base.Set("selected", f.Base.Document.Id); f.Base.Native.Vbe.MainWindow.HWnd = 0; }
                    f.Override = (method, values) => { if (method == "snapshots") { if (state == "late-closing") f.Base.Set("closing", true); if (state == "late-disposed") f.Window.Dispose(); } return null; };
                    f.Private("TimerTick", null, System.EventArgs.Empty); dispatcher.Drain();
                }
            using (var f = new Editor.ModernEditorToolFixture())
            {
                f.Window.Drafts = new EditorDraftStore(System.IO.Path.Combine(f.Module.Root, "blocked")); System.IO.Directory.CreateDirectory(f.Module.Root); System.IO.File.WriteAllText(f.Window.Drafts.Root, "owned blocker");
                f.Private("PreserveDrafts"); Assert.AreEqual("owned blocker", System.IO.File.ReadAllText(f.Window.Drafts.Root)); System.IO.File.Delete(f.Window.Drafts.Root);
            }
            using (var f = new Editor.ModernEditorToolFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                f.Base.Set("selected", f.Base.Document.Id); f.Private("CloseModuleClick", null, System.EventArgs.Empty); dispatcher.Drain(); Assert.IsFalse(f.Window.Documents.Contains(f.Base.Document));
            }
        }
        [STATestMethod]
        public void OwnedBrowserContractsInitializeSecurityPoliciesWithoutStartingChromium()
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                ModernEditorDebugFixture.Wait(f.Initialize()); Assert.AreEqual(ModernEditorWindow.Origin, f.Navigation, f.Diagnostic + string.Join("\n", f.Core.Errors.Concat(f.Controller.Errors).Concat(f.Settings.Errors))); Assert.AreEqual(0, f.Scripts);
                Assert.AreEqual(0, f.Settings.Values["AreDevToolsEnabled"]); Assert.AreEqual(0, f.Settings.Values["AreDefaultContextMenusEnabled"]); Assert.AreEqual(0, f.Settings.Values["IsStatusBarEnabled"]); Assert.AreEqual(0, f.Settings.Values["AreHostObjectsAllowed"]);
                Assert.IsTrue(f.Core.Handlers.ContainsKey("NavigationStarting")); f.Editor.Base.Ready(true); f.Window.ScriptExecution = null; Assert.AreEqual("null", ModernEditorDebugFixture.Wait(f.Window.Script("owned"))); Assert.AreEqual(1, f.Executes);
                Assert.IsTrue(f.Environment.Handlers.ContainsKey("BrowserProcessExited"), "Profile cleanup must subscribe to the owned environment's process exit.");
                var profile = f.Profiles.Single();
                Assert.AreEqual(Editor.ModernEditorBrowserFixture.BrowserProcessId, f.Core.Values["BrowserProcessId"]);
                f.BrowserExited(Editor.ModernEditorBrowserFixture.BrowserProcessId + 1);
                f.Window.Dispose();
                Assert.IsTrue(System.IO.Directory.Exists(profile.Path), "An unrelated process exit must not permit profile deletion.");
                f.BrowserExited(Editor.ModernEditorBrowserFixture.BrowserProcessId);
                ModernEditorDebugFixture.Wait(profile.Cleanup);
                Assert.IsFalse(System.IO.Directory.Exists(profile.Path), "The retired owned profile is removed only after its matching browser exits.");
            }
        }
        [STATestMethod]
        public void OwnedBrowserAwaitsHonorEveryInitialAndLateLifetimeGuardAndTranslationChoice()
        {
            foreach (string state in new[] { "initializing", "ready", "disposed", "closing", "missing-assets", "environment-disposed", "environment-closing", "core-disposed", "core-closing", "translation-disposed", "translation-closing", "environment-error", "error-disposed", "error-closing", "plain", "existing-browser" })
                using (var f = new Editor.ModernEditorBrowserFixture())
                {
                    if (state == "initializing" || state == "closing") f.Editor.Base.Set(state, true); if (state == "ready") f.Editor.Base.Ready(true); if (state == "disposed") f.Window.Dispose();
                    if (state == "missing-assets") System.IO.File.Delete(System.IO.Path.Combine(f.Assets, "index.html"));
                    f.EnvironmentCreated = () => { if (state == "environment-disposed" || state == "error-disposed") f.Window.Dispose(); if (state == "environment-closing" || state == "error-closing") f.Editor.Base.Set("closing", true); if (state.StartsWith("error-") || state == "environment-error") throw new System.IO.IOException("owned environment unavailable"); };
                    f.CoreEnsured = () => { if (state == "core-disposed") f.Window.Dispose(); if (state == "core-closing") f.Editor.Base.Set("closing", true); };
                    if (state.StartsWith("translation-"))
                    { string language = UiText.Culture.TwoLetterISOLanguageName; System.IO.File.WriteAllText(System.IO.Path.Combine(f.Assets, "nls.messages." + language + ".js"), "owned translation"); f.TranslationAdded = () => { if (state == "translation-disposed") f.Window.Dispose(); else f.Editor.Base.Set("closing", true); }; }
                    if (state == "existing-browser") ModernEditorDebugFixture.Wait(f.Initialize());
                    ModernEditorDebugFixture.Wait(f.Initialize());
                    if (state == "plain" || state == "existing-browser") Assert.AreEqual(ModernEditorWindow.Origin, f.Navigation);
                    else Assert.IsNull(f.Navigation, state);
                    if (state != "initializing") Assert.IsFalse(f.Editor.Base.Get<bool>("initializing"));
                }
            var culture = UiText.Culture;
            try
            {
                foreach (string language in new[] { "en-US", "pt-BR", "zh-CN" })
                    using (var f = new Editor.ModernEditorBrowserFixture())
                    {
                        LocalizationScope.Set(language);
                        string suffix = language == "en-US" ? "en" : language.ToLowerInvariant(); System.IO.File.WriteAllText(System.IO.Path.Combine(f.Assets, "nls.messages." + suffix + ".js"), "owned translation");
                        ModernEditorDebugFixture.Wait(f.Initialize()); Assert.AreEqual(1, f.Scripts, language);
                    }
            }
            finally { LocalizationScope.Set(culture.Name); }
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                var child = new System.Windows.Forms.Panel(); f.Window.Controls.Add(child);
                child.Disposed += (sender, args) => { Assert.IsTrue(f.Window.Disposing); Assert.IsFalse(f.Window.IsDisposed); ModernEditorDebugFixture.Wait(f.Initialize()); Assert.IsNull(f.Navigation); };
                f.Window.Dispose();
            }
            foreach (bool failed in new[] { false, true })
                using (var f = new Editor.ModernEditorBrowserFixture())
                {
                    var completion = new System.Threading.Tasks.TaskCompletionSource<Microsoft.Web.WebView2.Core.CoreWebView2Environment>();
                    f.Window.CreateBrowserEnvironment = cache => completion.Task; var task = f.Initialize();
                    var child = new System.Windows.Forms.Panel(); f.Window.Controls.Add(child);
                    child.Disposed += (sender, args) => { Assert.IsTrue(f.Window.Disposing); Assert.IsFalse(f.Window.IsDisposed); if (failed) completion.SetException(new System.IO.IOException("owned disposing environment")); else completion.SetResult(Editor.OwnedWebRaw.Wrap<Microsoft.Web.WebView2.Core.CoreWebView2Environment>(f.Environment.Proxy)); ModernEditorDebugFixture.Wait(task); };
                    f.Window.Dispose(); Assert.IsTrue(task.IsCompleted); Assert.IsNull(f.Navigation);
                }
        }
        [STATestMethod]
        public void OwnedBrowserEventsEnforceNavigationDownloadPermissionAndResourcePolicies()
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                ModernEditorDebugFixture.Wait(f.Initialize());
                foreach (string uri in new[] { ModernEditorWindow.Origin, "https://external.example" })
                {
                    var args = new Editor.OwnedWebRaw(Editor.OwnedWebRaw.Raw("ICoreWebView2NavigationStartingEventArgs")); args.Values["uri"] = uri; args.Values["Cancel"] = 0;
                    f.Core.Fire("NavigationStarting", f.Core.Proxy, args.Proxy); Assert.AreEqual(uri == ModernEditorWindow.Origin ? 0 : 1, args.Values["Cancel"]);
                    f.Core.Fire("FrameNavigationStarting", f.Core.Proxy, args.Proxy); Assert.AreEqual(1, args.Values["Cancel"]);
                }
                foreach (var pair in new[] { new[] { "NewWindowRequested", "ICoreWebView2NewWindowRequestedEventArgs", "Handled" }, new[] { "DownloadStarting", "ICoreWebView2DownloadStartingEventArgs", "Cancel" }, new[] { "PermissionRequested", "ICoreWebView2PermissionRequestedEventArgs", "State" } })
                {
                    var args = new Editor.OwnedWebRaw(Editor.OwnedWebRaw.Raw(pair[1])); f.Core.Fire(pair[0], f.Core.Proxy, args.Proxy);
                    if (pair[2] == "State") StringAssert.Contains(args.Values[pair[2]].ToString(), "DENY"); else Assert.AreEqual(1, args.Values[pair[2]]);
                }
                foreach (string uri in new[] { "https://editor.vbai.local/asset.js", "https://external.example/asset.js" })
                {
                    var request = new Editor.OwnedWebRaw(Editor.OwnedWebRaw.Raw("ICoreWebView2WebResourceRequest")); request.Values["uri"] = uri;
                    var args = new Editor.OwnedWebRaw(Editor.OwnedWebRaw.Raw("ICoreWebView2WebResourceRequestedEventArgs")); args.Values["Request"] = request.Proxy; f.CreatedResponse = null;
                    f.Core.Fire("WebResourceRequested", f.Core.Proxy, args.Proxy); Assert.AreEqual(!uri.Contains("editor.vbai.local"), f.CreatedResponse != null);
                    if (f.CreatedResponse != null) Assert.AreEqual(403, f.CreatedResponse.Values["StatusCode"]);
                }
                foreach (string state in new[] { "plain", "closing", "disposed" })
                {
                    if (state == "closing") f.Editor.Base.Set("closing", true); if (state == "disposed") { f.Editor.Base.Set("closing", false); f.Window.Dispose(); }
                    f.Editor.Base.Ready(true); var args = new Editor.OwnedWebRaw(Editor.OwnedWebRaw.Raw("ICoreWebView2ProcessFailedEventArgs"));
                    // Retain the subscribed managed handler after the owned controller is released.
                    f.Core.Fire("ProcessFailed", f.Core.Proxy, args.Proxy); Assert.AreEqual(state != "plain", f.Window.Ready);
                }
            }
        }
        [STATestMethod]
        public void OwnedWebMessagesValidateAuthoritySizeRevisionAndEverySupportedDispatch()
        {
            foreach (string state in new[] { "untrusted", "disposed", "closing", "oversized", "malformed", "ready", "ready-no-selection", "change-null", "change-missing", "change-stale", "change", "command-wrong", "command", "language", "definition", "editorCommand", "unknown" })
                using (var f = new Editor.ModernEditorBrowserFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    if (state == "disposed") f.Window.Dispose(); if (state == "closing") f.Editor.Base.Set("closing", true);
                    if (state == "ready-no-selection") f.Editor.Base.Set("selected", null);
                    string type = state.StartsWith("ready") ? "ready" : state.StartsWith("change") ? "change" : state.StartsWith("command") ? "command" : state;
                    string payload = f.Editor.Json.Serialize(new { type, id = state == "change-null" ? null : state == "change-missing" ? "missing" : f.Editor.Document.Id, version = state == "change-stale" ? 0 : 2, text = f.Editor.Document.Text + "\n' owned message", name = state == "command-wrong" ? "wrong" : state == "editorCommand" ? "unknown" : "sync", request = 42, module = f.Editor.Module.Name, line = 1, column = 1 });
                    if (state == "oversized") payload = new string('x', 16 * 1024 * 1024 + 1); if (state == "malformed") payload = "malformed";
                    f.Message(payload, state == "untrusted" ? "https://external.example" : ModernEditorWindow.Origin); dispatcher.Drain();
                    Assert.AreEqual(state == "ready" || state == "ready-no-selection", f.Window.Ready, state);
                    Assert.AreEqual(state == "change", f.Editor.Document.Dirty, state);
                    UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "timer").Stop();
                }
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                var child = new System.Windows.Forms.Panel(); f.Window.Controls.Add(child); child.Disposed += (sender, args) => { Assert.IsTrue(f.Window.Disposing); f.Message("{}"); }; f.Window.Dispose();
            }
            foreach (string state in new[] { "closing", "disposed", "disposing" })
                using (var f = new Editor.ModernEditorBrowserFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    f.Editor.Override = (method, values) => { if (method == "theme") { if (state == "closing") f.Editor.Base.Set("closing", true); if (state == "disposed") f.Window.Dispose(); } return null; };
                    if (state != "disposing") { f.Message("{\"type\":\"ready\"}"); dispatcher.Drain(); Assert.IsFalse(UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "timer").Enabled); }
                }
        }
        [STATestMethod]
        public void OwnedModuleLimitRecoveryAndUncreatedStatusPreserveEveryBoundary()
        {
            using (var window = new ModernEditorWindow())
            { LlmBoundaryScope.Call(window, "SetStatus"); Assert.IsFalse(window.IsHandleCreated); UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Dispose(); LlmBoundaryScope.Set(window, "timer", null); }
            using (var f = new Editor.ModernEditorToolFixture())
            {
                foreach (bool same in new[] { false, true })
                {
                    var module = new Editor.EditorLifetimeModule(); var draft = new EditorDraft { Key = module.Key, Baseline = "other baseline", Text = module.Code + (same ? "" : "\n' recovered draft") };
                    f.Window.Drafts.SaveSnapshot(System.Guid.NewGuid().ToString(), draft); f.Override = (method, values) => method == "open" ? "5" : null;
                    var doc = ModernEditorDebugFixture.Wait(f.Window.OpenModule(module)); Assert.AreEqual(!same, f.Base.Get<System.Collections.Generic.Dictionary<string, EditorDraft>>("recovered").ContainsKey(doc.Id)); Assert.AreEqual(5, f.Versions[doc.Id]);
                }
                while (f.Window.Documents.Count() < 30) ModernEditorDebugFixture.Wait(f.Window.OpenModule(new Editor.EditorLifetimeModule()));
                Assert.ThrowsException<System.InvalidOperationException>(() => ModernEditorDebugFixture.Wait(f.Window.OpenModule(new Editor.EditorLifetimeModule()))); Assert.AreEqual(30, f.Window.Documents.Count());
            }
        }
        [STATestMethod]
        public void UnexpectedOwnedLoggerFailureStillRunsAsynchronousCloseFinallyBeforePropagation()
        {
            using (var f = new Editor.ModernEditorToolFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                f.Document.Edit(f.Document.Text + "\n' recovery before failing logger");
                var append = LoadLog.AppendText;
                try
                {
                    LoadLog.AppendText = (path, text) => { if (text.Contains("Monaco:")) throw new System.InvalidOperationException("owned unexpected logger failure"); };
                    f.Override = (method, values) => method == "snapshots" ? throw new System.IO.IOException("owned capture failure") : (string)null;
                    f.Private("ClosingWindow", null, new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false)); dispatcher.Drain(typeof(System.InvalidOperationException));
                    Assert.IsTrue(f.Base.Get<bool>("closeAllowed")); Assert.IsTrue(f.Window.IsDisposed); Assert.IsNotNull(f.Window.Drafts.Recover(f.Document.RecoveryKey));
                }
                finally { LoadLog.AppendText = append; }
            }
        }
        [STATestMethod]
        public void NativeBrowserFactoriesUseOnlyAnOwnedTemporaryProfileAndLocalSyntheticDocument()
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            {
                string profile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "owned-vbai-browser-" + System.Guid.NewGuid().ToString("N"));
                var environmentFactory = typeof(ModernEditorWindow).GetMethod("NewBrowserEnvironment", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
                var ensure = typeof(ModernEditorWindow).GetMethod("EnsureBrowser", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
                f.Window.CreateBrowserEnvironment = cache => (System.Threading.Tasks.Task<Microsoft.Web.WebView2.Core.CoreWebView2Environment>)environmentFactory.Invoke(null, new object[] { profile });
                f.Window.EnsureBrowserEnvironment = (browser, environment) => (System.Threading.Tasks.Task)ensure.Invoke(null, new object[] { browser, environment });
                // The default assets path is observed with a guarded initialization; native mapping uses only the owned fixture folder.
                f.Window.BrowserAssetsDirectory = null;
                var nativeFactory = f.Window.CreateBrowserEnvironment; f.Window.CreateBrowserEnvironment = cache => { f.Editor.Base.Set("closing", true); return System.Threading.Tasks.Task.FromResult(Editor.OwnedWebRaw.Wrap<Microsoft.Web.WebView2.Core.CoreWebView2Environment>(f.Environment.Proxy)); };
                ModernEditorDebugFixture.Wait(f.Initialize()); f.Editor.Base.Set("closing", false); f.Window.CreateBrowserEnvironment = nativeFactory; f.Window.BrowserAssetsDirectory = f.Assets;
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                { f.Window.Show(); dispatcher.Drain(); }
                Assert.IsNotNull(f.Window.Browser.CoreWebView2); Assert.IsTrue(System.IO.Directory.Exists(profile)); Assert.AreEqual(f.Assets, f.Window.BrowserAssetsDirectory);
                f.Window.Dispose();
                // The dedicated profile is never reused. WebView2 can retain its owned process briefly after controller close.
                for (int attempt = 0; attempt < 100 && System.IO.Directory.Exists(profile); attempt++)
                { try { System.IO.Directory.Delete(profile, true); } catch (System.IO.IOException) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(15); } catch (System.UnauthorizedAccessException) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(15); } }
            }
        }
        [STATestMethod]
        public void Pr10AssistantActionsValidateNativeIdentitySelectionVersionAndOptionalCallback()
        {
            foreach (string state in new[] { "null-id", "missing", "managed", "unknown", "invalid-text", "invalid-selection", "stale", "no-callback", "/expliquer", "/corriger", "/refactoriser" })
                using (var f = new Editor.ModernEditorBrowserFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    var doc = f.Editor.Base.Document; int actions = 0; ChatAttachment attachment = null; string command = null;
                    if (state != "no-callback") f.Window.AssistantAction += (name, value) => { actions++; attachment = value; command = name; };
                    f.Message(f.Editor.Json.Serialize(new
                    {
                        type = "assistantAction",
                        id = state == "null-id" ? null : state == "missing" ? "missing" : state == "managed" ? f.Editor.Document.Id : doc.Id,
                        name = state == "unknown" ? "/unknown" : state.StartsWith("/") ? state : "/expliquer",
                        text = state == "invalid-text" ? null : doc.Text,
                        selectedText = state == "invalid-selection" ? null : "Debug.Print 1",
                        version = state == "stale" ? 0 : 1,
                        line = 3
                    }));
                    dispatcher.Drain(); bool valid = state.StartsWith("/"); Assert.AreEqual(valid ? 1 : 0, actions, state);
                    if (valid) { Assert.AreEqual(state, command); Assert.AreEqual("Debug.Print 1", attachment.Text); Assert.AreEqual("Project1", attachment.Project); Assert.AreEqual("Module1", attachment.Module); Assert.AreEqual(3, attachment.StartLine); Assert.AreEqual(doc.Id, attachment.EditorDocumentId); Assert.AreEqual(EditorDocument.Hash(doc.Text), attachment.Sha256); }
                    if (state.StartsWith("invalid") || state == "stale") Assert.IsFalse(string.IsNullOrEmpty(f.Editor.Base.Get<System.Windows.Forms.Label>("status").Text));
                }
        }

        [STATestMethod]
        public void Pr10ChangeClearsSavedAndSynchronizationErrorsAndSaveDispatchUsesOwnedFile()
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                var errors = f.Editor.Base.Get<System.Collections.Generic.Dictionary<string, string>>("documentSynchronizationErrors");
                f.Editor.Base.Set("lastSaveError", "previous save failure"); errors[f.Editor.Document.Id] = "previous sync failure";
                f.Message(f.Editor.Json.Serialize(new { type = "change", id = f.Editor.Document.Id, version = 2, text = f.Editor.Document.Text + "\n' new edit" })); dispatcher.Drain();
                Assert.IsNull(f.Editor.Base.Get<string>("lastSaveError")); Assert.IsFalse(errors.ContainsKey(f.Editor.Document.Id));
                f.Editor.Base.Ready(true); System.IO.Directory.CreateDirectory(f.Editor.Module.Root);
                f.Editor.Base.Native.Project.Path = System.IO.Path.Combine(f.Editor.Module.Root, "owned.bas"); System.IO.File.WriteAllText(f.Editor.Base.Native.Project.Path, "owned");
                int saves = 0; f.Window.NativeSave = native => saves++; f.Window.NativeHostSaved = native => true;
                f.Message(f.Editor.Json.Serialize(new { type = "command", name = "save", id = f.Editor.Base.Document.Id }));
                ModernEditorDebugFixture.Wait(System.Threading.Tasks.Task.Run(() =>
                {
                    var deadline = System.Diagnostics.Stopwatch.StartNew();
                    while (System.Threading.Volatile.Read(ref saves) == 0 || f.Editor.Base.Get<bool>("busy"))
                    {
                        if (deadline.ElapsedMilliseconds > 5000) throw new System.TimeoutException("Owned save command did not complete.");
                        System.Threading.Thread.Sleep(1);
                    }
                }));
                dispatcher.Drain();
                Assert.AreEqual(1, saves); Assert.AreEqual(UiText.Get("Saved."), f.Editor.Base.Get<System.Windows.Forms.Label>("status").Text);
                f.Editor.Base.Set("lastSaveError", "save has priority"); errors[f.Window.Current.Id] = "sync failure"; f.Editor.Private("UpdateStatus");
                Assert.AreEqual(UiText.Get("save has priority"), f.Editor.Base.Get<System.Windows.Forms.Label>("status").Text);
                f.Editor.Base.Set("lastSaveError", null); f.Editor.Private("UpdateStatus");
                Assert.AreEqual(UiText.Get("sync failure"), f.Editor.Base.Get<System.Windows.Forms.Label>("status").Text);
            }
        }

        [STATestMethod]
        public void Pr10OwnedBrowserKeysDispatchOnlyF9AndControlSaveAndReportScriptFailure()
        {
            foreach (var keys in new[] { System.Windows.Forms.Keys.F9, System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S, System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.F9, System.Windows.Forms.Keys.S, System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.S })
                using (var f = new Editor.ModernEditorBrowserFixture())
                using (var dispatcher = new Editor.OwnedEditorDispatcher())
                {
                    ModernEditorDebugFixture.Wait(f.Initialize()); f.Editor.Base.Ready(true); f.Editor.Base.Scripts.Clear();
                    var key = new System.Windows.Forms.KeyEventArgs(keys);
                    UiInvoke.Call(typeof(System.Windows.Forms.Control), "OnKeyDown", f.Window.Browser, key);
                    ModernEditorDebugFixture.Wait(System.Threading.Tasks.Task.Delay(30)); dispatcher.Drain();
                    bool handled = keys == System.Windows.Forms.Keys.F9 || keys == (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S);
                    Assert.AreEqual(handled, key.Handled); Assert.AreEqual(handled, key.SuppressKeyPress);
                    var commands = f.Editor.Base.Scripts.Where(item => item.Item1 == "command").ToArray(); Assert.AreEqual(handled ? 1 : 0, commands.Length);
                    if (handled) Assert.AreEqual(keys == System.Windows.Forms.Keys.F9 ? "vbai.toggle_breakpoint" : "vbai.save", commands[0].Item2[0]);
                }
            using (var f = new Editor.ModernEditorBrowserFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                ModernEditorDebugFixture.Wait(f.Initialize()); f.Editor.Base.Ready(true);
                f.Editor.Override = (method, values) => { if (method == "command") throw new System.IO.IOException("owned key command failed"); return null; };
                UiInvoke.Call(typeof(System.Windows.Forms.Control), "OnKeyDown", f.Window.Browser, new System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.F9));
                ModernEditorDebugFixture.Wait(System.Threading.Tasks.Task.Delay(30)); dispatcher.Drain();
                Assert.AreEqual(UiText.Get("owned key command failed"), f.Editor.Base.Get<System.Windows.Forms.Label>("status").Text);
            }
        }
        [STATestMethod]
        public void Pr10OwnedTabCloseRespectsBusySelectionAndVisibleToolbarLayout()
        {
            using (var f = new Editor.ModernEditorBrowserFixture())
            using (var dispatcher = new Editor.OwnedEditorDispatcher())
            {
                var tabs = f.Editor.Base.Get<ThemedTabControl>("tabs"); var nativeTab = tabs.TabPages[0];
                f.Editor.Base.Set("busy", true);
                f.Editor.Private("CloseTabRequested", null, new System.Windows.Forms.TabControlEventArgs(nativeTab, 0, System.Windows.Forms.TabControlAction.Selected));
                Assert.AreEqual(2, tabs.TabPages.Count);
                f.Editor.Base.Set("busy", false); f.Editor.Base.Set("initializing", true);
                f.Window.StartPosition = System.Windows.Forms.FormStartPosition.Manual; f.Window.Location = new System.Drawing.Point(-10000, -10000); f.Window.Show();
                f.Editor.Base.Set("showingDiff", true); f.Editor.Private("UpdateStatus"); Assert.IsTrue(f.Editor.Base.Get<System.Windows.Forms.TableLayoutPanel>("layout").RowStyles[0].Height >= f.Editor.Base.Get<System.Windows.Forms.FlowLayoutPanel>("toolbar").GetPreferredSize(System.Drawing.Size.Empty).Height);
                f.Editor.Base.Set("showingDiff", false); f.Editor.Private("UpdateStatus"); Assert.AreEqual(0f, f.Editor.Base.Get<System.Windows.Forms.TableLayoutPanel>("layout").RowStyles[0].Height);
                f.Editor.Private("CloseTabRequested", null, new System.Windows.Forms.TabControlEventArgs(nativeTab, 0, System.Windows.Forms.TabControlAction.Selected)); dispatcher.Drain();
                Assert.AreEqual(1, tabs.TabPages.Count); Assert.AreEqual(1, f.Editor.Base.Native.Original.CodeModule.CodePane.Window.Closes);
            }
        }
    }
}
