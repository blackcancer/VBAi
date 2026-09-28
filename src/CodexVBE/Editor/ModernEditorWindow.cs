using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace CodexVBE
{
    /// <summary>Designer shell with a local Monaco surface and revision-checked COM synchronization.</summary>
    internal sealed partial class ModernEditorWindow : Form
    {
        internal const string Origin = "https://editor.vbai.local/index.html";
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        private readonly Dictionary<string, EditorDocument> documents = new Dictionary<string, EditorDocument>();
        private readonly Dictionary<string, int> versions = new Dictionary<string, int>();
        private readonly Dictionary<string, EditorDraft> recovered = new Dictionary<string, EditorDraft>();
        private readonly Dictionary<string, string> reviewed = new Dictionary<string, string>();
        internal EditorDraftStore Drafts = new EditorDraftStore();
        internal WebView2 Browser { get; private set; }
        internal bool Ready { get; private set; }
        internal event Action<string, ChatAttachment> AssistantAction;
        internal bool WorkspaceHosted { get; set; }
        private bool busy, initializing, closing, closeAllowed, showingDiff;
        private int activeStatusLayouts;
        private string selected, synchronizationError;
        private DateTime lastEdit;
        private EditorSyncWorker synchronizationWorker;
        private Task<EditorSyncPlan> PrepareSynchronization(EditorDocument document)
        {
            if (synchronizationWorker == null) synchronizationWorker = new EditorSyncWorker();
            return synchronizationWorker.Prepare(document, Drafts);
        }


        internal IEnumerable<EditorDocument> Documents => documents.Values;
        internal EditorDocument Current => selected != null && documents.ContainsKey(selected) ? documents[selected] : null;
        public ModernEditorWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            Icon = VbeWindowIcons.Icon("assistant"); UiText.Apply(this, components); UiTheme.Attach(this);
            tabs.RightToLeft = RightToLeft.No;
            UiTheme.Changed += ThemeChanged;
        }
        protected override async void OnShown(EventArgs e) { base.OnShown(e); await InitializeBrowser(); }
        private async Task InitializeBrowser()
        {
            if (initializing || Ready || IsDisposed || Disposing || closing) return;
            initializing = true;
            try
            {
                string folder = Path.Combine(Path.GetDirectoryName(typeof(ModernEditorWindow).Assembly.Location), "EditorAssets");
                if (!File.Exists(Path.Combine(folder, "index.html"))) throw new FileNotFoundException("Monaco assets are missing.");
                Browser?.Dispose(); Browser = new WebView2 { Dock = DockStyle.Fill, DefaultBackgroundColor = UiTheme.Background };
                surface.Controls.Add(Browser);
                string cache = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "EditorWebView", System.Diagnostics.Process.GetCurrentProcess().Id.ToString());
                var environment = await CoreWebView2Environment.CreateAsync(null, cache);
                if (IsDisposed || Disposing || closing) return;
                await Browser.EnsureCoreWebView2Async(environment);
                if (IsDisposed || Disposing || closing) return;
                var core = Browser.CoreWebView2;
                string language = UiText.Culture.Name.ToLowerInvariant();
                if (language != "pt-br" && !language.StartsWith("zh-")) language = UiText.Culture.TwoLetterISOLanguageName;
                string translation = Path.Combine(folder, "nls.messages." + language + ".js");
                if (File.Exists(translation)) await core.AddScriptToExecuteOnDocumentCreatedAsync(File.ReadAllText(translation));
                if (IsDisposed || Disposing || closing) return;
                core.Settings.AreDevToolsEnabled = false; core.Settings.AreDefaultContextMenusEnabled = false;
                core.Settings.IsStatusBarEnabled = false; core.Settings.AreHostObjectsAllowed = false;
                core.SetVirtualHostNameToFolderMapping("editor.vbai.local", folder, CoreWebView2HostResourceAccessKind.DenyCors);
                core.NavigationStarting += (s, e) => { if (!Trusted(e.Uri)) e.Cancel = true; };
                core.FrameNavigationStarting += (s, e) => e.Cancel = true;
                core.NewWindowRequested += (s, e) => e.Handled = true;
                core.DownloadStarting += (s, e) => e.Cancel = true;
                core.PermissionRequested += (s, e) => e.State = CoreWebView2PermissionState.Deny;
                core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
                core.WebResourceRequested += (s, e) =>
                {
                    if (!LocalResource(e.Request.Uri)) e.Response = environment.CreateWebResourceResponse(new MemoryStream(), 403, "Forbidden", "Content-Type: text/plain");
                };
                Browser.KeyDown += (sender, key) =>
                {
                    if (key.KeyCode != Keys.F9 || key.Modifiers != Keys.None) return;
                    key.Handled = true; key.SuppressKeyPress = true;
                    BeginInvoke(new Action(async () => { try { await Script("command", "vbai.toggle_breakpoint"); } catch (Exception error) { Report(error); } }));
                };
                core.WebMessageReceived += MessageReceived;
                core.ProcessFailed += (s, e) => { if (IsDisposed || Disposing || closing) return; Ready = false; timer.Stop(); PreserveDrafts(); status.Text = UiText.Get("The editor stopped. Drafts are preserved; reopen the editor."); };
                core.Navigate(Origin);
                status.Text = UiText.Get("Loading editor…");
            }
            catch (Exception error) { LoadLog.Write("Monaco initialization: " + error.GetType().Name); if (!IsDisposed && !Disposing && !closing) status.Text = UiText.Get("Editor unavailable. Install WebView2 Runtime or use the native editor."); }
            finally { initializing = false; }
        }
        internal static bool Trusted(string uri) => string.Equals(uri, Origin, StringComparison.Ordinal);
        internal static bool LocalResource(string uri)
        { return Uri.TryCreate(uri, UriKind.Absolute, out var u) && u.Scheme == "https" && u.Host == "editor.vbai.local" && u.IsDefaultPort && u.UserInfo.Length == 0; }
        private async void MessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            if (!Trusted(args.Source) || IsDisposed || Disposing || closing) return;
            try
            {
                if (args.WebMessageAsJson.Length > 16 * 1024 * 1024) return;
                var message = json.Deserialize<EditorMessage>(args.WebMessageAsJson);
                if (message.type == "ready")
                {
                    Ready = true; await Theme();
                    await Script("labels", new[] { "Compile project", "Toggle breakpoint", "Show next statement", "Step into", "Step over", "Step out", "Breakpoint request sent; verify in VBE.", "Explain", "Fix", "Refactor" }.ToDictionary(key => key, UiText.Get));
                    foreach (var document in documents.Values.ToArray()) await RenderDocument(document);
                    if (selected != null) await Script("select", selected);
                    if (!closing && !IsDisposed && !Disposing) { timer.Start(); SetStatus(); }
                }
                else if (message.type == "change" && documents.TryGetValue(message.id ?? "", out var doc) && message.version > versions[doc.Id])
                { doc.Edit(message.text); versions[doc.Id] = message.version; lastEdit = DateTime.UtcNow; synchronizationError = null; SetStatus(); }
                else if (message.type == "command" && message.name == "sync") await ProcessDocuments(true);
                else if (message.type == "language") await LanguageRequest(message);
                else if (message.type == "definition") await OpenDefinition(message);
                else if (message.type == "assistantAction" && documents.TryGetValue(message.id ?? "", out var actionDoc) && actionDoc.Module is EditorVbeModule actionModule)
                {
                    if (!new[] { "/expliquer", "/corriger", "/refactoriser" }.Contains(message.name)) return;
                    EditorDocument.Validate(message.text); EditorDocument.Validate(message.selectedText);
                    if (message.version < versions[actionDoc.Id]) throw new InvalidOperationException("The editor selection changed. Select it again.");
                    actionDoc.Edit(message.text); versions[actionDoc.Id] = message.version;
                    await Task.Yield();
                    AssistantAction?.Invoke(message.name, new ChatAttachment { Label = actionModule.Name + " · Monaco L" + message.line,
                        Project = actionModule.ProjectName, Module = actionModule.ModuleName, Text = message.selectedText,
                        StartLine = message.line, EditorDocumentId = actionDoc.Id, Sha256 = EditorDocument.Hash(message.text) });
                }
                else if (message.type == "editorCommand") await EditorCommand(message);
            }
            catch (Exception error) { Report(error); }
        }
        private sealed class EditorMessage { public string type { get; set; } public string id { get; set; } public string text { get; set; } public string selectedText { get; set; } public int version { get; set; } public string name { get; set; } public int request { get; set; } public string module { get; set; } public int line { get; set; } public int column { get; set; } }
        internal async Task<string> Script(string method, params object[] values)
        {
            if (!Ready || Browser?.CoreWebView2 == null || IsDisposed) return "null";
            // Method names are internal constants; all document text is serialized as data.
            string result = await Browser.CoreWebView2.ExecuteScriptAsync("window.vbai." + method + "(" + string.Join(",", values.Select(json.Serialize)) + ")");
            // WebView completes script tasks inside its COM callback. Unwind that callback
            // before callers resize controls, change tabs, activate windows or close them.
            await Task.Yield();
            return result;
        }
        internal async Task<EditorDocument> OpenModule(IEditorModule module)
        {
            var existing = documents.Values.FirstOrDefault(d => ReferenceEquals(d.Module, module) ||
                (d.Module is EditorVbeModule vm && module is EditorVbeModule other && vm.IsComponent(other.Component)));
            if (existing != null) { selected = existing.Id; SelectTab(existing.Id); if (Ready) await Script("select", existing.Id); return existing; }
            if (documents.Count >= 30) throw new InvalidOperationException("Close the editor before opening more than 30 modules.");
            var document = new EditorDocument(module);
            if (module is EditorVbeModule nativeModule) nativeModule.EnsureNativeWindow();
            var draft = Drafts.Recover(module.Key);
            documents.Add(document.Id, document); versions[document.Id] = 1;
            if (draft != null && EditorDocument.Normalize(draft.Text) != document.Text) recovered[document.Id] = draft;
            var tab = new TabPage(module.Name) { Tag = document.Id }; tabs.TabPages.Add(tab); selected = document.Id; tabs.SelectedTab = tab;
            if (Ready) await RenderDocument(document);
            SetStatus(); Activate(); return document;
        }
        private async Task RenderDocument(EditorDocument doc)
        { int version; if (int.TryParse(await Script("open", doc.Id, doc.Text), out version)) versions[doc.Id] = version; }
        private void SelectTab(string id) { foreach (TabPage tab in tabs.TabPages) if ((string)tab.Tag == id) { tabs.SelectedTab = tab; break; } }
        private async void TabChanged(object sender, EventArgs e)
        { if (tabs.SelectedTab == null) return; selected = (string)tabs.SelectedTab.Tag; try { if (Ready) await Script("select", selected); SetStatus(); } catch (Exception error) { Report(error); } }
        private async Task CaptureDocuments()
        {
            var snapshots = json.Deserialize<EditorMessage[]>(await Script("snapshots"));
            if (snapshots == null) return;
            foreach (var item in snapshots)
                if (documents.TryGetValue(item.id, out var doc) && item.version >= versions[item.id])
                { doc.Edit(item.text); versions[item.id] = item.version; }
        }
        internal async Task ProcessDocuments(bool synchronize)
        {
            if (busy || !Ready || closing) return;
            busy = true;
            try { await ProcessDocumentsCore(synchronize); }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        private async Task ProcessDocumentsCore(bool synchronize)
        {
            await CaptureDocuments();
            if (IsDisposed || closing) return;
            Exception lastFailure = null;
            foreach (var doc in documents.Values.ToArray())
            {
                try
                {
                    var plan = await PrepareSynchronization(doc);
                    if (IsDisposed || closing) return;
                    if (doc.Text != plan.After || doc.Baseline != plan.Before) continue;
                    string captured = doc.Text;
                    int capturedVersion = versions[doc.Id];
                    string native = doc.Observe();
                    if (synchronize && doc.Dirty && !doc.Conflict && doc.Writable) native = doc.Synchronize(plan);
                    if (native != null)
                    {
                        int applied;
                        if (int.TryParse(await Script("apply", doc.Id, capturedVersion, native), out applied) && applied > 0)
                        { doc.Acknowledge(native, captured); versions[doc.Id] = Math.Max(versions[doc.Id], applied); }
                    }
                    if (!doc.Dirty) Drafts.ClearOwn(doc);
                }
                catch (Exception error) { lastFailure = error; }
            }
            synchronizationError = lastFailure?.Message;
            SetStatus();
            if (lastFailure != null) Report(lastFailure);
        }
        private async void TimerTick(object sender, EventArgs e)
        {
            if (busy || closing) return;
            await ProcessDocuments(DateTime.UtcNow - lastEdit > TimeSpan.FromMilliseconds(600));
            try { if (!closing && !IsDisposed) await ObserveDebugMode(); } catch (Exception error) { LoadLog.Write("Monaco debug observation: " + error.Message); }
        }
        private void SetStatus()
        {
            if (IsDisposed || Disposing || closing) return;
            // WebView callbacks must unwind before changing WinForms visibility/layout.
            if (IsHandleCreated) BeginInvoke(new Action(UpdateStatus)); else UpdateStatus();
        }
        private void UpdateStatus()
        {
            if (IsDisposed || Disposing || closing) return;
            activeStatusLayouts++;
            try
            {
            toolbar.Visible = showingDiff || (Current != null && (Current.Conflict || recovered.ContainsKey(Current.Id)));
            layout.RowStyles[0].Height = toolbar.Visible ? 44 : 0;
            compare.Visible = Current != null && Current.Conflict;
            edit.Visible = showingDiff; reload.Visible = Current != null && Current.Conflict;
            resolve.Visible = Current != null && Current.Conflict;
            restore.Visible = Current != null && recovered.ContainsKey(Current.Id);
            resolve.Enabled = Current != null && Current.Conflict && reviewed.ContainsKey(Current.Id);
            restore.Enabled = Current != null && recovered.ContainsKey(Current.Id);
            foreach (TabPage tab in tabs.TabPages)
            { var doc = documents[(string)tab.Tag]; try { tab.Text = doc.Module.Name + (doc.Dirty ? " *" : ""); } catch { } }
            status.Text = UiText.Get(synchronizationError ?? (Current == null ? "Open a VBA module to start editing." : Current.Conflict ? "The module changed in VBA. Resolve the conflict first." : Current.Dirty ? "Changes pending synchronization with VBA." : "Synchronized with VBA. Save the macro in its host application."));
            }
            finally { activeStatusLayouts--; }
        }
        private void Report(Exception error) { if (!IsDisposed && !Disposing && !closing) status.Text = UiText.Get(error.Message); LoadLog.Write("Monaco: " + error.GetType().Name); }
        private void CloseTabRequested(object sender, TabControlEventArgs e)
        {
            if (busy) return;
            tabs.SelectedTab = e.TabPage;
            CloseModuleClick(sender, EventArgs.Empty);
        }
        private async void DiffClick(object sender, EventArgs e)
        {
            try
            {
                await ProcessDocuments(false);
                if (Current != null)
                {
                    reviewed[Current.Id] = Current.Native;
                    await Script("compare", Current.Native);
                    showingDiff = true; SetStatus();
                }
            }
            catch (Exception error) { Report(error); }
        }
        private async void ResolveClick(object sender, EventArgs e)
        {
            if (busy || Current == null || !reviewed.TryGetValue(Current.Id, out var revision)) return;
            var doc = Current;
            busy = true;
            try
            {
                await CaptureDocuments(); Drafts.Save(doc);
                string captured = doc.Text, actual = doc.ResolveWithDraft(revision);
                int applied;
                if (int.TryParse(await Script("apply", doc.Id, versions[doc.Id], actual), out applied) && applied > 0)
                { doc.Acknowledge(actual, captured); versions[doc.Id] = Math.Max(versions[doc.Id], applied); }
                reviewed.Remove(doc.Id); await Script("hideDiff"); showingDiff = false; SetStatus();
            }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        private async void CloseModuleClick(object sender, EventArgs e)
        {
            if (busy || Current == null) return;
            var doc = Current;
            busy = true;
            try
            {
                await CaptureDocuments(); Drafts.Save(doc);
                await Script("close", doc.Id);
                var closed = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                BeginInvoke(new Action(() =>
                {
                    try
                    {
                        (doc.Module as EditorVbeModule)?.CloseNativeWindow();
                        var page = tabs.TabPages.Cast<TabPage>().First(t => (string)t.Tag == doc.Id);
                        if (selected == doc.Id) selected = null; tabs.TabPages.Remove(page); page.Dispose();
                        if (tabs.SelectedTab != null) selected = (string)tabs.SelectedTab.Tag;
                        documents.Remove(doc.Id); versions.Remove(doc.Id); reviewed.Remove(doc.Id); recovered.Remove(doc.Id);
                        SetStatus(); closed.SetResult(true);
                    }
                    catch (Exception error) { closed.SetException(error); }
                }));
                await closed.Task;
            }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        private async void EditClick(object sender, EventArgs e) { try { await Script("hideDiff"); showingDiff = false; SetStatus(); } catch (Exception error) { Report(error); } }
        private async void ReloadClick(object sender, EventArgs e)
        {
            if (busy || Current == null) return;
            var doc = Current;
            busy = true;
            try
            {
                await CaptureDocuments();
                if (doc.Dirty)
                {
                    new EditorDraftStore(Drafts.Root).Save(doc);
                    recovered[doc.Id] = new EditorDraft { Key = doc.RecoveryKey, Baseline = doc.Baseline, Text = doc.Text };
                }
                string text = EditorDocument.Normalize(doc.Module.Read());
                int applied; if (int.TryParse(await Script("apply", doc.Id, versions[doc.Id], text), out applied) && applied > 0)
                { doc.AcceptRemote(text); versions[doc.Id] = applied; }
                SetStatus();
            }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        private async void RestoreClick(object sender, EventArgs e)
        {
            if (busy || Current == null || !recovered.TryGetValue(Current.Id, out var draft)) return;
            var doc = Current;
            busy = true;
            try
            {
                await CaptureDocuments(); Drafts.Save(doc); int applied;
                if (int.TryParse(await Script("apply", doc.Id, versions[doc.Id], EditorDocument.Normalize(draft.Text)), out applied) && applied > 0)
                { doc.Restore(draft.Baseline, draft.Text); versions[doc.Id] = applied; recovered.Remove(doc.Id); }
                SetStatus();
            }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        private async void ThemeChanged() { try { await Theme(); } catch (Exception error) { Report(error); } }
        private Task<string> Theme() => Script("theme", UiTheme.Dark, UiTheme.HighContrast());
        private void PreserveDrafts() { foreach (var doc in documents.Values) try { Drafts.Save(doc); } catch (Exception error) { LoadLog.Write("Editor recovery failed: " + error.GetType().Name); } }
        private async void ClosingWindow(object sender, FormClosingEventArgs e)
        {
            if (WorkspaceHosted && e.CloseReason == CloseReason.UserClosing && !closeAllowed) { e.Cancel = true; return; }
            if (closeAllowed) { PreserveDrafts(); return; }
            if (!Ready && !initializing && !busy && activeStatusLayouts == 0) { closing = true; PreserveDrafts(); return; }
            e.Cancel = true; if (closing) return; closing = true; timer.Stop();
            try { if (Ready) await CaptureDocuments(); PreserveDrafts(); }
            catch (Exception error) { Report(error); PreserveDrafts(); }
            finally
            {
                // A WebView/WinForms callback can pump WM_CLOSE during a control's
                // CreateHandle. Let initialization, synchronization and status layout
                // unwind before allowing Form.Dispose to destroy their controls.
                while (!IsDisposed && (initializing || busy || activeStatusLayouts > 0)) await Task.Delay(15);
                closeAllowed = true;
                if (!IsDisposed && !Disposing) BeginInvoke(new Action(Close));
            }
        }
        private void DisposeRuntime()
        {
            timer?.Stop(); PreserveDrafts(); synchronizationWorker?.Dispose(); UiTheme.Changed -= ThemeChanged; Browser?.Dispose();
            foreach (var doc in documents.Values) (doc.Module as EditorVbeModule)?.CloseNativeWindow();
        }
    }
}
