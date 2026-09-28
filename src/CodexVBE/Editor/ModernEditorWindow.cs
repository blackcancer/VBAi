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
        /// <summary>Origine locale virtuelle autorisée pour l’interface Monaco.</summary>
        internal const string Origin = "https://editor.vbai.local/index.html";
        /// <summary>Sérialiseur des messages JSON entre WebView2 et le code hôte.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        /// <summary>Documents ouverts, indexés par leur identifiant de session.</summary>
        private readonly Dictionary<string, EditorDocument> documents = new Dictionary<string, EditorDocument>();
        /// <summary>Dernières révisions Monaco observées pour chaque document.</summary>
        private readonly Dictionary<string, int> versions = new Dictionary<string, int>();
        /// <summary>Brouillons récupérés au chargement et proposés séparément du code natif.</summary>
        private readonly Dictionary<string, EditorDraft> recovered = new Dictionary<string, EditorDraft>();
        /// <summary>Versions natives comparées par l’utilisateur avant résolution de conflit.</summary>
        private readonly Dictionary<string, string> reviewed = new Dictionary<string, string>();
        /// <summary>Magasin chiffré des brouillons de cette fenêtre.</summary>
        internal EditorDraftStore Drafts = new EditorDraftStore();
        /// <summary>Obtient la vue WebView2 qui affiche Monaco.</summary>
        /// <value>Contrôle de navigateur, nul avant son initialisation.</value>
        internal WebView2 Browser { get; private set; }
        /// <summary>Stores the create browser used by ModernEditorWindow.</summary>
        internal Func<WebView2> CreateBrowser = NewBrowser;
        /// <summary>Stores the create browser environment used by ModernEditorWindow.</summary>
        internal Func<string, Task<CoreWebView2Environment>> CreateBrowserEnvironment = NewBrowserEnvironment;
        /// <summary>Stores the ensure browser environment used by ModernEditorWindow.</summary>
        internal Func<WebView2, CoreWebView2Environment, Task> EnsureBrowserEnvironment = EnsureBrowser;
        /// <summary>Stores the browser assets directory used by ModernEditorWindow.</summary>
        internal string BrowserAssetsDirectory;
        /// <summary>Performs the new browser operation for ModernEditorWindow.</summary>
        /// <returns>The result produced by this operation.</returns>
        private static WebView2 NewBrowser() => new WebView2 { Dock = DockStyle.Fill, DefaultBackgroundColor = UiTheme.Background };
        /// <summary>Performs the new browser environment operation for ModernEditorWindow.</summary>
        /// <param name="cache">Text containing the cache.</param>
        /// <returns>The result produced by this operation.</returns>
        private static Task<CoreWebView2Environment> NewBrowserEnvironment(string cache) => CoreWebView2Environment.CreateAsync(null, cache);
        /// <summary>Performs the ensure browser operation for ModernEditorWindow.</summary>
        /// <param name="browser">The browser used by this operation.</param>
        /// <param name="environment">The environment used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        private static Task EnsureBrowser(WebView2 browser, CoreWebView2Environment environment) => browser.EnsureCoreWebView2Async(environment);
        /// <summary>Indique que l’application Monaco a signalé être prête.</summary>
        /// <value><see langword="true"/> après la réception du message ready.</value>
        internal bool Ready { get; private set; }
        /// <summary>Optional renderer boundary for an embedded surface or an isolated contract host.</summary>
        internal Func<string, object[], Task<string>> ScriptExecution;
        /// <summary>Indique qu’une opération asynchrone ou un état transitoire interdit une autre opération.</summary>
        private bool busy, initializing, closing, closeAllowed, showingDiff;
        /// <summary>Nombre de mises à jour d’état qui manipulent actuellement la disposition des contrôles.</summary>
        private int activeStatusLayouts;
        /// <summary>Stores the status generation used by ModernEditorWindow.</summary>
        private int statusGeneration;
                /// <summary>Notifies subscribers when assistant action occurs.</summary>
                internal event Action<string, ChatAttachment> AssistantAction;
        /// <summary>Gets or sets the workspace hosted.</summary>
        /// <value>The current value represented by this member.</value>
        internal bool WorkspaceHosted { get; set; }
        /// <summary>Identifiant du document sélectionné dans les onglets.</summary>
        private string selected, synchronizationError;
        /// <summary>Heure du dernier changement de texte, utilisée pour différer la synchronisation automatique.</summary>
        private DateTime lastEdit;
        /// <summary>Worker dédié à la préparation des instantanés et des diffs.</summary>
        private EditorSyncWorker synchronizationWorker;
        /// <summary>Persiste le document puis calcule en arrière-plan son plan de synchronisation.</summary>
        /// <param name="document">Document dont l’état doit être capturé.</param>
        /// <returns>Tâche qui produit un plan lié aux textes capturés.</returns>
        private Task<EditorSyncPlan> PrepareSynchronization(EditorDocument document)
        {
            if (synchronizationWorker == null) synchronizationWorker = new EditorSyncWorker();
            return synchronizationWorker.Prepare(document, Drafts);
        }


        /// <summary>Obtient les documents actuellement ouverts.</summary>
        /// <value>Vue des documents présents dans la fenêtre.</value>
        internal IEnumerable<EditorDocument> Documents => documents.Values;
        /// <summary>Obtient le document sélectionné.</summary>
        /// <value>Document actif ou <see langword="null"/> si aucun onglet n’est sélectionné.</value>
        internal EditorDocument Current => selected != null && documents.ContainsKey(selected) ? documents[selected] : null;
        /// <summary>Crée la fenêtre et relie son thème et ses contrôles localisés.</summary>
        public ModernEditorWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            Icon = VbeWindowIcons.Icon("assistant"); UiText.Apply(this, components); UiTheme.Attach(this);
            tabs.RightToLeft = RightToLeft.No;
            UiTheme.Changed += ThemeChanged;
        }
        /// <summary>Initialise WebView2 lorsque la fenêtre devient visible.</summary>
        /// <param name="e">Données de l’événement d’affichage.</param>
        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            await InitializeBrowser();
        }
        /// <summary>Configure WebView2 avec les ressources locales, les permissions refusées et les points d’entrée de messages.</summary>
        /// <returns>Tâche terminée lorsque la navigation locale a été lancée ou que l’initialisation a échoué.</returns>
        private async Task InitializeBrowser()
        {
            if (initializing || Ready || IsDisposed || Disposing || closing) return;
            initializing = true;
            try
            {
                string folder = BrowserAssetsDirectory ?? Path.Combine(Path.GetDirectoryName(typeof(ModernEditorWindow).Assembly.Location), "EditorAssets");
                if (!File.Exists(Path.Combine(folder, "index.html"))) throw new FileNotFoundException("Monaco assets are missing.");
                Browser?.Dispose(); Browser = CreateBrowser();
                surface.Controls.Add(Browser);
                string cache = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "EditorWebView", System.Diagnostics.Process.GetCurrentProcess().Id.ToString());
                var environment = await CreateBrowserEnvironment(cache);
                if (IsDisposed || Disposing || closing) return;
                await EnsureBrowserEnvironment(Browser, environment);
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
                    string action = key.KeyCode == Keys.F9 && key.Modifiers == Keys.None ? "vbai.toggle_breakpoint" :
                        key.KeyCode == Keys.S && key.Modifiers == Keys.Control ? "vbai.save" : null;
                    if (action == null) return;
                    key.Handled = true; key.SuppressKeyPress = true;
                    BeginInvoke(new Action(async () => { try { await Script("command", action); } catch (Exception error) { Report(error); } }));
                };
                core.WebMessageReceived += MessageReceived;
                core.ProcessFailed += (s, e) => { if (IsDisposed || Disposing || closing) return; Ready = false; timer.Stop(); PreserveDrafts(); status.Text = UiText.Get("The editor stopped. Drafts are preserved; reopen the editor."); };
                core.Navigate(Origin);
                status.Text = UiText.Get("Loading editor…");
            }
            catch (Exception error) { LoadLog.Write("Monaco initialization: " + error.GetType().Name); if (!IsDisposed && !Disposing && !closing) status.Text = UiText.Get("Editor unavailable. Install WebView2 Runtime or use the native editor."); }
            finally { initializing = false; }
        }
        /// <summary>Indique si l’URI correspond exactement à la page d’édition locale autorisée.</summary>
        /// <param name="uri">URI de navigation ou de message à vérifier.</param>
        /// <returns><see langword="true"/> uniquement pour l’origine locale configurée.</returns>
        internal static bool Trusted(string uri) => string.Equals(uri, Origin, StringComparison.Ordinal);
        /// <summary>Vérifie qu’une ressource appartient à l’hôte virtuel HTTPS local sans identifiants ni port non standard.</summary>
        /// <param name="uri">URI de ressource demandée.</param>
        /// <returns><see langword="true"/> si la ressource peut provenir du dossier Monaco local.</returns>
        internal static bool LocalResource(string uri)
        { return Uri.TryCreate(uri, UriKind.Absolute, out var u) && u.Scheme == "https" && u.Host == "editor.vbai.local" && u.IsDefaultPort && u.UserInfo.Length == 0; }
        /// <summary>Valide et distribue les messages JSON reçus depuis la page Monaco locale.</summary>
        /// <param name="sender">WebView2 à l’origine de l’événement.</param>
        /// <param name="args">Origine et contenu du message reçu.</param>
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
                { doc.Edit(message.text); versions[doc.Id] = message.version; lastEdit = DateTime.UtcNow; synchronizationError = null; lastSaveError = null; SetStatus(); }
                else if (message.type == "command" && message.name == "sync") await ProcessDocuments(true);
                else if (message.type == "command" && message.name == "save") await SaveDocument(message.id);
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
        /// <summary>Message typé échangé entre l’interface Monaco et la fenêtre hôte.</summary>
        private sealed class EditorMessage { /// <summary>Type d’opération demandée.</summary>
            /// <value>Message de changement, commande, navigation ou réponse.</value>
public string type { get; set; } /// <summary>Identifiant du document concerné.</summary>
            /// <value>Identifiant de session du document Monaco.</value>
public string id { get; set; } /// <summary>Texte transmis avec un changement de brouillon.</summary>
            /// <value>Contenu source envoyé par Monaco.</value>
public string text { get; set; } /// <summary>Gets or sets the selected text.</summary>
/// <value>The current value represented by this member.</value>
public string selectedText { get; set; } /// <summary>Révision Monaco associée au message.</summary>
            /// <value>Numéro de version du document.</value>
public int version { get; set; } /// <summary>Nom de commande d’éditeur ou de débogage.</summary>
            /// <value>Commande interne, par exemple compile ou step_into.</value>
public string name { get; set; } /// <summary>Identifiant de la requête de langage à laquelle répondre.</summary>
            /// <value>Numéro de requête généré par Monaco.</value>
public int request { get; set; } /// <summary>Nom du module cible d’une navigation vers définition.</summary>
            /// <value>Nom du composant cible.</value>
public string module { get; set; } /// <summary>Ligne de navigation ou de sélection.</summary>
            /// <value>Numéro de ligne indexé à partir de un.</value>
public int line { get; set; } /// <summary>Colonne de navigation ou de sélection.</summary>
            /// <value>Numéro de colonne indexé à partir de un.</value>
public int column { get; set; } }
        /// <summary>Appelle une méthode de l’interface Monaco avec des arguments sérialisés en données JSON.</summary>
        /// <param name="method">Nom de méthode interne exposée par l’application Web.</param>
        /// <param name="values">Arguments sérialisés individuellement avant l’appel.</param>
        /// <returns>Résultat JSON renvoyé par WebView2, ou chaîne null si la surface n’est pas prête.</returns>
        internal async Task<string> Script(string method, params object[] values)
        {
            if (ScriptExecution != null)
            {
                string rendered = await ScriptExecution(method, values);
                await Task.Yield();
                return rendered;
            }
            if (!Ready || Browser?.CoreWebView2 == null || IsDisposed) return "null";
            // Method names are internal constants; all document text is serialized as data.
            string result = await Browser.CoreWebView2.ExecuteScriptAsync("window.vbai." + method + "(" + string.Join(",", values.Select(json.Serialize)) + ")");
            // WebView completes script tasks inside its COM callback. Unwind that callback
            // before callers resize controls, change tabs, activate windows or close them.
            await Task.Yield();
            return result;
        }
        /// <summary>Ouvre ou sélectionne le document correspondant au module fourni et restaure tout brouillon récupéré.</summary>
        /// <param name="module">Adaptateur du module source.</param>
        /// <returns>Document créé ou déjà ouvert.</returns>
        /// <exception cref="InvalidOperationException">La limite de documents ouverts est atteinte.</exception>
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
        /// <summary>Envoie le contenu d’un document nouvellement ouvert à Monaco et actualise sa révision.</summary>
        /// <param name="doc">Document à afficher.</param>
        /// <returns>Tâche terminée après la réponse de la page.</returns>
        private async Task RenderDocument(EditorDocument doc)
        { int version; if (int.TryParse(await Script("open", doc.Id, doc.Text), out version)) versions[doc.Id] = version; }
        /// <summary>Sélectionne l’onglet portant l’identifiant du document.</summary>
        /// <param name="id">Identifiant de session du document.</param>
        private void SelectTab(string id) { foreach (TabPage tab in tabs.TabPages) if ((string)tab.Tag == id) { tabs.SelectedTab = tab; break; } }
        /// <summary>Met à jour le document courant et demande à Monaco de sélectionner son onglet.</summary>
        /// <param name="sender">Onglets à l’origine de l’événement.</param>
        /// <param name="e">Données de sélection.</param>
        private async void TabChanged(object sender, EventArgs e)
        { if (tabs.SelectedTab == null) return; selected = (string)tabs.SelectedTab.Tag; try { if (Ready) await Script("select", selected); SetStatus(); } catch (Exception error) { Report(error); } }
        /// <summary>Capture les textes et révisions les plus récents depuis la surface Monaco.</summary>
        /// <returns>Tâche terminée après la mise à jour des documents hôtes.</returns>
        private async Task CaptureDocuments()
        {
            var snapshots = json.Deserialize<EditorMessage[]>(await Script("snapshots"));
            if (snapshots == null) return;
            foreach (var item in snapshots)
                if (documents.TryGetValue(item.id, out var doc) && item.version >= versions[item.id])
                { doc.Edit(item.text); versions[item.id] = item.version; }
        }
        /// <summary>Capture les documents, observe les changements natifs et synchronise les brouillons admissibles.</summary>
        /// <param name="synchronize">Autorise l’écriture native des brouillons propres et modifiés.</param>
        /// <returns>Tâche terminée après le traitement de tous les documents ouverts.</returns>
        internal async Task ProcessDocuments(bool synchronize)
        {
            if (busy || !Ready || closing) return;
            busy = true;
            try { await ProcessDocumentsCore(synchronize); }
            catch (Exception error) { Report(error); }
            finally { busy = false; }
        }
        /// <summary>Performs the process documents core operation for ModernEditorWindow.</summary>
        /// <param name="synchronize">Indicates whether synchronize is enabled.</param>
        /// <returns>The result produced by this operation.</returns>
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
        /// <summary>Traite les changements après un court délai de repos puis observe le mode de débogage VBE.</summary>
        /// <param name="sender">Minuterie de la fenêtre.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void TimerTick(object sender, EventArgs e)
        {
            if (busy || closing) return;
            await ProcessDocuments(DateTime.UtcNow - lastEdit > TimeSpan.FromMilliseconds(600));
            try { if (!closing && !IsDisposed) await ObserveDebugMode(); } catch (Exception error) { LoadLog.Write("Monaco debug observation: " + error.Message); }
        }
        /// <summary>Planifie la mise à jour des contrôles de statut après la fin des callbacks WebView2.</summary>
        private void SetStatus()
        {
            if (IsDisposed || Disposing || closing) return;
            // WebView callbacks must unwind before changing WinForms visibility/layout.
            int generation = ++statusGeneration;
            if (IsHandleCreated) BeginInvoke(new Action(() => { if (generation == statusGeneration) UpdateStatus(); })); else UpdateStatus();
        }
                /// <summary>Publishes a result and invalidates older queued synchronization status updates.</summary>
                /// <param name="text">Text containing the text.</param>
        private void SetResultStatus(string text) { statusGeneration++; status.Text = text; }
        /// <summary>Met à jour les boutons de conflit, les titres d’onglets et le statut du document actif.</summary>
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
            status.Text = UiText.Get(lastSaveError ?? synchronizationError ?? (Current == null ? "Open a VBA module to start editing." : Current.Conflict ? "The module changed in VBA. Resolve the conflict first." : Current.Dirty ? "Changes pending synchronization with VBA." : "Synchronized with VBA. Save the macro in its host application."));
            }
            finally { activeStatusLayouts--; }
        }
        /// <summary>Affiche le message d’une erreur si la fenêtre reste active et consigne son type.</summary>
        /// <param name="error">Erreur à présenter et à journaliser.</param>
        private void Report(Exception error) { if (!IsDisposed && !Disposing && !closing) SetResultStatus(UiText.Get(error.Message)); LoadLog.Write("Monaco: " + error.GetType().Name); }
        /// <summary>Performs the close tab requested operation for ModernEditorWindow.</summary>
        /// <param name="sender">The sender used by this operation.</param>
        /// <param name="e">The e used by this operation.</param>
        private void CloseTabRequested(object sender, TabControlEventArgs e)
        {
            if (busy) return;
            tabs.SelectedTab = e.TabPage;
            CloseModuleClick(sender, EventArgs.Empty);
        }
        /// <summary>Capture l’état avant d’afficher la comparaison entre le brouillon et le code natif.</summary>
        /// <param name="sender">Bouton de comparaison.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Applique le brouillon à la version native explicitement comparée par l’utilisateur.</summary>
        /// <param name="sender">Bouton de résolution du conflit.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Ferme l’onglet sélectionné après capture et sauvegarde de son brouillon.</summary>
        /// <param name="sender">Bouton de fermeture du module.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Masque la comparaison et revient à l’édition du brouillon.</summary>
        /// <param name="sender">Bouton d’édition.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void EditClick(object sender, EventArgs e) { try { await Script("hideDiff"); showingDiff = false; SetStatus(); } catch (Exception error) { Report(error); } }
        /// <summary>Recharge le code natif en conservant séparément tout brouillon local modifié.</summary>
        /// <param name="sender">Bouton de rechargement.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Restaure dans Monaco le brouillon précédemment récupéré pour le document actif.</summary>
        /// <param name="sender">Bouton de restauration du brouillon.</param>
        /// <param name="e">Données de l’événement.</param>
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
        /// <summary>Applique à Monaco les couleurs du thème hôte après un changement de thème.</summary>
        private async void ThemeChanged()
        {
            if (IsDisposed || Disposing || closing || !IsHandleCreated) return;
            // Windows preference notifications can arrive on the SystemEvents worker.
            // WebView2, tabs and their palette must be updated on the owning UI thread.
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action(ThemeChanged)); }
                catch (InvalidOperationException) { /* Window closed during dispatch. */ }
                return;
            }
            try { await Theme(); } catch (Exception error) { Report(error); }
        }
        /// <summary>Envoie à Monaco les indicateurs de thème sombre et de contraste élevé.</summary>
        /// <returns>Résultat JSON de l’appel de thème.</returns>
        private Task<string> Theme()
        {
            UiTheme.Apply(this);
            if (Browser != null) Browser.DefaultBackgroundColor = UiTheme.Surface;
            return Script("theme", UiTheme.Dark, UiTheme.HighContrast(),
                ThemeColor(UiTheme.Surface), ThemeColor(UiTheme.Foreground));
        }
        /// <summary>Monaco requires hexadecimal colors, including named and system colors.</summary>
        private static string ThemeColor(System.Drawing.Color color) => "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        /// <summary>Enregistre chaque brouillon modifié avant un arrêt ou une fermeture de la fenêtre.</summary>
        private void PreserveDrafts() { foreach (var doc in documents.Values) try { Drafts.Save(doc); } catch (Exception error) { LoadLog.Write("Editor recovery failed: " + error.GetType().Name); } }
        /// <summary>Capture et préserve les brouillons avant de permettre la fermeture définitive du formulaire.</summary>
        /// <param name="sender">Formulaire en cours de fermeture.</param>
        /// <param name="e">Événement annulable de fermeture.</param>
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
        /// <summary>Arrête les minuteries et workers, détache le thème, libère WebView2 et ferme les CodePane détenus.</summary>
        private void DisposeRuntime()
        {
            timer?.Stop(); PreserveDrafts(); synchronizationWorker?.Dispose(); UiTheme.Changed -= ThemeChanged; Browser?.Dispose();
            foreach (var doc in documents.Values) (doc.Module as EditorVbeModule)?.CloseNativeWindow();
        }
    }
}
