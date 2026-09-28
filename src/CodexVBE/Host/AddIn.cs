using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Point d’entrée COM qui démarre le serveur de commandes et les fenêtres VBAi.</summary>
    [ComVisible(true)]
    [Guid("8E854243-087F-4D6C-9E0E-8622B0E50883")]
    [ProgId("CodexVBE.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDTExtensibility2))]
        public sealed class AddIn : IDTExtensibility2
    {
        /// <summary>Écrit une entrée dans le journal du complément.</summary>
        internal static Action<string> WriteLog = LoadLog.Write;
        /// <summary>Démarre les mises à jour sur un déploiement géré par l’installeur.</summary>
        internal static Action StartUpdateCheck = UpdateCoordinator.Start;
        /// <summary>Arrête les vérifications en cours quand le complément est déconnecté.</summary>
        internal static Action StopUpdateCheck = UpdateCoordinator.Stop;
        /// <summary>Crée le collecteur d’erreurs lié à la session et à son stockage local.</summary>
        internal static Func<Action<CrashReport>, CrashReporter> CreateCrashReporter = show => new CrashReporter(show);
        /// <summary>Démarre le pont de commandes local pour la session active.</summary>
        internal static Action<BridgeServer> StartBridge = (Action<BridgeServer>)Delegate.CreateDelegate(typeof(Action<BridgeServer>), typeof(BridgeServer).GetMethod("Start"));
        /// <summary>Crée la fenêtre de discussion liée à une session VBE.</summary>
        internal static Func<VbeSession, ChatWindow> CreateChat = CreateChatNative;
        /// <summary>Crée le gestionnaire de menus du VBE.</summary>
        internal static Func<object, Action, Action, Action, Action<string>, VbeMenu> CreateMenu = CreateMenuNative;
        /// <summary>Charge les paramètres des fournisseurs LLM.</summary>
        internal static Func<LlmSettings> ReadSettings = LlmSettings.Load;
        /// <summary>Affiche une fenêtre modale avec son propriétaire Win32.</summary>
        internal static Func<Form, IWin32Window, DialogResult> ShowModal = (Func<Form, IWin32Window, DialogResult>)Delegate.CreateDelegate(typeof(Func<Form, IWin32Window, DialogResult>), typeof(Form).GetMethod("ShowDialog", new[] { typeof(IWin32Window) }));
        /// <summary>Affiche une notification WinForms et renvoie le choix de l’utilisateur.</summary>
        internal static Func<string, string, MessageBoxButtons, MessageBoxIcon, DialogResult> ShowNotice = MessageBox.Show;
        /// <summary>Crée directement une fenêtre de discussion.</summary>
        /// <param name="session">Session VBE associée à la fenêtre.</param>
        /// <returns>Nouvelle fenêtre de discussion.</returns>
        private static ChatWindow CreateChatNative(VbeSession session) { return new ChatWindow(session); }
        /// <summary>Crée directement le gestionnaire des commandes de menu VBE.</summary>
        /// <param name="host">Instance hôte dont les barres de commandes seront utilisées.</param>
        /// <param name="chat">Action d’ouverture de la discussion.</param>
        /// <param name="settings">Action d’ouverture des paramètres.</param>
        /// <param name="github">Action d’ouverture de GitHub.</param>
        /// <param name="editor">Action de commande associée au texte fourni.</param>
        /// <returns>Gestionnaire des menus installé sur l’hôte.</returns>
        private static VbeMenu CreateMenuNative(object host, Action chat, Action settings, Action github, Action<string> editor) { return new VbeMenu(host, chat, settings, github, editor, () => AboutWindow.ShowForVbe(host), () => CrashReportWindow.ShowForVbe(host), () => UpdateWindow.ShowForVbe(host)); }
        /// <summary>Contrôle WinForms fournissant un contexte de synchronisation pour le serveur local.</summary>
        private Control dispatcher;
        private CrashReporter crashReporter;
        /// <summary>Serveur de commandes local rattaché à l’instance du VBE.</summary>
        private BridgeServer server;
        /// <summary>Fenêtre de conversation actuellement ouverte.</summary>
        private ChatWindow chat;
        /// <summary>Commandes ajoutées aux barres du VBE.</summary>
        private VbeMenu menu;
        /// <summary>Instance VBE fournie par l’hôte COM.</summary>
        private object vbe;
        /// <summary>Instance COM de l’add-in enregistrée dans l’hôte.</summary>
        private object addIn;
        /// <summary>Fenêtre native du VBE qui héberge le contrôle de conversation.</summary>
        private object nativeChatWindow;
        /// <summary>Contrôle utilisateur contenu dans la fenêtre native du VBE.</summary>
        private ChatToolWindow nativeChatControl;
        /// <summary>Indique si la fenêtre de conversation est attachée au cadre VBE.</summary>
        private bool docked;
        internal static Func<ModernEditorWindow> CreateModernEditor = CreateModernEditorNative;
        private static ModernEditorWindow CreateModernEditorNative() => new ModernEditorWindow();
        private ModernEditorWindow modernEditor;
        private EditorWorkspaceHost editorWorkspace;
        private EditorProjectNavigation editorNavigation;

        /// <summary>Crée l’instance COM et journalise le processus hôte.</summary>
        public AddIn()
        {
            var process = Process.GetCurrentProcess();
            WriteLog("Constructed: " + process.ProcessName + " PID=" + process.Id);
        }

        /// <summary>Initialise la session, démarre le pont local et ajoute les commandes de menu.</summary>
        /// <param name="application">Objet dont les barres de commande sont recherchées.</param>
        /// <param name="connectMode">Mode de connexion communiqué par l’hôte.</param>
        /// <param name="addInInstance">Instance COM de l’add-in hôte.</param>
        /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnConnection(object application, int connectMode, object addInInstance, ref object[] custom)
        {
            try
            {
                var process = Process.GetCurrentProcess();
                WriteLog("OnConnection: " + process.ProcessName + " PID=" + process.Id);
                vbe = application;
                UiText.Initialize(vbe);
                bool nativeDark = false;
                try { nativeDark = ReadSettings().NativeVbeDarkTheme; }
                catch (Exception settingsError) { WriteLog("Native VBE theme setting unavailable: " + settingsError.Message); }
                try
                {
                    var editor = new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd));
                    VbeNativeTheme.Initialize(editor, nativeDark, vbe);
                    if (nativeDark || VbeNativeTheme.ExperimentEnabled()) WriteLog("Native VBE dark mode enabled.");
                }
                catch (Exception themeError) { WriteLog("Native VBE dark mode unavailable: " + themeError.ToString()); }
                addIn = addInInstance;
                WriteLog("AddInInst: " + (addIn == null ? "null" : addIn.GetType().FullName)
                    + ", COM=" + (addIn != null && Marshal.IsComObject(addIn)));
                try { WriteLog("AddInInst ProgId: " + ((dynamic)addIn).ProgId); }
                catch (Exception infoError) { WriteLog("AddInInst ProgId unavailable: " + infoError.Message); }
                dispatcher = new Control();
                var handle = dispatcher.Handle;
                StartUpdateCheck();
                crashReporter = CreateCrashReporter(report => CrashReportWindow.ShowReportForVbe(vbe, report));
                server = new BridgeServer(dispatcher, CreateEditorSession(), process.Id);
                StartBridge(server);
                WriteLog("Bridge started: CodexVBE." + process.Id);
                try { menu = CreateMenu(vbe, ShowChat, ShowSettings, ShowGitHub, PrepareEditorAction); }
                catch (Exception menuError) { WriteLog("VBE menu failed: " + menuError.ToString()); crashReporter.ReportUnexpected(menuError); }
                editorNavigation = new EditorProjectNavigation(vbe, dispatcher, OpenModernModule);
                try { ShowChat(); ToggleDock(); }
                catch (Exception uiError) { WriteLog("Assistant window failed: " + uiError.ToString()); crashReporter.ReportUnexpected(uiError); }
                try
                {
                    GetModernEditor(true);
                    var module = ActiveEditorModule(false);
                    if (module != null) OpenModernModule(module);
                }
                catch (Exception editorError) { WriteLog("Modern editor startup failed: " + editorError.ToString()); }
                crashReporter.RecoverPending();
            }
            catch (Exception ex)
            {
                WriteLog("OnConnection failed: " + ex.ToString());
                crashReporter?.ReportUnexpected(ex);
                Dispose();
                throw;
            }
        }

        /// <summary>Wrapper de poignée HWND utilisé comme propriétaire WinForms.</summary>
        private sealed class VbeWindowOwner : IWin32Window
        {
        /// <summary>Wrapper de poignée HWND utilisé comme propriétaire WinForms.</summary>
        /// <param name="handle">Poignée HWND de la fenêtre propriétaire.</param>
            public VbeWindowOwner(IntPtr handle) { Handle = handle; }
        /// <summary>Poignée HWND du propriétaire VBE.</summary>
        /// <value>Poignée HWND du propriétaire VBE.</value>
            public IntPtr Handle { get; private set; }
        }

        /// <summary>Affiche ou réactive la fenêtre de conversation, intégrée si elle est attachée.</summary>
        private void ShowChat()
        {
            if (nativeChatControl != null && nativeChatControl.IsDisposed)
            { nativeChatControl = null; nativeChatWindow = null; docked = false; }
            if (chat == null || chat.IsDisposed)
            {
                chat = CreateChat(CreateEditorSession());
                chat.DockRequested += ToggleDock;
                chat.FormClosed += (sender, args) => chat = null;
                if (docked && nativeChatControl != null) nativeChatControl.Attach(chat);
            }
            if (docked && nativeChatWindow != null)
            {
                ((dynamic)nativeChatWindow).Visible = true;
                EnsureUsableChatPlacement();
                ((dynamic)nativeChatWindow).SetFocus();
                return;
            }
            if (!chat.Visible)
            {
                try
                {
                    var owner = new VbeWindowOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd)));
                    chat.Show(owner);
                }
                catch (Exception ownerError)
                {
                    WriteLog("VBE window owner unavailable: " + ownerError.Message);
                    chat.Show();
                }
            }
            else chat.Activate();
            WriteLog("Assistant window shown.");
        }

        /// <summary>Ouvre les paramètres depuis la fenêtre de conversation ou directement.</summary>
        private void ShowSettings()
        {
            try
            {
                if (chat != null && !chat.IsDisposed) chat.ShowSettings(VbeOwner());
                else using (var dialog = new LlmSettingsWindow(ReadSettings()))
                    ShowModal(dialog, VbeOwner());
            }
            catch (Exception ex) { ReportMenuError(ex); }
        }

        private async void PrepareEditorAction(string command)
        {
            try
            {
                if (command == "/editor") { ShowModernEditor(); return; }
                if (modernEditor != null && !modernEditor.IsDisposed && modernEditor.Visible && modernEditor.Ready && modernEditor.Current != null)
                { await modernEditor.Script("command", "vbai." + command.TrimStart('/')); return; }
                ShowChat(); chat.PrepareEditorAction(command);
            }
            catch (Exception error) { ReportMenuError(error); }
        }

        private IEditorModule ActiveEditorModule(bool followOnly)
        {
            dynamic host = vbe;
            if (followOnly && (host.ActiveWindow == null || (int)host.ActiveWindow.Type != 0)) return null;
            dynamic pane = host.ActiveCodePane;
            if (pane == null) return null;
            object component = pane.CodeModule.Parent;
            object project = ((dynamic)component).Collection.Parent;
            if (followOnly && (int)((dynamic)project).Mode != 2) return null;
            return new EditorVbeModule(vbe, project, component);
        }
        private VbeSession CreateEditorSession() => new VbeSession(vbe) { ModernEditor = GetModernEditor };
        private async void OpenModernModule(IEditorModule module)
        {
            try { await GetModernEditor(true).OpenModule(module); }
            catch (Exception error) { ReportMenuError(error); }
        }
        private ModernEditorWindow GetModernEditor(bool show)
        {
            if (!show) return modernEditor != null && !modernEditor.IsDisposed ? modernEditor : null;
            if (modernEditor == null || modernEditor.IsDisposed)
            {
                editorWorkspace?.Dispose();
                modernEditor = CreateModernEditor();
                modernEditor.AssistantAction += (command, attachment) => { ShowChat(); chat.PrepareMonacoAction(command, attachment); };
                try { editorWorkspace = new EditorWorkspaceHost(vbe, modernEditor); }
                catch { modernEditor.Dispose(); modernEditor = null; editorWorkspace = null; throw; }
            }
            editorWorkspace.Show();
            return modernEditor;
        }
        private async void ShowModernEditor()
        {
            try
            {
                var module = ActiveEditorModule(false);
                var editor = GetModernEditor(true);
                if (module != null) await editor.OpenModule(module);
            }
            catch (Exception error) { ReportMenuError(error); }
        }
        /// <summary>Construit le propriétaire WinForms à partir de la fenêtre principale du VBE.</summary>
        /// <returns>Fenêtre propriétaire WinForms ancrée sur la fenêtre principale du VBE.</returns>
        private IWin32Window VbeOwner()
        {
            return new VbeWindowOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd)));
        }

        /// <summary>Ouvre l’interface GitHub pour le projet actif enregistré.</summary>
        private void ShowGitHub()
        {
            try
            {
                dynamic project = null;
                try { project = ((dynamic)vbe).ActiveVBProject; }
                catch (System.IO.DirectoryNotFoundException) { }
                catch (COMException) { }
                if (project == null) throw new InvalidOperationException(UiText.Get("Select a saved VBA project to open GitHub."));
                string path = null;
                try { path = (string)project.FileName; }
                catch (System.IO.DirectoryNotFoundException) { }
                catch (COMException) { }
                if (string.IsNullOrWhiteSpace(path) || !System.IO.Path.IsPathRooted(path) ||
                    !System.IO.File.Exists(path))
                    throw new InvalidOperationException(UiText.Get("Save the macro before opening GitHub."));
                var session = CreateEditorSession();
                string scope = System.IO.Path.GetFullPath(path);
                using (var dialog = new GitWindow(session.GitProject(path, scope), scope,
                    (string)project.Name, ReadSettings().GitHubAccount))
                    ShowModal(dialog, VbeOwner());
            }
            catch (Exception ex) { ReportMenuError(ex); }
        }

        /// <summary>Journalise et affiche une erreur issue d’une commande du menu.</summary>
                /// <param name="ex">Exception levée pendant une action de menu.</param>
private void ReportMenuError(Exception ex)
        {
            WriteLog("VBE menu action failed: " + ex.ToString());
            crashReporter?.ReportUnexpected(ex);
            ShowNotice(ex.Message, "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>Attache ou détache la fenêtre de conversation au cadre principal du VBE.</summary>
        private void ToggleDock()
        {
            try
            {
                if (docked)
                {
                    nativeChatControl.Detach(chat); docked = false;
                    ((dynamic)nativeChatWindow).Visible = false; ShowChat();
                    return;
                }
                if (nativeChatWindow == null)
                {
                    object document = null;
                    object addInForWindow = addIn;
                    try
                    {
                        addInForWindow = ((dynamic)vbe).AddIns.Item("CodexVBE.AddIn");
                        WriteLog("Tool window AddIn from collection: " + ((dynamic)addInForWindow).ProgId);
                    }
                    catch (Exception lookupError) { WriteLog("Tool window AddIn lookup failed: " + lookupError.Message); }
                    nativeChatWindow = ((IVbeWindows)((dynamic)vbe).Windows).CreateToolWindow((IVbeAddIn)addInForWindow, "CodexVBE.ChatToolWindow",
                        "VBAi", "{B5C96ED5-1B16-497C-8441-B3F471F9F92B}", ref document);
                    nativeChatControl = document as ChatToolWindow;
                    if (nativeChatControl == null) throw new InvalidOperationException(UiText.Get("The window's COM control was not created."));
                }
                ((dynamic)nativeChatWindow).Visible = true;
                nativeChatControl.Attach(chat); docked = true;
                // CreateToolWindow restores the saved layout. Forcing MainWindow linkage can
                // collapse a new pane to a six-pixel strip at the bottom of the VBE.
                EnsureUsableChatPlacement();
                ((dynamic)nativeChatWindow).SetFocus();
            }
            catch (Exception ex)
            {
                WriteLog("Native chat docking failed: " + ex.ToString());
                try { if (!chat.TopLevel) nativeChatControl?.Detach(chat); } catch { }
                docked = false;
                try { if (nativeChatWindow != null) ((dynamic)nativeChatWindow).Close(); } catch { }
                nativeChatWindow = null; nativeChatControl = null;
                ShowChat();
                chat.ReportDockFailure(ex.Message);
            }
        }

        private void EnsureUsableChatPlacement()
        {
            System.Drawing.Size siteSize;
            if (nativeChatControl == null || !nativeChatControl.TryGetNativeSiteSize(out siteSize) ||
                (siteSize.Width >= chat.MinimumSize.Width && siteSize.Height >= chat.MinimumSize.Height)) return;

            dynamic window = nativeChatWindow;
            dynamic frame = window.LinkedWindowFrame;
            // Bounds on a docked pane describe its containing frame, not the actual site.
            // Detach only the unusable pane before assigning a readable floating layout.
            if (frame != null) frame.LinkedWindows.Remove(window);
            var area = Screen.FromHandle(VbeOwner().Handle).WorkingArea;
            int width = Math.Min(600, area.Width);
            int height = Math.Min(820, area.Height);
            window.Width = width;
            window.Height = height;
            window.Left = area.Right - width;
            window.Top = area.Top + Math.Max(0, (area.Height - height) / 2);
            WriteLog("Recovered unusable chat pane (" + siteSize.Width + "x" + siteSize.Height +
                ") as a native dockable window (" + width + "x" + height + ").");
        }

        /// <summary>Libère les services et fenêtres quand l’hôte déconnecte l’add-in.</summary>
        /// <param name="removeMode">Mode de suppression transmis par l’hôte COM.</param>
        /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnDisconnection(int removeMode, ref object[] custom)
        {
            WriteLog("OnDisconnection: " + removeMode);
            CleanupTemporaryToolbarCommands();
            Dispose();
        }

        /// <summary>Point d’extension COM appelé après la mise à jour de la collection d’add-ins.</summary>
                /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnAddInsUpdate(ref object[] custom) { }
        /// <summary>Point d’extension COM appelé à la fin du démarrage de l’hôte.</summary>
                /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnStartupComplete(ref object[] custom) { }
        /// <summary>Libère les services lorsque l’hôte commence son arrêt.</summary>
                /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnBeginShutdown(ref object[] custom) { CleanupTemporaryToolbarCommands(); Dispose(); }

        /// <summary>Nettoie les boutons de session avant que la référence VBE soit libérée.</summary>
        private void CleanupTemporaryToolbarCommands()
        {
            try { if (vbe != null) new VbeEditorWindows(vbe).RemoveTemporaryToolbarCommands(); }
            catch (Exception error) { WriteLog("Temporary toolbar cleanup failed: " + error.Message); }
        }

        /// <summary>Détache et ferme les fenêtres, menus, serveur et contrôle de synchronisation.</summary>
        private void Dispose()
        {
            try { if (!VbeNativeTheme.Disconnect()) WriteLog("Native VBE theme cleanup deferred: renderer still active."); }
            catch (Exception error) { WriteLog("Native VBE theme cleanup failed: " + error.ToString()); }
            editorNavigation?.Dispose(); editorNavigation = null;
            editorWorkspace?.Dispose(); editorWorkspace = null;
            modernEditor?.Dispose(); modernEditor = null;
            StopUpdateCheck();
            crashReporter?.Dispose();
            crashReporter = null;
            menu?.Dispose();
            menu = null;
            if (nativeChatControl != null && chat != null && !chat.IsDisposed && docked) nativeChatControl.Detach(chat);
            docked = false;
            chat?.Dispose();
            chat = null;
            try { if (nativeChatWindow != null) ((dynamic)nativeChatWindow).Close(); } catch { }
            nativeChatWindow = null; nativeChatControl = null; addIn = null;
            server?.Dispose();
            server = null;

            dispatcher?.Dispose();
            dispatcher = null;
            vbe = null;
        }
    }
}
