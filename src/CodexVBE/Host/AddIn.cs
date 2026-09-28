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
        /// <summary>Contrôle WinForms fournissant un contexte de synchronisation pour le serveur local.</summary>
        private Control dispatcher;
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

        /// <summary>Crée l’instance COM et journalise le processus hôte.</summary>
        public AddIn()
        {
            var process = Process.GetCurrentProcess();
            LoadLog.Write("Constructed: " + process.ProcessName + " PID=" + process.Id);
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
                LoadLog.Write("OnConnection: " + process.ProcessName + " PID=" + process.Id);
                vbe = application;
                UiText.Initialize(vbe);
                addIn = addInInstance;
                LoadLog.Write("AddInInst: " + (addIn == null ? "null" : addIn.GetType().FullName)
                    + ", COM=" + (addIn != null && Marshal.IsComObject(addIn)));
                try { LoadLog.Write("AddInInst ProgId: " + ((dynamic)addIn).ProgId); }
                catch (Exception infoError) { LoadLog.Write("AddInInst ProgId unavailable: " + infoError.Message); }
                dispatcher = new Control();
                var handle = dispatcher.Handle;
                server = new BridgeServer(dispatcher, new VbeSession(vbe), process.Id);
                server.Start();
                LoadLog.Write("Bridge started: CodexVBE." + process.Id);
                try { menu = new VbeMenu(vbe, ShowChat, ShowSettings, ShowGitHub, command => { ShowChat(); chat.PrepareEditorAction(command); }); }
                catch (Exception menuError) { LoadLog.Write("VBE menu failed: " + menuError); }
                try { ShowChat(); ToggleDock(); }
                catch (Exception uiError) { LoadLog.Write("Assistant window failed: " + uiError); }
            }
            catch (Exception ex)
            {
                LoadLog.Write("OnConnection failed: " + ex);
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
                chat = new ChatWindow(new VbeSession(vbe));
                chat.DockRequested += ToggleDock;
                chat.FormClosed += (sender, args) => chat = null;
                if (docked && nativeChatControl != null) nativeChatControl.Attach(chat);
            }
            if (docked && nativeChatWindow != null)
            {
                ((dynamic)nativeChatWindow).Visible = true;
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
                    LoadLog.Write("VBE window owner unavailable: " + ownerError.Message);
                    chat.Show();
                }
            }
            else chat.Activate();
            LoadLog.Write("Assistant window shown.");
        }

        /// <summary>Ouvre les paramètres depuis la fenêtre de conversation ou directement.</summary>
        private void ShowSettings()
        {
            try
            {
                if (chat != null && !chat.IsDisposed) chat.ShowSettings(VbeOwner());
                else using (var dialog = new LlmSettingsWindow(LlmSettings.Load()))
                    dialog.ShowDialog(VbeOwner());
            }
            catch (Exception ex) { ReportMenuError(ex); }
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
                dynamic project = ((dynamic)vbe).ActiveVBProject;
                if (project == null) throw new InvalidOperationException(UiText.Get("Select a saved VBA project to open GitHub."));
                string path = (string)project.FileName;
                if (string.IsNullOrWhiteSpace(path) || !System.IO.Path.IsPathRooted(path))
                    throw new InvalidOperationException(UiText.Get("Save the macro before opening GitHub."));
                var session = new VbeSession(vbe);
                string scope = session.GitScope(path);
                using (var dialog = new GitWindow(session.GitProject(path, scope), scope,
                    (string)project.Name, LlmSettings.Load().GitHubAccount))
                    dialog.ShowDialog(VbeOwner());
            }
            catch (Exception ex) { ReportMenuError(ex); }
        }

        /// <summary>Journalise et affiche une erreur issue d’une commande du menu.</summary>
                /// <param name="ex">Exception levée pendant une action de menu.</param>
private void ReportMenuError(Exception ex)
        {
            LoadLog.Write("VBE menu action failed: " + ex);
            MessageBox.Show(ex.Message, "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        LoadLog.Write("Tool window AddIn from collection: " + ((dynamic)addInForWindow).ProgId);
                    }
                    catch (Exception lookupError) { LoadLog.Write("Tool window AddIn lookup failed: " + lookupError.Message); }
                    nativeChatWindow = ((IVbeWindows)((dynamic)vbe).Windows).CreateToolWindow((IVbeAddIn)addInForWindow, "CodexVBE.ChatToolWindow",
                        "VBAi", "{B5C96ED5-1B16-497C-8441-B3F471F9F92B}", ref document);
                    nativeChatControl = document as ChatToolWindow;
                    if (nativeChatControl == null) throw new InvalidOperationException(UiText.Get("The window's COM control was not created."));
                }
                ((dynamic)nativeChatWindow).Visible = true;
                nativeChatControl.Attach(chat); docked = true;
                try { ((dynamic)vbe).MainWindow.LinkedWindows.Add(nativeChatWindow); }
                catch (Exception positionError) { LoadLog.Write("Native chat main-frame docking unavailable: " + positionError.Message); }
                ((dynamic)nativeChatWindow).SetFocus();
            }
            catch (Exception ex)
            {
                LoadLog.Write("Native chat docking failed: " + ex);
                try { if (!chat.TopLevel) nativeChatControl?.Detach(chat); } catch { }
                docked = false;
                try { if (nativeChatWindow != null) ((dynamic)nativeChatWindow).Close(); } catch { }
                nativeChatWindow = null; nativeChatControl = null;
                ShowChat();
                chat.ReportDockFailure(ex.Message);
            }
        }

        /// <summary>Libère les services et fenêtres quand l’hôte déconnecte l’add-in.</summary>
        /// <param name="removeMode">Mode de suppression transmis par l’hôte COM.</param>
        /// <param name="custom">Données personnalisées transmises par l’hôte, éventuellement modifiées par l’add-in.</param>
public void OnDisconnection(int removeMode, ref object[] custom)
        {
            LoadLog.Write("OnDisconnection: " + removeMode);
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
public void OnBeginShutdown(ref object[] custom) { Dispose(); }

        /// <summary>Détache et ferme les fenêtres, menus, serveur et contrôle de synchronisation.</summary>
        private void Dispose()
        {
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
