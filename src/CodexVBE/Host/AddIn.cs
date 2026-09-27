using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CodexVBE
{
    [ComVisible(true)]
    [Guid("8E854243-087F-4D6C-9E0E-8622B0E50883")]
    [ProgId("CodexVBE.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDTExtensibility2))]
    public sealed class AddIn : IDTExtensibility2
    {
        private Control dispatcher;
        private BridgeServer server;
        private ChatWindow chat;
        private VbeMenu menu;
        private object vbe;
        private object addIn;
        private object nativeChatWindow;
        private ChatToolWindow nativeChatControl;
        private bool docked;

        public AddIn()
        {
            var process = Process.GetCurrentProcess();
            LoadLog.Write("Constructed: " + process.ProcessName + " PID=" + process.Id);
        }

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

        private sealed class VbeWindowOwner : IWin32Window
        {
            public VbeWindowOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; private set; }
        }

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

        private IWin32Window VbeOwner()
        {
            return new VbeWindowOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd)));
        }

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

        private void ReportMenuError(Exception ex)
        {
            LoadLog.Write("VBE menu action failed: " + ex);
            MessageBox.Show(ex.Message, "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

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

        public void OnDisconnection(int removeMode, ref object[] custom)
        {
            LoadLog.Write("OnDisconnection: " + removeMode);
            Dispose();
        }

        public void OnAddInsUpdate(ref object[] custom) { }
        public void OnStartupComplete(ref object[] custom) { }
        public void OnBeginShutdown(ref object[] custom) { Dispose(); }

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
