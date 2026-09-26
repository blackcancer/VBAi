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
                dispatcher = new Control();
                var handle = dispatcher.Handle;
                server = new BridgeServer(dispatcher, new VbeSession(vbe), process.Id);
                server.Start();
                LoadLog.Write("Bridge started: CodexVBE." + process.Id);
                try { ShowChat(); }
                catch (Exception uiError) { LoadLog.Write("Assistant window failed: " + uiError); }
                try { menu = new VbeMenu(vbe, ShowChat, ShowSettings); }
                catch (Exception menuError) { LoadLog.Write("VBE menu failed: " + menuError); }
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
            if (chat == null || chat.IsDisposed)
            {
                chat = new ChatWindow(new VbeSession(vbe));
                chat.FormClosed += (sender, args) => chat = null;
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
            ShowChat();
            chat.ShowSettings();
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
            chat?.Dispose();
            chat = null;
            server?.Dispose();
            server = null;

            dispatcher?.Dispose();
            dispatcher = null;
            vbe = null;
        }
    }
}
