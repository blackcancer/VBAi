using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CodexVBE
{
    // COM control created by VBIDE.Windows.CreateToolWindow. The VBE owns docking and placement.
    [ComVisible(true)]
    [Guid("0F4D723B-97D8-42E5-9B31-70646B97C8D2")]
    [ProgId("CodexVBE.ChatToolWindow")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ChatToolWindow : UserControl
    {
        public ChatToolWindow() { Dock = DockStyle.Fill; }
        internal void Attach(ChatWindow chat)
        {
            chat.Hide(); chat.TopLevel = false; chat.FormBorderStyle = FormBorderStyle.None;
            chat.Dock = DockStyle.Fill; Controls.Add(chat); chat.Show();
        }
        internal void Detach(ChatWindow chat)
        {
            chat.Hide(); Controls.Remove(chat); chat.Dock = DockStyle.None;
            chat.TopLevel = true; chat.FormBorderStyle = FormBorderStyle.Sizable;
        }
    }
}
