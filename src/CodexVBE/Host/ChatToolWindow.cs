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
        private readonly Timer siteResizeTimer;

        public ChatToolWindow()
        {
            Dock = DockStyle.Fill;
            Size = new System.Drawing.Size(520, 760);
            siteResizeTimer = new Timer { Interval = 300 };
            siteResizeTimer.Tick += (sender, args) => FitNativeSite();
        }

        internal void Attach(ChatWindow chat)
        {
            chat.Hide(); chat.TopLevel = false; chat.FormBorderStyle = FormBorderStyle.None;
            chat.Dock = DockStyle.Fill; Controls.Add(chat); chat.Show();
            siteResizeTimer.Start();
            BeginInvoke((Action)FitNativeSite);
        }
        internal void Detach(ChatWindow chat)
        {
            siteResizeTimer.Stop();
            chat.Hide(); Controls.Remove(chat); chat.Dock = DockStyle.None;
            chat.TopLevel = true; chat.FormBorderStyle = FormBorderStyle.Sizable;
        }

        private void FitNativeSite()
        {
            if (!IsHandleCreated || IsDisposed) return;
            var site = GetParent(Handle);
            if (site == IntPtr.Zero) return;
            NativeRect client;
            NativeRect control;
            if (!GetClientRect(site, out client) || !GetWindowRect(Handle, out control)) return;
            var origin = new NativePoint { X = control.Left, Y = control.Top };
            if (!ScreenToClient(site, ref origin)) return;
            var width = Math.Max(1, client.Right - origin.X);
            var height = Math.Max(1, client.Bottom - origin.Y);
            if (Width == width && Height == height) return;
            SetWindowPos(Handle, IntPtr.Zero, origin.X, origin.Y, width, height, 0x0004 | 0x0010);
            Size = new System.Drawing.Size(width, height);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) siteResizeTimer.Dispose();
            base.Dispose(disposing);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)]
        private struct NativePoint { public int X, Y; }
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr handle);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetClientRect(IntPtr handle, out NativeRect rect);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr handle, out NativeRect rect);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ScreenToClient(IntPtr handle, ref NativePoint point);
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr handle, IntPtr after, int x, int y, int width, int height, uint flags);
    }
}
