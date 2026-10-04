using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VBAi
{
    // COM control created by VBIDE.Windows.CreateToolWindow. The VBE owns docking and placement.
    /// <summary>Contrôle COM hébergé et positionné par la fenêtre native créée par le VBE.</summary>
    [ComVisible(true)]
    [Guid("0F4D723B-97D8-42E5-9B31-70646B97C8D2")]
    [ProgId("VBAi.ChatToolWindow")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed partial class ChatToolWindow : UserControl
    {
        /// <summary>Lit un rectangle Win32 associé à une fenêtre.</summary>
        /// <param name="handle">Handle de la fenêtre source.</param>
        /// <param name="rect">Rectangle obtenu.</param>
        /// <returns><see langword="true"/> lorsque le rectangle a été lu.</returns>
        internal delegate bool RectReader(IntPtr handle, out NativeRect rect);
        /// <summary>Convertit un point entre les coordonnées écran et client.</summary>
        /// <param name="handle">Handle de la fenêtre dont le repère client est utilisé.</param>
        /// <param name="point">Point à convertir, remplacé par ses coordonnées converties.</param>
        /// <returns><see langword="true"/> lorsque la conversion a réussi.</returns>
        internal delegate bool PointConverter(IntPtr handle, ref NativePoint point);
        /// <summary>Lit le parent natif du contrôle.</summary>
        internal Func<IntPtr, IntPtr> ParentReader = GetParent;
        /// <summary>Lit la zone client de la fenêtre hôte.</summary>
        internal RectReader ClientReader = GetClientRect;
        /// <summary>Lit le rectangle écran du contrôle.</summary>
        internal RectReader WindowReader = GetWindowRect;
        /// <summary>Convertit l’origine du contrôle en coordonnées client de l’hôte.</summary>
        internal PointConverter CoordinateConverter = ScreenToClient;
        /// <summary>Redimensionne et repositionne le contrôle dans la fenêtre native.</summary>
        internal Func<IntPtr, IntPtr, int, int, int, int, uint, bool> PositionWindow = SetWindowPos;

        /// <summary>Crée le contrôle COM et démarre le suivi du site natif.</summary>
        public ChatToolWindow()
        {
            InitializeComponent();
        }

        /// <summary>Allows native focus loss after the host has detached its ActiveX site.</summary>
        /// <param name="e">The focus notification received on the owning UI thread.</param>
        protected override void OnLostFocus(EventArgs e)
        {
            try { base.OnLostFocus(e); }
            catch (InvalidComObjectException)
            {
                // WinForms ActiveXImpl.OnFocus can notify an already detached site
                // during native host shutdown. Do not repeat that notification.
                LoadLog.Write("Native tool-window focus site already detached.");
            }
        }

        /// <summary>Intègre la fenêtre de conversation comme contrôle enfant de ce conteneur.</summary>
        /// <param name="chat">Fenêtre de conversation à attacher ou détacher.</param>
        internal void Attach(Form chat)
        {
            Dock = DockStyle.Fill;
            chat.Hide(); chat.TopLevel = false; chat.FormBorderStyle = FormBorderStyle.None;
            chat.Dock = DockStyle.Fill; Controls.Add(chat); chat.Show();
            siteResizeTimer.Start();
            BeginInvoke((Action)FitNativeSite);
        }
        /// <summary>Retire la fenêtre du conteneur et la restaure comme fenêtre autonome.</summary>
        /// <param name="chat">Fenêtre de conversation à attacher ou détacher.</param>
        internal void Detach(Form chat)
        {
            siteResizeTimer.Stop();
            chat.Hide(); Controls.Remove(chat); chat.Dock = DockStyle.None;
            chat.TopLevel = true; chat.FormBorderStyle = FormBorderStyle.Sizable;
        }

        /// <summary>Actualise la taille du contrôle lors du tick du minuteur.</summary>
        /// <param name="sender">Minuteur déclencheur.</param>
        /// <param name="e">Événement du minuteur.</param>
        private void SiteResizeTimer_Tick(object sender, EventArgs e) { FitNativeSite(); }

                /// <summary>Lit la taille réelle de la zone cliente du site natif, qui peut différer du cadre VBIDE.</summary>
                /// <param name="size">Reçoit la taille lue, ou une taille vide si le site ne peut pas être interrogé.</param>
                /// <returns><see langword="true"/> si la zone cliente native a été lue.</returns>
        internal bool TryGetNativeSiteSize(out System.Drawing.Size size)
        {
            size = System.Drawing.Size.Empty;
            if (!IsHandleCreated || IsDisposed) return false;
            var site = ParentReader(Handle);
            NativeRect client;
            if (site == IntPtr.Zero || !ClientReader(site, out client)) return false;
            size = new System.Drawing.Size(Math.Max(0, client.Right - client.Left), Math.Max(0, client.Bottom - client.Top));
            return true;
        }
        /// <summary>Ajuste la taille du contrôle à la zone client du site VBE.</summary>
        private void FitNativeSite()
        {
            if (!IsHandleCreated || IsDisposed) return;
            var site = ParentReader(Handle);
            if (site == IntPtr.Zero) return;
            NativeRect client;
            NativeRect control;
            if (!ClientReader(site, out client) || !WindowReader(Handle, out control)) return;
            var origin = new NativePoint { X = control.Left, Y = control.Top };
            if (!CoordinateConverter(site, ref origin)) return;
            var width = Math.Max(1, client.Right - origin.X);
            var height = Math.Max(1, client.Bottom - origin.Y);
            if (Width == width && Height == height) return;
            PositionWindow(Handle, IntPtr.Zero, origin.X, origin.Y, width, height, 0x0004 | 0x0010);
            Size = new System.Drawing.Size(width, height);
        }

        /// <summary>Libère le minuteur de synchronisation de taille.</summary>
        /// <param name="disposing">Indique si les ressources gérées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Structure Win32 de rectangle en coordonnées écran.</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct NativeRect
        {
            /// <summary>Bord gauche du rectangle.</summary>
            public int Left;
            /// <summary>Bord supérieur du rectangle.</summary>
            public int Top;
            /// <summary>Bord droit du rectangle.</summary>
            public int Right;
            /// <summary>Bord inférieur du rectangle.</summary>
            public int Bottom;
        }
        /// <summary>Point Win32 utilisé pour convertir les coordonnées écran en coordonnées client.</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct NativePoint
        {
            /// <summary>Coordonnée horizontale du point.</summary>
            public int X;
            /// <summary>Coordonnée verticale du point.</summary>
            public int Y;
        }
        /// <summary>Obtient le HWND parent du contrôle dans le site VBE.</summary>
        /// <param name="handle">Handle Win32 du contrôle ou site à interroger.</param>
        /// <returns>Handle du parent, ou zéro sans parent natif.</returns>
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr handle);
        /// <summary>Obtient le rectangle client d’une fenêtre.</summary>
        /// <param name="handle">Handle Win32 du contrôle ou site à interroger.</param>
        /// <param name="rect">Rectangle de sortie fourni par Windows.</param>
        /// <returns>true si Windows a fourni le rectangle client.</returns>
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetClientRect(IntPtr handle, out NativeRect rect);
        /// <summary>Obtient le rectangle écran d’une fenêtre.</summary>
        /// <param name="handle">Handle Win32 du contrôle ou site à interroger.</param>
        /// <param name="rect">Rectangle de sortie fourni par Windows.</param>
        /// <returns>true si Windows a fourni le rectangle écran.</returns>
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr handle, out NativeRect rect);
        /// <summary>Convertit un point écran vers les coordonnées client d’une fenêtre.</summary>
        /// <param name="handle">Handle Win32 de la fenêtre de destination.</param>
        /// <param name="point">Point à convertir, mis à jour par Windows.</param>
        /// <returns>true si le point a été converti.</returns>
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ScreenToClient(IntPtr handle, ref NativePoint point);
        /// <summary>Positionne et redimensionne une fenêtre enfant.</summary>
        /// <param name="handle">Handle Win32 du contrôle ou site à interroger.</param>
        /// <param name="after">Fenêtre placée avant le contrôle, ou null.</param>
        /// <param name="x">Coordonnée horizontale dans l’espace client parent.</param>
        /// <param name="y">Coordonnée verticale dans l’espace client parent.</param>
        /// <param name="width">Nouvelle largeur de la fenêtre.</param>
        /// <param name="height">Nouvelle hauteur de la fenêtre.</param>
        /// <param name="flags">Options de positionnement Win32.</param>
        /// <returns>true si le déplacement ou redimensionnement a réussi.</returns>
        [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr handle, IntPtr after, int x, int y, int width, int height, uint flags);
    }
}
