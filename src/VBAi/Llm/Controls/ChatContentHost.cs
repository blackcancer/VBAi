using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;

namespace VBAi
{

    /// <summary>Surface interop dont la position et l’apparence utilisent le concepteur WinForms standard.</summary>
    /// <remarks>Le contenu riche est fourni par le moteur de chat ; la surface fixe reste éditable sans le concepteur WPF externe.</remarks>
    [ToolboxItem(false)]
    [Designer("System.Windows.Forms.Design.ControlDesigner, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
    public sealed class ChatContentHost : ElementHost
    {

        /// <summary>Crée l’hôte standard sans connexion à l’IDE ou au fournisseur.</summary>
        public ChatContentHost() { }

        /// <summary>Allows a child focus notification after the native ActiveX site has detached.</summary>
        /// <param name="e">The notification received on the owning UI thread.</param>
        protected override void OnGotFocus(EventArgs e)
        {
            try { base.OnGotFocus(e); }
            catch (InvalidComObjectException)
            {
                // ElementHost forwards child focus to the containing ActiveX site.
                // Shutdown may detach that site first; do not repeat the notification.
                LoadLog.Write("Native chat content focus site already detached.");
            }
        }
    }
}
