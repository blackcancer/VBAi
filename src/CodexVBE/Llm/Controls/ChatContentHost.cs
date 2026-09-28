using System.ComponentModel;
using System.Windows.Forms.Integration;

namespace CodexVBE
{
    /// <summary>Surface interop dont la position et l’apparence utilisent le concepteur WinForms standard.</summary>
    /// <remarks>Le contenu riche est fourni par le moteur de chat ; la surface fixe reste éditable sans le concepteur WPF externe.</remarks>
    [ToolboxItem(false)]
    [Designer("System.Windows.Forms.Design.ControlDesigner, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
    public sealed class ChatContentHost : ElementHost
    {
        /// <summary>Crée l’hôte standard sans connexion à l’IDE ou au fournisseur.</summary>
        public ChatContentHost() { }
    }
}
