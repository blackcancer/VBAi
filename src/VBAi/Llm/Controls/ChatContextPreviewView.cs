using System.Windows.Forms;
namespace VBAi
{
    /// <summary>Modèle Designer d’un segment de contexte ; seules ses données varient.</summary>
    public sealed partial class ChatContextPreviewView : UserControl
    {
        /// <summary>Crée le cadre et la zone de lecture depuis le Designer.</summary>
        public ChatContextPreviewView() { InitializeComponent(); }
                /// <summary>Remplace le titre et le contenu sans reconstruire les contrôles.</summary>
                /// <param name="title">Titre affiché au-dessus du contenu.</param>
                /// <param name="text">Texte présenté dans la zone de lecture.</param>
        public void ShowContent(string title, string text) { section.Text = title ?? ""; content.Text = text ?? ""; }
    }
}
