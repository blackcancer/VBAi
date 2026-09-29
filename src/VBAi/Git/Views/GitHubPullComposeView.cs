using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Vue de préparation d’une pull request.</summary>
    public sealed partial class GitHubPullComposeView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullComposeView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
