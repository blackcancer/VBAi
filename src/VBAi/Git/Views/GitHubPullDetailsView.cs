using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue des détails d’une pull request.</summary>
    public sealed partial class GitHubPullDetailsView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullDetailsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
