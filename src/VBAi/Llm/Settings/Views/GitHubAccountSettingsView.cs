using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue de configuration du compte GitHub.</summary>
    public sealed partial class GitHubAccountSettingsView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubAccountSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
