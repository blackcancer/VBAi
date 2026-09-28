using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue de connexion du document à un dépôt Git.</summary>
    public sealed partial class GitConnectionView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitConnectionView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
