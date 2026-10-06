using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue de configuration du fournisseur LLM.</summary>
    public sealed partial class ProviderSettingsView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public ProviderSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
