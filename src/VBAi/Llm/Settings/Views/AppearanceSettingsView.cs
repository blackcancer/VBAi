using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue des préférences d’apparence.</summary>
    public sealed partial class AppearanceSettingsView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public AppearanceSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
