using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des changements du document VBA.</summary>
    public sealed partial class GitChangesView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitChangesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
