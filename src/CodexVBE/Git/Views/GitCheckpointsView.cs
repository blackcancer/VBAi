using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des points de contrôle Git.</summary>
    public sealed partial class GitCheckpointsView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitCheckpointsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
