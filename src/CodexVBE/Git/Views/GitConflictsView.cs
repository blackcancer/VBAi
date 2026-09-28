using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des conflits Git.</summary>
    public sealed partial class GitConflictsView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitConflictsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
