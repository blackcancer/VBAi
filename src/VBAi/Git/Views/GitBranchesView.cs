using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Vue des branches Git.</summary>
    public sealed partial class GitBranchesView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitBranchesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
