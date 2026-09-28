using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class AppearanceSettingsView : UserControl
    {
        public AppearanceSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
