using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class ProviderSettingsView : UserControl
    {
        public ProviderSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
