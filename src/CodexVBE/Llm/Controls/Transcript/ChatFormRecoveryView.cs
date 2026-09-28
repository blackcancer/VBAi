using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatFormRecoveryView layout.</summary>
    public sealed partial class ChatFormRecoveryView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatFormRecoveryView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
