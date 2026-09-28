using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatActivityGroupView layout.</summary>
    public sealed partial class ChatActivityGroupView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatActivityGroupView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
