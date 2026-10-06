using System;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays recoverable changes made to a VBA form.</summary>
    public sealed partial class ChatFormRecoveryView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatFormRecoveryView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
