using System;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays a proposed code change with its diff and available undo actions.</summary>
    public sealed partial class ChatChangeCardView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatChangeCardView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
