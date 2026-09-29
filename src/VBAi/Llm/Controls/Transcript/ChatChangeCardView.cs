using System;
using System.Windows.Forms;
namespace VBAi
{
    /// <summary>Designer-editable ChatChangeCardView layout.</summary>
    public sealed partial class ChatChangeCardView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatChangeCardView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
