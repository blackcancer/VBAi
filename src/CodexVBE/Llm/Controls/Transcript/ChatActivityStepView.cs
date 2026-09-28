using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatActivityStepView layout.</summary>
    public sealed partial class ChatActivityStepView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatActivityStepView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
