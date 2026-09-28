using System;
using System.Windows.Forms;
namespace CodexVBE
{
        /// <summary>Groups related chat activity entries under one expandable transcript section.</summary>
public sealed partial class ChatActivityGroupView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatActivityGroupView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}
