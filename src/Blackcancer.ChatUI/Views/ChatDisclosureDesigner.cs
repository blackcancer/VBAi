// SPDX-License-Identifier: MPL-2.0
using System.ComponentModel;
using System.Windows.Forms.Design;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Allows the Designer to edit the fixed disclosure's nested content container.</summary>
    public sealed class ChatDisclosureDesigner : ParentControlDesigner
    {
        public override void Initialize(IComponent component)
        { base.Initialize(component); EnableDesignMode(((ChatDisclosureView)component).ContentPanel, "ContentPanel"); }
    }
}
