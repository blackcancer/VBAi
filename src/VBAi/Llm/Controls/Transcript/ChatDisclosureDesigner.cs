using System.ComponentModel;
using System.Windows.Forms.Design;
namespace VBAi
{
    /// <summary>Makes the disclosure body an editable nested container in the WinForms Designer.</summary>
    public sealed class ChatDisclosureDesigner : ParentControlDesigner
    {
        /// <summary>Exposes the body to drag-and-drop and Designer serialization.</summary>
        /// <param name="component">Disclosure being edited.</param>
        public override void Initialize(IComponent component)
        {
            base.Initialize(component);
            EnableDesignMode(((ChatDisclosureView)component).ContentPanel, "ContentPanel");
        }
    }
}
