using System.Drawing;
using System.Windows.Forms;
namespace CodexVBE.Tests.Unit
{
    internal sealed class PreferredHeightView : UserControl
    {
        internal int HeightWanted;
        public override Size GetPreferredSize(Size proposedSize) { return new Size(proposedSize.Width,HeightWanted); }
    }
}