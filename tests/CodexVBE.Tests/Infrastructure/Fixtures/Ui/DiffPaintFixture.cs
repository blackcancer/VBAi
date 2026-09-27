using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
namespace CodexVBE.Tests.Infrastructure
{
    internal static class DiffPaintFixture
    {
        internal static DataGridViewCellPaintingEventArgs Args(DataGridView grid,Graphics graphics,Rectangle bounds,int row,int column,DataGridViewElementStates state,object value,DataGridViewCellStyle style)
        {
            var args=new DataGridViewCellPaintingEventArgs(grid,graphics,new Rectangle(0,0,500,100),bounds,row,column,state,value,value,null,style,new DataGridViewAdvancedBorderStyle(),DataGridViewPaintParts.All);
            // net48's public constructor omits the owner used by PaintBackground;
            // actual DataGridView events initialize it through their internal constructor.
            typeof(DataGridViewCellPaintingEventArgs).GetField("dataGridView",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(args,grid);
            return args;
        }
    }
}
