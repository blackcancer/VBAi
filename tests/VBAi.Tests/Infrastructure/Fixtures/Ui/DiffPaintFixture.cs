using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
namespace VBAi.Tests.Infrastructure
{
    /// <summary>Construit les arguments de peinture utilisés pour tester les cellules de différence.</summary>
    internal static class DiffPaintFixture
    {
        /// <summary>Crée des arguments de peinture avec la grille comme propriétaire pour les tests.</summary>
        /// <param name="grid">Grille contenant la cellule.</param>
        /// <param name="graphics">Contexte graphique de peinture.</param>
        /// <param name="bounds">Zone de la cellule à peindre.</param>
        /// <param name="row">Index de ligne de la cellule.</param>
        /// <param name="column">Index de colonne de la cellule.</param>
        /// <param name="state">État visuel de la cellule.</param>
        /// <param name="value">Valeur affichée dans la cellule.</param>
        /// <param name="style">Style appliqué à la cellule.</param>
        /// <returns>Arguments prêts à être transmis au gestionnaire de peinture.</returns>
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
