using System.Windows;
using System.Windows.Forms.Integration;

namespace CodexVBE
{
    /// <summary>Héberge dans le transcript la vue de diff WinForms éditable dans le Designer.</summary>
    internal sealed class ChatDiffView : WindowsFormsHost
    {
                /// <summary>Les lignes varient ; la grille et les commandes proviennent du Designer de CodeDiffView.</summary>
        /// <param name="before">Original code text.</param><param name="after">Updated code text.</param>
        internal ChatDiffView(string before, string after)
        {
            Height = 300;
            FlowDirection = FlowDirection.LeftToRight;
            var view = new CodeDiffView { UnifiedDiff = true };
            Child = view;
            UiTheme.Apply(view);
            view.ShowDiff(before, after);
        }
    }
}
