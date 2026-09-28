using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Modèle Designer d’une référence ou pièce jointe avec ouverture et retrait.</summary>
    public sealed partial class ChatContextChipView : UserControl
    {
        private bool canOpen;
        /// <summary>Demande l’ouverture de la référence sélectionnée.</summary>
        public event EventHandler OpenRequested;
        /// <summary>Demande le retrait de l’élément du contexte.</summary>
        public event EventHandler RemoveRequested;
        /// <summary>Crée le modèle et ses commandes Designer.</summary>
        public ChatContextChipView() { InitializeComponent(); }
        /// <summary>Configure les données et les actions disponibles pour un élément du contexte.</summary>
        public void ShowItem(string text, bool canOpen, string openTip, string removeTip)
        {
            this.canOpen = canOpen;
            open.Text = text ?? ""; remove.Visible = canOpen;
            toolTips.SetToolTip(open, canOpen ? openTip : removeTip);
            toolTips.SetToolTip(remove, removeTip);
        }
        private void Open_Click(object sender, EventArgs e)
        {
            if (canOpen) OpenRequested?.Invoke(this, e);
            else RemoveRequested?.Invoke(this, e);
        }
        private void Remove_Click(object sender, EventArgs e) { RemoveRequested?.Invoke(this, e); }
    }
}
