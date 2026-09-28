using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Modèle Designer d’une référence ou pièce jointe avec ouverture et retrait.</summary>
    public sealed partial class ChatContextChipView : UserControl
    {
        /// <summary>Indique si la vignette représente un élément ouvrable.</summary>
        private bool canOpen;
        /// <summary>Demande l’ouverture de la référence sélectionnée.</summary>
        public event EventHandler OpenRequested;
        /// <summary>Demande le retrait de l’élément du contexte.</summary>
        public event EventHandler RemoveRequested;
        /// <summary>Crée le modèle et ses commandes Designer.</summary>
        public ChatContextChipView() { InitializeComponent(); }
                /// <summary>Configure les données et les actions disponibles pour un élément du contexte.</summary>
                /// <param name="text">Texte présenté dans la vignette.</param>
                /// <param name="canOpen">Indique si l’action principale ouvre l’élément ou le retire.</param>
                /// <param name="openTip">Info-bulle associée à l’action principale lorsqu’elle ouvre l’élément.</param>
                /// <param name="removeTip">Info-bulle associée à l’action de retrait.</param>
        public void ShowItem(string text, bool canOpen, string openTip, string removeTip)
        {
            this.canOpen = canOpen;
            open.Text = text ?? ""; remove.Visible = canOpen;
            toolTips.SetToolTip(open, canOpen ? openTip : removeTip);
            toolTips.SetToolTip(remove, removeTip);
        }
        /// <summary>Déclenche l’ouverture ou le retrait selon la capacité de l’élément.</summary>
        /// <param name="sender">Contrôle qui a déclenché l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void Open_Click(object sender, EventArgs e)
        {
            if (canOpen) OpenRequested?.Invoke(this, e);
            else RemoveRequested?.Invoke(this, e);
        }
        /// <summary>Déclenche le retrait de l’élément du contexte.</summary>
        /// <param name="sender">Contrôle qui a déclenché l’événement.</param>
        /// <param name="e">Données de l’événement WinForms.</param>
        private void Remove_Click(object sender, EventArgs e) { RemoveRequested?.Invoke(this, e); }
    }
}
