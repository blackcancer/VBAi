using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private readonly Dictionary<FormCutChange, System.Windows.Forms.Button> formCutButtons = new Dictionary<FormCutChange, System.Windows.Forms.Button>();
        private FrameworkElement RenderFormCut(FormCutChange change)
        {
            var card = new ChatFormRecoveryView();
            card.title.Text = change.Form + (string.IsNullOrEmpty(change.ParentPath) ? "" : " / " + change.ParentPath);
            card.count.Text = change.ControlCount + " · " + UiText.Get("Controls cut");
            card.recover.Click += (sender,args) => {
                if (busy || tools == null || !tools.CanRecoverFormCut(change)) return;
                card.recover.Enabled = false;
                var result = tools.RecoverFormCut(change);
                string status = !result.Ok ? result.Error : change.Restored ? UiText.Get("Controls restored; some properties could not be verified.") :
                    Convert.ToString(((dynamic)result.Data).NativeError) ?? UiText.Get("Recovery incomplete. Inspect the Designer before another action.");
                SetStatus(status); AddTranscriptMessage(change.Restored ? "Designer" : "Erreur", status);
                RefreshFormCutCards(); SaveCurrentSession();
            };
            formCutButtons[change] = card.recover; RefreshFormCutCards();
            return new ChatDesignerHost(card) { Margin = new Thickness(0,0,4,14) };
        }
        private void RefreshFormCutCards()
        {
            foreach (var pair in formCutButtons)
            {
                bool available = !busy && tools != null && tools.CanRecoverFormCut(pair.Key);
                pair.Value.Enabled = available;
                pair.Value.Text = UiText.Get(pair.Key.Restored ? "Controls restored" : pair.Key.Attempted ? "Recovery attempted" :
                    pair.Key.Owner == null || (!busy && !available) ? "Recovery unavailable" : "Restore cut controls");
            }
        }
    }
}
