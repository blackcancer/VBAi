using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private readonly Dictionary<FormCutChange, Button> formCutButtons = new Dictionary<FormCutChange, Button>();
        private FrameworkElement RenderFormCut(FormCutChange change)
        {
            var body = new StackPanel();
            body.Children.Add(new TextBlock { Text = change.Form + (string.IsNullOrEmpty(change.ParentPath) ? "" : " / " + change.ParentPath),
                Foreground = Ink("#334155"), FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap });
            body.Children.Add(new TextBlock { Text = change.ControlCount + " · " + UiText.Get("Controls cut"),
                Foreground = Ink("#64748B"), Margin = new Thickness(0, 6, 0, 8) });
            var recover = ChatButton(UiText.Get("Restore cut controls"));
            recover.ToolTip = UiText.Get("Restore names, position, size and tab order. Replaces the clipboard. Refuses intervening changes.");
            recover.Click += (sender, args) => {
                if (busy || tools == null || !tools.CanRecoverFormCut(change)) return;
                recover.IsEnabled = false;
                var result = tools.RecoverFormCut(change);
                string status = !result.Ok ? result.Error : change.Restored ? UiText.Get("Controls restored; some properties could not be verified.") :
                    Convert.ToString(((dynamic)result.Data).NativeError) ?? UiText.Get("Recovery incomplete. Inspect the Designer before another action.");
                SetStatus(status); AddTranscriptMessage(change.Restored ? "Designer" : "Erreur", status);
                RefreshFormCutCards(); SaveCurrentSession();
            };
            body.Children.Add(recover); formCutButtons[change] = recover;
            RefreshFormCutCards();
            return new Border { Child = body, Background = Ink("#FFFFFF"), BorderBrush = Ink("#E2E8F0"), BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10), Padding = new Thickness(12), Margin = new Thickness(0, 0, 4, 14) };
        }
        private void RefreshFormCutCards()
        {
            foreach (var pair in formCutButtons)
            {
                bool available = !busy && tools != null && tools.CanRecoverFormCut(pair.Key);
                pair.Value.IsEnabled = available;
                pair.Value.Content = UiText.Get(pair.Key.Restored ? "Controls restored" : pair.Key.Attempted ? "Recovery attempted" :
                    pair.Key.Owner == null || (!busy && !available) ? "Recovery unavailable" : "Restore cut controls");
            }
        }
    }
}
