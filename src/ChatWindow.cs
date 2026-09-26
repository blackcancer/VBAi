using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class ChatWindow : Form
    {
        private readonly TextBox transcript;
        private readonly TextBox prompt;
        private readonly Button send;
        private readonly Label status;
        private readonly ComboBox providerPicker;
        private readonly Button configure;
        private readonly LlmSettings settings;
        private readonly LlmVbeTools tools;
        private CodexAppServerClient codex;
        private readonly List<object> messages = new List<object>();
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private bool busy;

        public ChatWindow(VbeSession session)
        {
            Text = "CodexVBE — Assistant";
            Width = 620;
            Height = 720;
            MinimumSize = new Size(420, 400);
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = true;
            Font = new Font("Segoe UI", 9F);

            transcript = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BackColor = Color.White, Font = new Font("Consolas", 9F) };
            prompt = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, Height = 100 };
            send = new Button { Text = "Envoyer", Dock = DockStyle.Right, Width = 100 };
            status = new Label { Text = "Modèle non configuré ou prêt à répondre", Dock = DockStyle.Top,
                Height = 25, AutoEllipsis = true };
            providerPicker = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            providerPicker.Items.AddRange(LlmProvider.All);
            configure = new Button { Text = "Configuration…", Dock = DockStyle.Right, Width = 130 };
            var toolbar = new Panel { Dock = DockStyle.Top, Height = 32 };
            toolbar.Controls.Add(providerPicker);
            toolbar.Controls.Add(configure);
            var composer = new Panel { Dock = DockStyle.Bottom, Height = 110 };
            composer.Controls.Add(prompt);
            composer.Controls.Add(send);
            Controls.Add(transcript);
            Controls.Add(composer);
            Controls.Add(status);
            Controls.Add(toolbar);

            tools = new LlmVbeTools(session, this);
            try { settings = LlmSettings.Load(); }
            catch (Exception ex) { LoadLog.Write("LLM settings load failed: " + ex.Message); settings = new LlmSettings(); }
            providerPicker.SelectedIndexChanged += (sender, args) => ResetConversation();
            int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
            providerPicker.SelectedIndex = selected < 0 ? 0 : selected;
            configure.Click += (sender, args) => ShowSettings();
            send.Click += async (sender, args) => await SendAsync();
        }

        private void ResetConversation()
        {
            codex?.Dispose();
            codex = null;
            messages.Clear();
            messages.Add(new { role = "system", content =
                "You are a VBE assistant. Use tools to inspect the live VBA project before stating facts about it. " +
                "Read current code or form state before edits; pass the returned revision to every edit. " +
                "Do not invent project, module, form or control names. Keep answers concise and in the user's language." });
            transcript.Text = "CodexVBE est connecté au VBE. Écrivez une demande puis cliquez sur Envoyer.\r\n";
            status.Text = ((LlmProvider)providerPicker.SelectedItem).IsCodex ?
                "Codex : utilise la connexion ChatGPT du CLI local." :
                ((LlmProvider)providerPicker.SelectedItem).Available ?
                "Configurez le modèle du fournisseur sélectionné, puis envoyez une demande." :
                "Ce fournisseur n'est pas encore implémenté.";
            settings.ProviderName = ((LlmProvider)providerPicker.SelectedItem).Name;
            try { settings.Save(); }
            catch (Exception saveError) { LoadLog.Write("LLM settings save failed: " + saveError.Message); }
        }

        public void ShowSettings()
        {
            if (busy) return;
            using (var dialog = new LlmSettingsWindow(settings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                int selected = Array.FindIndex(LlmProvider.All, item => item.Name == settings.ProviderName);
                if (selected >= 0 && providerPicker.SelectedIndex != selected) providerPicker.SelectedIndex = selected;
                else ResetConversation();
            }
        }

        private void Append(string speaker, string content)
        {
            if (IsDisposed || transcript.IsDisposed) return;
            transcript.AppendText("\r\n" + speaker + " : " + content + "\r\n");
        }

        private void SetStatus(string text)
        {
            if (!IsDisposed && !status.IsDisposed) status.Text = text;
        }

        private async Task SendAsync()
        {
            string question = prompt.Text.Trim();
            if (busy || question.Length == 0) return;
            busy = true;
            send.Enabled = false;
            providerPicker.Enabled = false;
            configure.Enabled = false;
            prompt.Clear();
            Append("Vous", question);
            int checkpoint = messages.Count;
            var provider = (LlmProvider)providerPicker.SelectedItem;
            settings.ProviderName = provider.Name;
            try { settings.Save(); }
            catch (Exception saveError) { LoadLog.Write("LLM settings save failed: " + saveError.Message); }
            if (!provider.IsCodex) messages.Add(new { role = "user", content = question });
            try
            {
                if (provider.IsCodex)
                {
                    if (codex == null)
                        codex = new CodexAppServerClient(SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext(),
                            tools, SetStatus, settings);
                    SetStatus("Codex — en cours");
                    Append("Assistant", await codex.TurnAsync(question));
                    SetStatus("Codex — prêt");
                    return;
                }
                using (var client = new LlmChatClient((LlmProvider)providerPicker.SelectedItem, settings))
                {
                    SetStatus(client.DisplayName + " — en cours");
                    for (int turn = 0; turn < 8; turn++)
                    {
                        var message = await client.CompleteAsync(messages, LlmVbeTools.Definitions);
                        messages.Add(message);
                        object rawCalls;
                        var calls = message.TryGetValue("tool_calls", out rawCalls) ? rawCalls as object[] : null;
                        if (calls == null || calls.Length == 0)
                        {
                            string answer = message.ContainsKey("content") ? Convert.ToString(message["content"]) : "";
                            Append("Assistant", string.IsNullOrWhiteSpace(answer) ? "Aucune réponse textuelle." : answer);
                            SetStatus(client.DisplayName + " — prêt");
                            return;
                        }
                        foreach (object rawCall in calls)
                        {
                            var call = rawCall as IDictionary<string, object>;
                            var function = call != null && call.ContainsKey("function") ? call["function"] as IDictionary<string, object> : null;
                            if (function == null || !call.ContainsKey("id")) throw new InvalidOperationException("Invalid tool call.");
                            string name = Convert.ToString(function["name"]);
                            string arguments = Convert.ToString(function["arguments"]);
                            Append("Outil", name);
                            string result = tools.Invoke(name, arguments);
                            messages.Add(new { role = "tool", tool_call_id = Convert.ToString(call["id"]), content = result });
                        }
                    }
                    throw new InvalidOperationException("The assistant exceeded the tool-call limit.");
                }
            }
            catch (Exception ex)
            {
                if (provider.IsCodex) { codex?.Dispose(); codex = null; }
                // A failed request must not leave an orphaned tool call in the next API request.
                messages.RemoveRange(checkpoint, messages.Count - checkpoint);
                Append("Erreur", ex.Message);
                SetStatus("Erreur — vérifiez la configuration et réessayez");
            }
            finally
            {
                busy = false;
                if (!IsDisposed) { send.Enabled = true; providerPicker.Enabled = true; configure.Enabled = true; }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { codex?.Dispose(); codex = null; }
            base.Dispose(disposing);
        }
    }
}
