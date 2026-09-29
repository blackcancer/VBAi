using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Provides the chat window implementation.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Performs the migrate provider privacy operation for ChatWindow.</summary>
        private void MigrateProviderPrivacy()
        {
            if (currentSession == null || currentSession.ReadAccessPolicyVersion >= 1) return;
            bool hadContext = transcriptEntries.Count > 0 || messages.Count > 0 ||
                !string.IsNullOrEmpty(currentSession.CodexThreadId) || !string.IsNullOrEmpty(currentSession.ResumeContext);
            codex?.Dispose(); codex = null;
            messages.Clear();
            currentSession.MessagesJson = "[]";
            currentSession.CodexThreadId = null;
            currentSession.CodexThreadHome = null;
            currentSession.ResumeContext = null;
            currentSession.BudgetPaused = false;
            currentSession.ReadAccessPolicyVersion = 1;
            if (hadContext) AddEntry(new ChatEntry { Speaker = "Assistant", Text = UiText.Get("Project privacy has been upgraded. Previous history remains available locally; the provider context starts fresh to avoid retransmitting previously shared project data.") });
            currentSession.ProviderHistoryStartIndex = transcriptEntries.Count;
        }

        /// <summary>Performs the configure project access operation for ChatWindow.</summary>
        private void ConfigureProjectAccess()
        {
            if (busy || currentSession == null || tools == null) return;
            try
            {
                EnsureCurrentScope();
                var inventory = ReadHost(scopeSession, new Request { Command = "list_projects" });
                if (!inventory.Ok) throw new InvalidOperationException(inventory.Error);
                var rows = (json.DeserializeObject(json.Serialize(inventory.Data)) as object[] ?? new object[0])
                    .OfType<IDictionary<string, object>>().ToArray();
                var choices = new List<KeyValuePair<string, string>>();
                foreach (var row in rows)
                {
                    string name = Convert.ToString(row["Name"]), path = Convert.ToString(row["FileName"]);
                    // Unsaved duplicate names cannot provide a stable authorization target.
                    if (string.IsNullOrEmpty(path) && rows.Count(item => Convert.ToString(item["Name"]) == name) != 1) continue;
                    string selector = string.IsNullOrEmpty(path) ? name : path;
                    if (string.Equals(selector, tools.BoundProject, StringComparison.OrdinalIgnoreCase)) continue;
                    choices.Add(new KeyValuePair<string, string>(selector, name + " · " + (string.IsNullOrEmpty(path) ? UiText.Get("unsaved document") : path)));
                }
                using (var dialog = new ProjectAccessWindow())
                {
                    dialog.Populate(choices, currentSession.ReadProjectGrants, currentSession.SharedContextReadAllowed);
                    if (ShowModal(dialog, this) != DialogResult.OK) return;
                    NewSession();
                    currentSession.ReadProjectGrants = dialog.SelectedProjects;
                    currentSession.SharedContextReadAllowed = dialog.SharedContext;
                    EnsureCurrentScope();
                    SaveCurrentSession();
                    SetStatus(UiText.Get("Project access configured for the new conversation."));
                }
            }
            catch (Exception error) { SetStatus(error.Message); }
        }
    }
}
