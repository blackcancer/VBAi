using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    internal enum ChatMode { Discussion, Plan, Agent }

    internal sealed class ChatCommand
    {
        public string Token { get; set; }
        public string Kind { get; set; }
        public string Instruction { get; set; }
        public ChatMode Mode { get; set; }
        public static readonly ChatCommand[] All = {
            new ChatCommand { Token = "/expliquer", Kind = "Comprendre le code", Mode = ChatMode.Discussion, Instruction = "Explique le code et ses dépendances, sans le modifier." },
            new ChatCommand { Token = "/corriger", Kind = "Diagnostiquer et corriger", Mode = ChatMode.Agent, Instruction = "Diagnostique puis corrige le problème en conservant le comportement attendu." },
            new ChatCommand { Token = "/refactoriser", Kind = "Améliorer la structure", Mode = ChatMode.Agent, Instruction = "Refactorise le code sans changer son comportement public." },
            new ChatCommand { Token = "/documenter", Kind = "Documenter le code", Mode = ChatMode.Agent, Instruction = "Ajoute une documentation VBA concise et utile au code." },
            new ChatCommand { Token = "/tests", Kind = "Préparer des tests VBA", Mode = ChatMode.Agent, Instruction = "Propose puis crée des tests VBA isolés. Ne les exécute pas sans demande explicite et distingue tests écrits et exécutés." },
            new ChatCommand { Token = "/plan", Kind = "Préparer les étapes", Mode = ChatMode.Plan, Instruction = "Prépare un plan concret avec les modules concernés et les vérifications, sans modifier le projet." }
        };
        public static string Expand(string text)
        {
            var command = All.FirstOrDefault(x => text == x.Token || text.StartsWith(x.Token + " ", StringComparison.Ordinal) || text.StartsWith(x.Token + "\n", StringComparison.Ordinal));
            if (text.StartsWith("/") && command == null) throw new InvalidOperationException("Commande inconnue. Utilisez les suggestions /.");
            return command == null ? text : command.Instruction + "\n" + text.Substring(command.Token.Length).Trim();
        }
    }

    internal sealed class ChatAttachment
    {
        public string Label { get; set; }
        public string Text { get; set; }
        public string Project { get; set; }
        public string Module { get; set; }
        public string Sha256 { get; set; }
        public int StartLine { get; set; }
    }

    internal static class ChatHistory
    {
        public static bool Matches(ChatSessionState session, string query)
        {
            return session.Title.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                session.Entries.Any(x => (x.Text ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (x.Change != null && (x.Change.Module.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    x.Change.Before.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 || x.Change.After.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)));
        }
        public static string Export(ChatSessionState session)
        {
            var output = new StringBuilder("# " + session.Title + "\n\nDocument : " + session.Scope + "\n");
            foreach (var entry in session.Entries)
            {
                output.Append("\n## ").Append(entry.Speaker).Append("\n\n").Append(entry.Text).Append('\n');
                if (entry.Change != null) output.Append(entry.Change.Label).Append("\n\n").Append(Fence(entry.Change.Diff, "diff"));
                foreach (var attachment in entry.Attachments ?? new ChatAttachment[0])
                    output.Append("\n### ").Append(attachment.Label).Append('\n').Append(Fence(attachment.Text, "text"));
                if (!string.IsNullOrWhiteSpace(entry.AttachedMemory)) output.Append("\n### Mémoire jointe\n").Append(Fence(entry.AttachedMemory, "text"));
            }
            return output.ToString();
        }
        private static string Fence(string text, string language)
        {
            string fence = "```";
            while ((text ?? "").Contains(fence)) fence += "`";
            return fence + language + "\n" + text + "\n" + fence + "\n";
        }
    }
}
