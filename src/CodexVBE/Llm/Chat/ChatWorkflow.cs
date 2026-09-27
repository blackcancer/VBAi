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
        public string EnglishToken { get; set; }
        public string DisplayToken { get { return UiText.Culture.TwoLetterISOLanguageName == "fr" ? Token : EnglishToken; } }
        private string kind;
        private string instruction;
        public string Kind { get { return UiText.Get(kind); } set { kind = value; } }
        public string DisplayKind { get { return Kind; } }
        public string Instruction { get { return UiText.Get(instruction); } set { instruction = value; } }
        public ChatMode Mode { get; set; }
        public static readonly ChatCommand[] All = {
            new ChatCommand { Token = "/expliquer", EnglishToken = "/explain", Kind = "Understand the code", Mode = ChatMode.Discussion, Instruction = "Explain the code and its dependencies without changing it." },
            new ChatCommand { Token = "/corriger", EnglishToken = "/fix", Kind = "Diagnose and fix", Mode = ChatMode.Agent, Instruction = "Diagnose and fix the problem while preserving the expected behavior." },
            new ChatCommand { Token = "/refactoriser", EnglishToken = "/refactor", Kind = "Improve the structure", Mode = ChatMode.Agent, Instruction = "Refactor the code without changing its public behavior." },
            new ChatCommand { Token = "/documenter", EnglishToken = "/document", Kind = "Document the code", Mode = ChatMode.Agent, Instruction = "Add concise and useful VBA documentation to the code." },
            new ChatCommand { Token = "/tests", EnglishToken = "/tests", Kind = "Prepare VBA tests", Mode = ChatMode.Agent, Instruction = "Propose and create isolated VBA tests. Do not execute them without an explicit request; distinguish written tests from executed tests." },
            new ChatCommand { Token = "/plan", EnglishToken = "/plan", Kind = "Prepare the steps", Mode = ChatMode.Plan, Instruction = "Prepare a concrete plan with the affected modules and checks without changing the project." }
        };
        public static string Expand(string text)
        {
            string token = text.Split(new[] { ' ', '\r', '\n' }, 2)[0];
            var command = All.FirstOrDefault(x => token == x.Token || token == x.EnglishToken);
            if (text.StartsWith("/") && command == null) throw new InvalidOperationException(UiText.Get("Unknown command. Use the / suggestions."));
            return command == null ? text : command.Instruction + "\n" + text.Substring(token.Length).Trim();
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
            var output = new StringBuilder("# " + session.DisplayTitle + "\n\nDocument: " + session.Scope + "\n");
            foreach (var entry in session.Entries)
            {
                output.Append("\n## ").Append(UiText.Speaker(entry.Speaker)).Append("\n\n").Append(entry.Text).Append('\n');
                if (entry.Change != null) output.Append(entry.Change.Label).Append("\n\n").Append(Fence(entry.Change.Diff, "diff"));
                foreach (var attachment in entry.Attachments ?? new ChatAttachment[0])
                    output.Append("\n### ").Append(attachment.Label).Append('\n').Append(Fence(attachment.Text, "text"));
                if (!string.IsNullOrWhiteSpace(entry.AttachedMemory)) output.Append(UiText.Get("\n### Attached memory\n")).Append(Fence(entry.AttachedMemory, "text"));
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
