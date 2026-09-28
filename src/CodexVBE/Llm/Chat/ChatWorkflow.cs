using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    /// <summary>Mode d’exécution associé à une commande de conversation.</summary>
    internal enum ChatMode
    {
        /// <summary>Échange informatif sans modification du projet.</summary>
        Discussion,
        /// <summary>Préparation d’un plan sans appliquer les changements.</summary>
        Plan,
        /// <summary>Exécution d’une tâche qui peut modifier le projet.</summary>
        Agent
    }

    /// <summary>Commande de saisie rapide et instruction envoyée au modèle.</summary>
    internal sealed class ChatCommand
    {
        /// <summary>Jeton de commande affiché en français.</summary>
        /// <value>Jeton court reconnu dans la saisie.</value>
        public string Token { get; set; }
        /// <summary>Jeton de commande reconnu en anglais.</summary>
        /// <value>Jeton anglais accepté en entrée.</value>
        public string EnglishToken { get; set; }
        /// <summary>Jeton localisé selon la langue de l’interface.</summary>
        /// <value>Jeton traduit pour la culture active.</value>
        public string DisplayToken { get { return UiText.Culture.TwoLetterISOLanguageName == "fr" ? Token : EnglishToken; } }
        /// <summary>Clé de traduction du type de commande.</summary>
        private string kind;
        /// <summary>Clé de traduction de l’instruction associée.</summary>
        private string instruction;
        /// <summary>Type de commande traduit pour l’interface.</summary>
        /// <value>Valeur traduite depuis la clé de commande.</value>
        public string Kind { get { return UiText.Get(kind); } set { kind = value; } }
        /// <summary>Libellé du type de commande affiché.</summary>
        /// <value>Libellé affiché du type de commande.</value>
        public string DisplayKind { get { return Kind; } }
        /// <summary>Instruction traduite transmise au modèle.</summary>
        /// <value>Instruction traduite associée au jeton.</value>
        public string Instruction { get { return UiText.Get(instruction); } set { instruction = value; } }
        /// <summary>Mode d’exécution associé à la commande.</summary>
        /// <value>Mode assigné à cette commande.</value>
        public ChatMode Mode { get; set; }
        /// <summary>Commandes rapides prises en charge et leurs instructions.</summary>
        public static readonly ChatCommand[] All = {
            new ChatCommand { Token = "/expliquer", EnglishToken = "/explain", Kind = "Understand the code", Mode = ChatMode.Discussion, Instruction = "Explain the code and its dependencies without changing it." },
            new ChatCommand { Token = "/corriger", EnglishToken = "/fix", Kind = "Diagnose and fix", Mode = ChatMode.Agent, Instruction = "Diagnose and fix the problem while preserving the expected behavior." },
            new ChatCommand { Token = "/refactoriser", EnglishToken = "/refactor", Kind = "Improve the structure", Mode = ChatMode.Agent, Instruction = "Refactor the code without changing its public behavior." },
            new ChatCommand { Token = "/documenter", EnglishToken = "/document", Kind = "Document the code", Mode = ChatMode.Agent, Instruction = "Add concise and useful VBA documentation to the code." },
            new ChatCommand { Token = "/tests", EnglishToken = "/tests", Kind = "Prepare VBA tests", Mode = ChatMode.Agent, Instruction = "Propose and create isolated VBA tests. Do not execute them without an explicit request; distinguish written tests from executed tests." },
            new ChatCommand { Token = "/plan", EnglishToken = "/plan", Kind = "Prepare the steps", Mode = ChatMode.Plan, Instruction = "Prepare a concrete plan with the affected modules and checks without changing the project." }
        };
        /// <summary>Développe une commande reconnue en instruction suivie du texte saisi.</summary>
        /// <param name="text">Texte saisi, éventuellement précédé d’un jeton de commande.</param>
        /// <returns>Texte d’origine ou instruction développée suivie des arguments.</returns>
        public static string Expand(string text)
        {
            string token = text.Split(new[] { ' ', '\r', '\n' }, 2)[0];
            var command = All.FirstOrDefault(x => token == x.Token || token == x.EnglishToken);
            if (text.StartsWith("/") && command == null) throw new InvalidOperationException(UiText.Get("Unknown command. Use the / suggestions."));
            return command == null ? text : command.Instruction + "\n" + text.Substring(token.Length).Trim();
        }
    }

    /// <summary>Extrait de code joint à une entrée de conversation.</summary>
    internal sealed class ChatAttachment
    {
        /// <summary>Libellé affiché pour l’extrait joint.</summary>
        /// <value>Titre de la pièce jointe.</value>
        public string Label { get; set; }
        /// <summary>Texte de l’extrait joint.</summary>
        /// <value>Contenu textuel de la pièce jointe.</value>
        public string Text { get; set; }
        /// <summary>Nom du projet VBA source, le cas échéant.</summary>
        /// <value>Projet VBA d’origine lorsqu’il est connu.</value>
        public string Project { get; set; }
        /// <summary>Nom du module VBA source, le cas échéant.</summary>
        /// <value>Module VBA d’origine lorsqu’il est connu.</value>
        public string Module { get; set; }
        /// <summary>Empreinte du code source lors de sa lecture.</summary>
        /// <value>Empreinte SHA-256 du module lors de la capture.</value>
        public string Sha256 { get; set; }
                /// <summary>Document Monaco source pour contrôler la fraîcheur du brouillon.</summary>
/// <value>The current value represented by this member.</value>
        public string EditorDocumentId { get; set; }
        /// <summary>Première ligne source de l’extrait.</summary>
        /// <value>Numéro de la première ligne, selon l’indexation de la source.</value>
        public int StartLine { get; set; }
    }

    /// <summary>Recherche et exporte le contenu des sessions de conversation.</summary>
    internal static class ChatHistory
    {
        /// <summary>Indique si le titre, le texte ou une modification de session contient la recherche.</summary>
        /// <param name="session">Session dont le contenu est recherché ou exporté.</param>
        /// <param name="query">Texte recherché sans distinction de casse.</param>
        /// <returns>true si une occurrence est trouvée dans les champs examinés.</returns>
        public static bool Matches(ChatSessionState session, string query)
        {
            return session.Title.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                session.Entries.Any(x => (x.Text ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (x.Change != null && (x.Change.Module.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    x.Change.Before.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 || x.Change.After.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)));
        }
        /// <summary>Exporte la session en Markdown avec ses modifications et pièces jointes.</summary>
        /// <param name="session">Session dont le contenu est recherché ou exporté.</param>
        /// <returns>Document Markdown représentant la session.</returns>
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
        /// <summary>Encadre le contenu dans un bloc Markdown dont le délimiteur ne ferme pas son contenu.</summary>
        /// <param name="text">Texte saisi, éventuellement précédé d’un jeton de commande.</param>
        /// <param name="language">Identifiant de langage indiqué après le délimiteur Markdown.</param>
        /// <returns>Bloc Markdown entourant le texte sans collision avec son contenu.</returns>
        private static string Fence(string text, string language)
        {
            string fence = "```";
            while ((text ?? "").Contains(fence)) fence += "`";
            return fence + language + "\n" + text + "\n" + fence + "\n";
        }
    }
}
