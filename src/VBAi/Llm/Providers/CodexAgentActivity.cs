using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{

    /// <summary>Étape réellement publiée par Codex, indépendante du contrôle qui l'affiche.</summary>
    internal sealed class CodexAgentActivity
    {

        /// <summary>Identité stable de l'étape.</summary><value>Identifiant de l'item ou de la section de résumé.</value>
        public string Id { get; set; }

        /// <summary>Catégorie native de l'action.</summary><value>Type d'item Codex.</value>
        public string Kind { get; set; }

        /// <summary>Action, outil, commande ou recherche reçue.</summary><value>Titre affichable.</value>
        public string Title { get; set; }

        /// <summary>Résumé ou résultat textuel borné.</summary><value>Détail selectable de l'étape.</value>
        public string Detail { get; set; }

        /// <summary>État natif du travail.</summary><value>inProgress, completed, failed, declined ou interrupted.</value>
        public string Status { get; set; }

        /// <summary>Durée fournie par le serveur, sans estimation locale.</summary><value>Millisecondes ou null si absentes.</value>
        public long? DurationMs { get; set; }

        /// <summary>Indique un fragment à ajouter au détail existant.</summary><value>Vrai pour les deltas de sortie ou de résumé.</value>
        public bool Append { get; set; }

        /// <summary>Convertit les actions natives prises en charge en données de présentation.</summary>
        /// <param name="item">Item reçu du protocole.</param>
        /// <param name="complete">Indique une notification de fin.</param>
        /// <returns>Action affichable, ou null pour un autre type d'item.</returns>
        internal static CodexAgentActivity FromItem(IDictionary<string, object> item, bool complete)
        {
            string kind = Text(item, "type"), title, detail;
            switch (kind)
            {
                case "commandExecution": title = ActionTitle("Running a command", Text(item, "command")); detail = Text(item, "command") + "\n" + Text(item, "cwd") + "\n" + Text(item, "aggregatedOutput"); break;
                case "fileChange":
                    title = UiText.Get("Files");
                    detail = item.TryGetValue("changes", out var changes) && changes is object[] rows ?
                        string.Join("\n", rows.OfType<IDictionary<string, object>>().Select(row => Text(row, "path"))) : ""; break;
                case "dynamicToolCall": title = ToolTitle(Text(item, "tool"), Target(item)); detail = Text(item, "tool") + "\n" + Target(item); break;
                case "mcpToolCall": title = ToolTitle(Text(item, "tool"), Target(item)); detail = Text(item, "server") + " / " + Text(item, "tool") + "\n" + Target(item); break;
                case "webSearch": title = ActionTitle("Searching the web", Text(item, "query")); detail = Text(item, "query"); break;
                case "imageView": title = ActionTitle("Viewing an image", Text(item, "path")); detail = Text(item, "path"); break;
                case "collabToolCall": title = UiText.Get("Coordinating agents"); detail = Text(item, "tool") + "\n" + Text(item, "receiverThreadId"); break;
                default: return null;
            }
            string status = Text(item, "status");
            if (complete && item.TryGetValue("success", out var success) && success is bool ok && !ok) status = "failed";
            if (complete && kind == "commandExecution" && item.TryGetValue("exitCode", out var exit) && exit != null && Convert.ToInt32(exit) != 0) status = "failed";
            long? duration = null;
            if (item.TryGetValue("durationMs", out var raw) && raw != null && long.TryParse(Convert.ToString(raw), out var milliseconds) && milliseconds >= 0) duration = milliseconds;
            return new CodexAgentActivity
            {
                Id = Text(item, "id"),
                Kind = kind,
                Title = Limit(title),
                Detail = Limit(detail),
                Status = string.IsNullOrEmpty(status) ? (complete ? "completed" : "inProgress") : status,
                DurationMs = duration
            };
        }

        /// <summary>Describes a declared tool action without exposing arbitrary arguments or inventing reasoning.</summary>
        /// <param name="tool">Technical tool name supplied by the provider.</param>
        /// <param name="target">Allowlisted public target fields, when supplied.</param>
        /// <returns>Localized action caption and its declared target.</returns>
        internal static string ToolTitle(string tool, string target = null)
        {
            string name = (tool ?? "").Trim();
            int separator = Math.Max(name.LastIndexOf('/'), name.LastIndexOf('.'));
            if (separator >= 0) name = name.Substring(separator + 1).Trim();
            string key;
            switch (name.ToLowerInvariant())
            {
                case "read_module": case "find_code": case "list_procedures": case "project_symbols":
                    key = "Reading VBA code"; break;
                case "write_module": case "replace_procedure": case "remove_procedure": case "insert_code_file": case "apply_code_edit":
                    key = "Updating VBA code"; break;
                case "invoke_monaco": case "open_editor": case "editor_layout": case "select_code": case "select_code_range": case "monaco_open": case "monaco_read": case "monaco_navigate":
                    key = "Working in the editor"; break;
                case "list_projects": case "list_modules": case "project_properties": case "component_properties":
                    key = "Inspecting VBA project"; break;
                default: key = "Using a tool"; break;
            }
            return UiText.Get(key) + (string.IsNullOrWhiteSpace(target) ? "" : " · " + Limit(target));
        }

        /// <summary>Combines a localized action with bounded metadata already published by the provider.</summary>
        /// <param name="key">Localized action resource key.</param>
        /// <param name="target">Declared command, search query or image path.</param>
        /// <returns>Primary activity caption without multiline or unbounded target text.</returns>
        private static string ActionTitle(string key, string target) => UiText.Get(key) +
            (string.IsNullOrWhiteSpace(target) ? "" : " · " + PublicTarget(target));

        /// <summary>Formats existing session titles while preserving meaningful public reasoning summaries.</summary>
        /// <returns>Localized primary caption, independent of the retained protocol identifier.</returns>
        internal string DisplayTitle()
        {
            if (Kind == "reasoning" && (string.IsNullOrWhiteSpace(Title) || Title == "Reasoning" || Title == UiText.Get("Reasoning") || Title == "Reasoning · summary" || Title == UiText.Get("Reasoning · summary")))
            {
                // This is an excerpt of the public summary, never generated or private reasoning.
                string heading = (Detail ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line))?.Trim(' ', '\t', '#', '*', '`', '>');
                if (!string.IsNullOrWhiteSpace(heading)) return heading.Length <= 96 ? heading : heading.Substring(0, 93) + "…";
                return UiText.Get("Reasoning");
            }
            if ((Kind == "dynamicToolCall" || Kind == "mcpToolCall") && IsTechnicalToolTitle()) return ToolTitle(Title);
            return string.IsNullOrWhiteSpace(Title) ? UiText.Get(Kind == "reasoning" ? "Reasoning" : "Tool") : Title;
        }

        /// <summary>Keeps legacy technical identifiers in selectable details, away from the primary caption.</summary>
        /// <returns>Public activity detail with a legacy identifier when applicable.</returns>
        internal string DisplayDetail() =>
            (IsTechnicalToolTitle() && (Kind == "dynamicToolCall" || Kind == "mcpToolCall") ? Title + "\n" : "") + Detail;

        /// <summary>Recognizes identifiers in older sessions without reclassifying a public natural-language caption.</summary>
        /// <returns>True for a technical tool identifier.</returns>
        internal bool IsTechnicalToolTitle() => !string.IsNullOrWhiteSpace(Title) && Title.IndexOf('_') >= 0 &&
            Title.Split('/').All(part => part.Trim().All(character => char.IsLetterOrDigit(character) || character == '_' || character == '.' || character == '-'));

        /// <summary>Ne reprend que les identités publiques de la cible d'un outil, sans ses arguments secrets ou son code.</summary>
        /// <param name="item">Item d'outil.</param>
        /// <returns>Cible déclarée par le fournisseur.</returns>
        private static string Target(IDictionary<string, object> item)
        {
            if (!item.TryGetValue("arguments", out var raw) || !(raw is IDictionary<string, object> arguments)) return "";
            return string.Join(" · ", new[] { "Project", "Module", "Form", "Procedure", "ProcedureName", "ControlPath", "Path" }
                .Select(key => arguments.TryGetValue(key, out var value) && value is string text ? PublicTarget(text) : "").Where(value => !string.IsNullOrWhiteSpace(value)));
        }

        /// <summary>Bounds a declared identity to one line so it cannot replace the action heading.</summary>
        /// <param name="text">Public target identity supplied by the provider.</param>
        /// <returns>At most 128 characters without control characters.</returns>
        private static string PublicTarget(string text)
        {
            string value = new string(text.TakeWhile(character => character != '\r' && character != '\n').Where(character => !char.IsControl(character)).ToArray()).Trim();
            return value.Length <= 128 ? value : value.Substring(0, 125) + "…";
        }

        /// <summary>Lit un champ texte optionnel.</summary><param name="item">Objet source.</param><param name="key">Champ demandé.</param><returns>Valeur ou chaîne vide.</returns>
        private static string Text(IDictionary<string, object> item, string key) => item != null && item.TryGetValue(key, out var value) ? Convert.ToString(value) : "";

        /// <summary>Borne les détails conservés dans la session.</summary><param name="text">Texte reçu.</param><returns>Au plus 16 384 caractères.</returns>
        internal static string Limit(string text) => string.IsNullOrEmpty(text) ? "" : text.Substring(0, Math.Min(text.Length, 16384));
    }
}
