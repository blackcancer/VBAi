using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
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
                case "commandExecution": title = Text(item, "command"); detail = Text(item, "cwd") + "\n" + Text(item, "aggregatedOutput"); break;
                case "fileChange":
                    title = UiText.Get("Files");
                    detail = item.TryGetValue("changes", out var changes) && changes is object[] rows ?
                        string.Join("\n", rows.OfType<IDictionary<string, object>>().Select(row => Text(row, "path"))) : ""; break;
                case "dynamicToolCall": title = Text(item, "tool"); detail = Target(item); break;
                case "mcpToolCall": title = Text(item, "server") + " / " + Text(item, "tool"); detail = Target(item); break;
                case "webSearch": title = Text(item, "query"); detail = ""; break;
                case "imageView": title = Text(item, "path"); detail = ""; break;
                case "collabToolCall": title = Text(item, "tool"); detail = Text(item, "receiverThreadId"); break;
                default: return null;
            }
            string status = Text(item, "status");
            if (complete && item.TryGetValue("success", out var success) && success is bool ok && !ok) status = "failed";
            if (complete && kind == "commandExecution" && item.TryGetValue("exitCode", out var exit) && exit != null && Convert.ToInt32(exit) != 0) status = "failed";
            long? duration = null;
            if (item.TryGetValue("durationMs", out var raw) && raw != null && long.TryParse(Convert.ToString(raw), out var milliseconds) && milliseconds >= 0) duration = milliseconds;
            return new CodexAgentActivity { Id = Text(item, "id"), Kind = kind, Title = Limit(title), Detail = Limit(detail),
                Status = string.IsNullOrEmpty(status) ? (complete ? "completed" : "inProgress") : status, DurationMs = duration };
        }
        /// <summary>Ne reprend que les identités publiques de la cible d'un outil, sans ses arguments secrets ou son code.</summary>
        /// <param name="item">Item d'outil.</param>
        /// <returns>Cible déclarée par le fournisseur.</returns>
        private static string Target(IDictionary<string, object> item)
        {
            if (!item.TryGetValue("arguments", out var raw) || !(raw is IDictionary<string, object> arguments)) return "";
            return string.Join(" · ", new[] { "Project", "Module", "ProcedureName", "ControlPath", "Path" }
                .Select(key => Text(arguments, key)).Where(value => !string.IsNullOrWhiteSpace(value)));
        }
        /// <summary>Lit un champ texte optionnel.</summary><param name="item">Objet source.</param><param name="key">Champ demandé.</param><returns>Valeur ou chaîne vide.</returns>
        private static string Text(IDictionary<string, object> item, string key) => item != null && item.TryGetValue(key, out var value) ? Convert.ToString(value) : "";
        /// <summary>Borne les détails conservés dans la session.</summary><param name="text">Texte reçu.</param><returns>Au plus 16 384 caractères.</returns>
        internal static string Limit(string text) => string.IsNullOrEmpty(text) ? "" : text.Substring(0, Math.Min(text.Length, 16384));
    }
}
