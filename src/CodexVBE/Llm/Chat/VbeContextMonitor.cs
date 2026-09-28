using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // Incremental STA polling complements native reference events, including code changes
    // without a documented notification. One module is read per tick to bound UI work.
    /// <summary>Détecte les changements de projet, modules et références par interrogation incrémentale du VBE.</summary>
    internal sealed class VbeContextMonitor
    {
        /// <summary>Exécuteur des requêtes VBE utilisées pour lire l’inventaire.</summary>
        private readonly Func<Request, Response> execute;
        /// <summary>Sérialiseur utilisé pour comparer les réponses sous forme JSON.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
        /// <summary>Requêtes de lecture différées entre deux cycles d’interrogation.</summary>
        private readonly Queue<Request> pending = new Queue<Request>();
        /// <summary>Dernières empreintes observées par catégorie ou module.</summary>
        private readonly Dictionary<string, string> fingerprints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Clé du projet actuellement surveillé.</summary>
        private string scope;
        /// <summary>Signalé lorsqu’une empreinte précédemment observée change.</summary>
        internal event Action Changed;
        /// <summary>Crée le moniteur avec l’exécuteur des commandes hôte.</summary>
        /// <param name="execute">Fonction qui exécute une requête et retourne sa réponse.</param>
        internal VbeContextMonitor(Func<Request, Response> execute) { this.execute = execute; }
        /// <summary>Traite une étape de surveillance; chaque appel lit au plus une réponse de module différée.</summary>
        /// <param name="project">Projet courant, ou valeur vide lorsqu’aucun projet n’est sélectionné.</param>
        internal void Step(string project)
        {
            if (!string.Equals(project, scope, StringComparison.OrdinalIgnoreCase)) { pending.Clear(); fingerprints.Clear(); scope = project; }
            if (pending.Count == 0)
            {
                Response projects = execute(new Request { Command = "list_projects" });
                Observe("projects", projects);
                if (string.IsNullOrEmpty(project)) return;
                Response modules = execute(new Request { Command = "list_modules", Project = project });
                Observe("modules", modules);
                if (modules.Ok)
                {
                    var items = json.DeserializeObject(json.Serialize(modules.Data)) as object[];
                    foreach (var entry in (items ?? new object[0]))
                    {
                        var fields = entry as IDictionary<string, object>;
                        if (fields != null && fields.ContainsKey("Name")) pending.Enqueue(new Request { Command = "read_module", Project = project, Module = Convert.ToString(fields["Name"]) });
                    }
                }
                pending.Enqueue(new Request { Command = "list_references", Project = project });
                return;
            }
            Request request = pending.Dequeue();
            Response result = execute(request);
            if (request.Command == "read_module" && result.Ok)
            {
                var values = json.DeserializeObject(json.Serialize(result.Data)) as IDictionary<string, object>;
                if (values != null && values.ContainsKey("Sha256")) ObserveValue("module:" + request.Module, Convert.ToString(values["Sha256"]));
            }
            else Observe(request.Command, result);
        }
        /// <summary>Enregistre la réponse réussie après sérialisation de ses données.</summary>
        /// <param name="key">Clé stable de la catégorie observée.</param>
        /// <param name="response">Réponse à comparer, ignorée si elle a échoué.</param>
        private void Observe(string key, Response response)
        { if (response.Ok) ObserveValue(key, json.Serialize(response.Data)); }
        /// <summary>Mémorise une valeur et signale uniquement les changements après la première observation.</summary>
        /// <param name="key">Clé de l’élément surveillé.</param>
        /// <param name="value">Nouvelle empreinte sérialisée.</param>
        private void ObserveValue(string key, string value)
        {
            string previous;
            bool changed = fingerprints.TryGetValue(key, out previous) && previous != value;
            fingerprints[key] = value;
            if (changed) Changed?.Invoke();
        }
    }
}
