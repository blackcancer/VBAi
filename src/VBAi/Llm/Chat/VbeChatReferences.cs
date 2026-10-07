using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Identifie un projet, module ou élément de code VBE et conserve son contexte de sélection.</summary>
    internal sealed class VbeChatReference
    {

        /// <summary>Nom du projet associé à la référence.</summary>
        public string Project;

        /// <summary>Nom du module, nul pour une référence au projet entier.</summary>
        public string Module;

        /// <summary>Nom de la procédure, nul pour une référence au module ou au projet.</summary>
        public string Name;

        /// <summary>Portée d’une déclaration, nulle pour les procédures et les symboles de module.</summary>
        public string DeclarationScope;

        /// <summary>Colonne de déclaration, indexée à partir de un.</summary>
        public int DeclarationColumn;

        /// <summary>Obtient ou définit la catégorie affichée de la référence.</summary>
        /// <value>Catégorie de référence.</value>
        public string Kind { get; set; }

        /// <summary>Obtient la catégorie traduite lorsque la référence désigne le projet.</summary>
        /// <value>Catégorie localisée du projet ou catégorie d’origine.</value>
        public string DisplayKind { get { return Kind == "Projet" ? UiText.Get("Project") : Kind; } }

        /// <summary>Index COM du type de procédure propriété.</summary>
        public int ProcKind;

        /// <summary>Première ligne de la procédure dans le module, indexée à partir de un.</summary>
        public int StartLine;

        /// <summary>Dernière ligne de la procédure dans le module, indexée à partir de un.</summary>
        public int EndLine;

        /// <summary>Empreinte du module à la découverte de la procédure.</summary>
        public string Sha256;

        /// <summary>Obtient le jeton textuel inséré dans une conversation.</summary>
        /// <value>Valeur de <see cref="Token"/>.</value>
        public string DisplayToken { get { return Token; } }

        /// <summary>Obtient le chemin textuel unique de la référence, préfixé par # pour un projet ou @ pour un élément nommé.</summary>
        /// <value>Jeton d’identification construit à partir du projet, du module et du nom.</value>
        public string Token
        {
            get
            {
                string path = (Name == null ? "#" : "@") + Project;
                if (Module != null) path += "." + Module;
                if (!string.IsNullOrEmpty(DeclarationScope)) path += "." + DeclarationScope;
                if (Name != null) path += "." + Name;
                if (DeclarationColumn > 0) path += ":L" + StartLine;
                if (Kind == "Property") path += ":" + new[] { "", "Let", "Set", "Get" }[ProcKind];
                return path;
            }
        }

        /// <summary>Obtient le jeton suivi de sa catégorie d’affichage.</summary>
        /// <value>Libellé de référence destiné à l’interface.</value>
        public string Display { get { return Token + "  —  " + DisplayKind; } }

        /// <summary>Retourne le libellé destiné à l’affichage.</summary>
        /// <returns>Valeur de <see cref="Display"/>.</returns>
        public override string ToString() { return Display; }
    }

    /// <summary>Charge progressivement les projets, modules et procédures disponibles pour les références de conversation.</summary>
    internal sealed class VbeChatReferences
    {

        /// <summary>Session VBE utilisée pour lire et naviguer dans le code.</summary>
        private readonly VbeSession session;

        /// <summary>Exécute les requêtes de lecture et navigation via la session hôte par défaut.</summary>
        internal Func<Request, Response> Execute;

        /// <summary>Sérialiseur utilisé pour convertir les résultats de pont en dictionnaires simples.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();

        /// <summary>Projets dont les modules doivent encore être énumérés.</summary>
        private readonly Queue<string> projects = new Queue<string>();

        /// <summary>Modules dont les procédures doivent encore être chargées.</summary>
        private readonly Queue<VbeChatReference> pending = new Queue<VbeChatReference>();

        /// <summary>Références découvertes et exposées à la recherche.</summary>
        private readonly List<VbeChatReference> entries = new List<VbeChatReference>();

        /// <summary>Se produit lorsque l’état de chargement ou la liste des références change.</summary>
        public event Action Changed;

        /// <summary>Obtient la dernière erreur de lecture, le cas échéant.</summary>
        /// <value>Message de la dernière exception capturée, ou nul.</value>
        public string Error { get; private set; }

        /// <summary>Indique si des projets ou modules restent à charger.</summary>
        /// <value><see langword="true"/> lorsqu’une file de chargement n’est pas vide.</value>
        public bool IsLoading { get { return projects.Count > 0 || pending.Count > 0; } }

        /// <summary>Crée le chargeur progressif lié à la session VBE.</summary>
        /// <param name="session">Session utilisée pour les commandes VBE.</param>
        public VbeChatReferences(VbeSession session) { this.session = session; Execute = request => session.Execute(request); }

        /// <summary>Obtient la liste actuelle des projets, modules et procédures découverts.</summary>
        /// <value>Références actuellement disponibles.</value>
        public IList<VbeChatReference> Entries { get { return entries; } }

        /// <summary>Efface les résultats précédents, charge la liste des projets et prépare le parcours progressif.</summary>
        public void Refresh()
        {
            projects.Clear();
            pending.Clear();
            entries.Clear();
            Error = null;
            try
            {
                foreach (var project in Read("list_projects", null, null))
                {
                    string projectName = Field(project, "Name");
                    entries.Add(new VbeChatReference { Project = projectName, Kind = "Projet" });
                    projects.Enqueue(projectName);
                }
            }
            catch (Exception ex) { Error = ex.Message; }
            Changed?.Invoke();
        }

        /// <summary>Traite un projet ou un module en attente et notifie les observateurs après l’étape.</summary>
        public void Step()
        {
            if (projects.Count > 0)
            {
                string project = projects.Dequeue();
                try
                {
                    foreach (var moduleInfo in Read("list_modules", project, null))
                    {
                        var item = new VbeChatReference
                        {
                            Project = project,
                            Module = Field(moduleInfo, "Name"),
                            Kind = "Module"
                        };
                        entries.Add(item);
                        pending.Enqueue(item);
                    }
                }
                catch (Exception ex) { Error = project + " : " + ex.Message; }
                Changed?.Invoke();
                return;
            }
            if (pending.Count == 0) return;
            var module = pending.Dequeue();
            try
            {
                Response response = Execute(new Request
                {
                    Command = "list_procedures",
                    Project = module.Project,
                    Module = module.Module
                });
                if (!response.Ok) throw new InvalidOperationException(response.Error);
                var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
                module.Sha256 = Field(data, "Sha256");
                var procedures = data != null && data.TryGetValue("Procedures", out object raw) ? raw as object[] : null;
                if (procedures != null)
                    foreach (var value in procedures)
                    {
                        if (!(value is IDictionary<string, object> procedure)) continue;
                        int kind = Convert.ToInt32(procedure["Kind"]);
                        string declaration = Field(procedure, "Declaration");
                        entries.Add(new VbeChatReference
                        {
                            Project = module.Project,
                            Module = module.Module,
                            Name = Field(procedure, "Name"),
                            Kind = kind == 0
                                ? (declaration.IndexOf("Function", StringComparison.OrdinalIgnoreCase) >= 0 ? "Function" : "Sub")
                                : "Property",
                            ProcKind = kind,
                            StartLine = Convert.ToInt32(procedure["StartLine"]),
                            EndLine = Convert.ToInt32(procedure["EndLine"]),
                            Sha256 = module.Sha256
                        });
                    }
                // Read declarations once per module; physical positions retain their own source SHA.
                Response sourceResponse = Execute(new Request { Command = "read_module", Project = module.Project, Module = module.Module });
                if (!sourceResponse.Ok) throw new InvalidOperationException(sourceResponse.Error);
                var sourceData = json.DeserializeObject(json.Serialize(sourceResponse.Data)) as IDictionary<string, object>;
                string code = Field(sourceData, "Code"), sourceSha = Field(sourceData, "Sha256");
                foreach (var declaration in VbaDeclarationIndex.Read(code))
                    entries.Add(new VbeChatReference
                    {
                        Project = module.Project,
                        Module = module.Module,
                        Name = declaration.Name,
                        Kind = declaration.Kind,
                        DeclarationScope = declaration.Scope == "Module" ? null : declaration.Scope,
                        DeclarationColumn = declaration.Column,
                        StartLine = declaration.Line,
                        EndLine = declaration.Line,
                        Sha256 = sourceSha
                    });
            }
            catch (Exception ex) { Error = module.Token + " : " + ex.Message; }
            Changed?.Invoke();
        }

        /// <summary>Recherche les références dont le jeton ou le nom contient la chaîne indiquée.</summary>
        /// <param name="query">Texte à rechercher dans les références.</param>
        /// <returns>Résultats ordonnés par correspondance de préfixe puis jeton, limités à quarante entrées.</returns>
        public IEnumerable<VbeChatReference> Match(string query)
        {
            return MatchPrefix(query, '\0');
        }

        /// <summary>Recherche les références en restreignant éventuellement les résultats au type désigné par le préfixe.</summary>
        /// <param name="query">Texte recherché après retrait des préfixes # et @.</param>
        /// <param name="prefix">Nul pour tous les types, @ pour les éléments nommés, ou # pour les projets et modules.</param>
        /// <returns>Résultats correspondants triés et limités à quarante entrées.</returns>
        public IEnumerable<VbeChatReference> MatchPrefix(string query, char prefix)
        {
            string term = query.TrimStart('#', '@');
            return entries.Where(item => (prefix == '\0' || (prefix == '@' ? item.Name != null : item.Name == null)) &&
                item.Token.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(item => item.Token.Substring(1).StartsWith(term, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(item => item.Token, StringComparer.OrdinalIgnoreCase).Take(40);
        }

        /// <summary>Demande au VBE d’ouvrir le code du module ou de sélectionner la procédure référencée.</summary>
        /// <param name="item">Référence de navigation.</param>
        /// <returns>Réponse du pont VBE ; une référence de projet seul retourne une erreur de sélection.</returns>
        public Response Navigate(VbeChatReference item)
        {
            if (item.Module == null)
                return Response.Failure(UiText.Get("Select a module or procedure in this project to open its code."));
            Response response = Execute(new Request { Command = "read_module", Project = item.Project, Module = item.Module });
            if (!response.Ok) return response;
            var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
            if (item.DeclarationColumn > 0)
            {
                if (!string.Equals(item.Sha256, Field(data, "Sha256"), StringComparison.OrdinalIgnoreCase))
                    return Response.Failure("The declaration changed since discovery; refresh its reference before navigating.");
                return Execute(new Request
                {
                    Command = "select_code",
                    Project = item.Project,
                    Module = item.Module,
                    StartLine = item.StartLine,
                    StartColumn = item.DeclarationColumn,
                    EndColumn = item.DeclarationColumn + item.Name.Length,
                    Expression = item.Name,
                    ExpectedSha256 = item.Sha256
                });
            }
            return Execute(new Request
            {
                Command = item.Name == null ? "select_code" : "select_procedure",
                Project = item.Project,
                Module = item.Module,
                Procedure = item.Name,
                ProcKind = item.ProcKind,
                StartLine = 1,
                ExpectedSha256 = Field(data, "Sha256")
            });
        }

        /// <summary>Résout une référence en code du projet, du module ou de la plage de procédure mémorisée.</summary>
        /// <param name="item">Référence à résoudre.</param>
        /// <returns>Texte descriptif et contenu de code associé.</returns>
        /// <exception cref="InvalidOperationException">La lecture échoue, l’empreinte du module a changé ou sa plage n’est plus valide.</exception>
        public string Resolve(VbeChatReference item)
        {
            if (item.Module == null)
            {
                var modules = Read("list_modules", item.Project, null);
                return item.Token + " (projet vivant ; " + modules.Length + " composants)\n" +
                    string.Join("\n", modules.Select(value => Field(value, "Name") +
                        " (type " + Field(value, "Type") + ", " + Field(value, "Lines") + " lignes)"));
            }
            Response response = Execute(new Request
            {
                Command = "read_module",
                Project = item.Project,
                Module = item.Module
            });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
            string code = Field(data, "Code");
            string sha = Field(data, "Sha256");
            if (item.Name == null) { item.Sha256 = sha; return item.Token + " (SHA-256 " + sha + ")\n" + code; }
            if (!string.Equals(sha, item.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(item.Token + UiText.Get(" changed since selection. Remove the token and select it again with #."));
            string[] lines = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            int start = Math.Max(1, item.StartLine);
            int end = Math.Min(lines.Length, item.EndLine);
            if (end < start) throw new InvalidOperationException(item.Token + " n'a plus de plage valide.");
            return item.Token + " (lignes " + start + "-" + end + ", SHA-256 " + sha + ")\n" +
                string.Join("\n", lines.Skip(start - 1).Take(end - start + 1));
        }

        /// <summary>Exécute une commande VBE et convertit son tableau de résultats en dictionnaires.</summary>
        /// <param name="command">Commande à transmettre à la session.</param>
        /// <param name="project">Projet ciblé, si applicable.</param>
        /// <param name="module">Module ciblé, si applicable.</param>
        /// <returns>Éléments objet retournés, ou tableau vide si la charge utile n’est pas un tableau.</returns>
        /// <exception cref="InvalidOperationException">La commande VBE échoue.</exception>
        private IDictionary<string, object>[] Read(string command, string project, string module)
        {
            Response response = Execute(new Request { Command = command, Project = project, Module = module });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            return !(json.DeserializeObject(json.Serialize(response.Data)) is object[] array) ? new IDictionary<string, object>[0] :
                array.OfType<IDictionary<string, object>>().ToArray();
        }

        /// <summary>Lit une valeur de dictionnaire et la convertit en chaîne, en renvoyant une chaîne vide si elle manque.</summary>
        /// <param name="value">Dictionnaire de données.</param>
        /// <param name="name">Clé à lire.</param>
        /// <returns>Valeur convertie ou chaîne vide.</returns>
        private static string Field(IDictionary<string, object> value, string name)
        {
            return value != null && value.TryGetValue(name, out object raw) ? Convert.ToString(raw) : "";
        }
    }
}
