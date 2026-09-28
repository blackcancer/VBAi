using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;


namespace CodexVBE
{
    /// <summary>Gère les signets et les commandes arrière/avant des emplacements de code VBA.</summary>
internal sealed class VbeNavigationHistory
    {
        /// <summary>Instance VBE facultative utilisée pour résoudre les projets et capturer le volet actif.</summary>
private readonly dynamic vbe;
        /// <summary>Exécuteur des commandes de lecture et de sélection de code.</summary>
private readonly Func<Request, Response> execute;
        /// <summary>Signets temporaires des projets non enregistrés, indexés par identité de session et nom.</summary>
private readonly Dictionary<string, Location> bookmarks = new Dictionary<string, Location>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Chemin de la base locale qui conserve les signets des projets enregistrés.</summary>
private readonly string bookmarkDatabase;
        /// <summary>Clés de session attribuées aux projets sans chemin persistant.</summary>
private readonly Dictionary<object, string> unsavedProjectScopes = new Dictionary<object, string>();
        /// <summary>Pile des emplacements précédents pour la navigation arrière.</summary>
private readonly List<Location> back = new List<Location>();
        /// <summary>Pile des emplacements futurs après un déplacement arrière.</summary>
private readonly List<Location> forward = new List<Location>();
        /// <summary>Emplacement de code identifié par projet, module, empreinte et position.</summary>
private sealed class Location
        {
            /// <summary>Stores the project,module,sha256 used by Location.</summary>
public string Project, Module, Sha256;
            /// <summary>Stores the line,column used by Location.</summary>
public int Line, Column;
        }
        /// <summary>Crée l’historique avec le VBE, le transport de commandes et le stockage des signets persistants.</summary>
        /// <param name="vbe">Instance VBE, ou null pour un transport sans automatisation directe.</param>
        /// <param name="execute">Transport des lectures et sélections de code.</param>
        /// <param name="bookmarkDatabase">Base SQLite facultative pour les projets enregistrés.</param>
internal VbeNavigationHistory(object vbe, Func<Request, Response> execute, string bookmarkDatabase = null)
        {
            this.vbe = vbe; this.execute = execute;
            this.bookmarkDatabase = bookmarkDatabase ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CodexVBE", "chat.db");
        }
        /// <summary>Ajoute, liste, supprime ou ouvre un signet pour un projet enregistré ou non enregistré.</summary>
        /// <param name="request">Action et nom du signet, ou cible lors de son ajout.</param>
        /// <returns>Signets listés, état de suppression ou résultat de navigation.</returns>
        /// <exception cref="ArgumentException">Le projet, le nom ou l’action du signet est invalide.</exception>
        /// <exception cref="InvalidOperationException">La cible est périmée ou le signet n’existe pas.</exception>
internal object Bookmark(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project)) throw new ArgumentException("Project is required.");
            string scope = request.Project;
            if (vbe != null)
            {
                dynamic project = VbeProjectResolver.Resolve(vbe, scope);
                try { string path = (string)project.FileName; if (!string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path)) scope = Path.GetFullPath(path); } catch { }
                if (!Path.IsPathRooted(scope))
                {
                    if (!unsavedProjectScopes.TryGetValue((object)project, out scope))
                    { scope = Guid.NewGuid().ToString("N"); unsavedProjectScopes.Add((object)project, scope); }
                }
            }
            if (Path.IsPathRooted(scope)) return PersistentBookmark(request, Path.GetFullPath(scope));
            if (request.Action == "list") return new { Persistence = "SessionOnlyUnsavedProject", Bookmarks = bookmarks.Where(x => x.Key.StartsWith(scope + "\0", StringComparison.OrdinalIgnoreCase)).Select(x => new { Name = x.Key.Substring(x.Key.IndexOf('\0') + 1), x.Value.Module, x.Value.Line, x.Value.Column, x.Value.Sha256 }).ToArray() };
            if (string.IsNullOrWhiteSpace(request.Query) || request.Query.Length > 100 || request.Query.Any(char.IsControl)) throw new ArgumentException("Query is the bookmark name, 1-100 characters.");
            string key = scope + "\0" + request.Query;
            if (request.Action == "remove") return new { Removed = bookmarks.Remove(key) };
            if (request.Action == "add")
            {
                if (bookmarks.Count >= 200 && !bookmarks.ContainsKey(key)) throw new InvalidOperationException("The session already has 200 bookmarks.");
                var location = new Location { Project = request.Project, Module = request.Module, Sha256 = request.ExpectedSha256, Line = request.StartLine, Column = request.StartColumn };
                Validate(location); bookmarks[key] = location; return new { Added = true, Name = request.Query };
            }
            if (request.Action == "go")
            {
                Location location;
                if (!bookmarks.TryGetValue(key, out location)) throw new InvalidOperationException("Bookmark not found in this session.");
                return Navigate(location, true);
            }
            throw new ArgumentException("Use add, list, remove or go.");
        }
        /// <summary>Effectue les opérations de signet SQLite sous la portée du chemin canonique d’un projet enregistré.</summary>
        /// <param name="request">Action et nom de signet.</param>
        /// <param name="scope">Chemin absolu du projet.</param>
        /// <returns>Signets persistants, résultat d’écriture ou navigation vers la cible.</returns>
private object PersistentBookmark(Request request, string scope)
        {
            if (request.Action != "list" && (string.IsNullOrWhiteSpace(request.Query) || request.Query.Length > 100 || request.Query.Any(char.IsControl)))
                throw new ArgumentException("Query is the bookmark name, 1-100 printable characters.");
            using (var store = new ChatSessionStore(bookmarkDatabase))
            {
                if (request.Action == "list") return new { Bookmarks = store.ListBookmarks(scope), Persistence = "SQLite", Project = scope };
                if (request.Action == "remove") return new { Removed = store.RemoveBookmark(scope, request.Query), Persistence = "SQLite" };
                if (request.Action == "add")
                {
                    var location = new Location { Project = scope, Module = request.Module, Sha256 = request.ExpectedSha256, Line = request.StartLine, Column = request.StartColumn };
                    Validate(location);
                    store.SaveBookmark(scope, new CodeBookmark { Name = request.Query, Module = location.Module, Sha256 = location.Sha256, Line = location.Line, Column = location.Column });
                    return new { Added = true, Name = request.Query, Persistence = "SQLite", Project = scope };
                }
                if (request.Action == "go")
                {
                    var bookmark = store.ListBookmarks(scope).SingleOrDefault(x => string.Equals(x.Name, request.Query, StringComparison.OrdinalIgnoreCase));
                    if (bookmark == null) throw new InvalidOperationException("Bookmark not found for this macro.");
                    return Navigate(new Location { Project = scope, Module = bookmark.Module, Sha256 = bookmark.Sha256, Line = bookmark.Line, Column = bookmark.Column }, true);
                }
                throw new ArgumentException("Use add, list, remove or go.");
            }
        }

        /// <summary>Résout un sélecteur vers le chemin du fichier du projet ou son nom lorsque celui-ci n’est pas enregistré.</summary>
        /// <param name="selector">Sélecteur de projet accepté par le résolveur VBE.</param>
        /// <returns>Identité canonique utilisée pour isoler l’historique de navigation.</returns>
private string CanonicalProject(string selector)
        {
            if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("Project is required.");
            if (vbe == null) return Path.IsPathRooted(selector) ? Path.GetFullPath(selector) : selector;
            dynamic project = VbeProjectResolver.Resolve(vbe, selector);
            try { string path = (string)project.FileName; if (!string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path)) return Path.GetFullPath(path); } catch { }
            return (string)project.Name;
        }
        /// <summary>Va vers une cible ou exécute une commande back/forward dans l’historique du projet courant.</summary>
        /// <param name="request">Action de navigation, cible et empreinte source attendue.</param>
        /// <returns>Données renvoyées par la commande de sélection VBE.</returns>
        /// <exception cref="ArgumentException">L’action ou l’emplacement demandé est invalide.</exception>
        /// <exception cref="InvalidOperationException">La pile est vide ou sa cible appartient à un autre projet.</exception>
internal object Go(Request request)
        {
            string project = CanonicalProject(request.Project);
            if (request.Action == "back" || request.Action == "forward")
            {
                var source = request.Action == "back" ? back : forward;
                var destination = request.Action == "back" ? forward : back;
                if (source.Count == 0) throw new InvalidOperationException("No " + request.Action + " VBAi navigation position.");
                Location target = source[source.Count - 1];
                if (!string.Equals(CanonicalProject(target.Project), project, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The history location belongs to another project.");
                Location current = CaptureCurrent();
                object result = Navigate(target, false);
                source.RemoveAt(source.Count - 1);
                if (current != null) { destination.Add(current); if (destination.Count > 100) destination.RemoveAt(0); }
                return result;
            }
            if (request.Action != "go") throw new ArgumentException("Use go, back or forward.");
            return Navigate(new Location { Project = project, Module = request.Module, Sha256 = request.ExpectedSha256, Line = request.StartLine, Column = request.StartColumn }, true);
        }
        /// <summary>Capture l’emplacement du volet de code actif s’il est lisible.</summary>
        /// <returns>Emplacement courant lié à l’empreinte du module, ou null si aucun volet n’est exploitable.</returns>
private Location CaptureCurrent()
        {
            Location current = null;
            try
            {
                dynamic pane = vbe.ActiveCodePane;
                if (pane != null)
                {
                    int line = 0, column = 0, endLine = 0, endColumn = 0;
                    pane.GetSelection(ref line, ref column, ref endLine, ref endColumn);
                    dynamic component = pane.CodeModule.Parent;
                    string project = (string)component.Collection.Parent.Name;
                    try { string path = (string)component.Collection.Parent.FileName; if (!string.IsNullOrWhiteSpace(path)) project = path; } catch { }
                    Response code = execute(new Request { Command = "read_module", Project = project, Module = (string)component.Name });
                    if (code.Ok) current = new Location { Project = project, Module = (string)component.Name, Line = line, Column = column, Sha256 = (string)((dynamic)code.Data).Sha256 };
                }
            }
            catch { /* A missing active pane does not prevent navigation to a valid target. */ }
            return current;
        }
        /// <summary>Valide une cible et demande au VBE de sélectionner sa position.</summary>
        /// <param name="target">Projet, module, source versionnée et position à ouvrir.</param>
        /// <param name="remember">Indique s’il faut capturer l’emplacement actif pour alimenter l’historique arrière.</param>
        /// <returns>Données retournées par la sélection de code.</returns>
        /// <exception cref="InvalidOperationException">La cible est périmée ou le transport de navigation a échoué.</exception>
private object Navigate(Location target, bool remember)
        {
            Validate(target);
            Location current = remember ? CaptureCurrent() : null;
            Response result = execute(new Request { Command = "select_code_range", Project = target.Project, Module = target.Module, ExpectedSha256 = target.Sha256, StartLine = target.Line, StartColumn = target.Column, EndLine = target.Line, EndColumn = target.Column });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            if (remember) { forward.Clear(); if (current != null) { back.Add(current); if (back.Count > 100) back.RemoveAt(0); } }
            return result.Data;
        }
        /// <summary>Vérifie l’empreinte du module et les limites de ligne/colonne avant une navigation.</summary>
        /// <param name="target">Emplacement à valider.</param>
        /// <exception cref="InvalidOperationException">La source a changé ou le module n’a pas pu être lu.</exception>
        /// <exception cref="ArgumentException">La position est en dehors du texte courant.</exception>
private void Validate(Location target)
        {
            Response code = execute(new Request { Command = "read_module", Project = target.Project, Module = target.Module });
            if (!code.Ok) throw new InvalidOperationException(code.Error);
            dynamic value = code.Data;
            if (string.IsNullOrEmpty(target.Sha256) || !string.Equals((string)value.Sha256, target.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This location is stale. Read the module and update the bookmark or navigation target.");
            string[] lines = CodeRollback.Lines((string)value.Code);
            if (target.Line < 1 || target.Line > lines.Length || target.Column < 1 || target.Column > lines[target.Line - 1].Length + 1)
                throw new ArgumentException("The location is outside the current source.");
        }
    }
}
