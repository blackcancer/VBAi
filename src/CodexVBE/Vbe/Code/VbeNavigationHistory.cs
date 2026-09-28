using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;


namespace CodexVBE
{
    internal sealed class VbeNavigationHistory
    {
        private readonly dynamic vbe;
        private readonly Func<Request, Response> execute;
        private readonly Dictionary<string, Location> bookmarks = new Dictionary<string, Location>(StringComparer.OrdinalIgnoreCase);
        private readonly string bookmarkDatabase;
        private readonly Dictionary<object, string> unsavedProjectScopes = new Dictionary<object, string>();
        private readonly List<Location> back = new List<Location>();
        private readonly List<Location> forward = new List<Location>();
        private sealed class Location
        {
            public string Project, Module, Sha256;
            public int Line, Column;
        }
        internal VbeNavigationHistory(object vbe, Func<Request, Response> execute, string bookmarkDatabase = null)
        {
            this.vbe = vbe; this.execute = execute;
            this.bookmarkDatabase = bookmarkDatabase ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CodexVBE", "chat.db");
        }
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

        private string CanonicalProject(string selector)
        {
            if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("Project is required.");
            if (vbe == null) return Path.IsPathRooted(selector) ? Path.GetFullPath(selector) : selector;
            dynamic project = VbeProjectResolver.Resolve(vbe, selector);
            try { string path = (string)project.FileName; if (!string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path)) return Path.GetFullPath(path); } catch { }
            return (string)project.Name;
        }
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
        private object Navigate(Location target, bool remember)
        {
            Validate(target);
            Location current = remember ? CaptureCurrent() : null;
            Response result = execute(new Request { Command = "select_code_range", Project = target.Project, Module = target.Module, ExpectedSha256 = target.Sha256, StartLine = target.Line, StartColumn = target.Column, EndLine = target.Line, EndColumn = target.Column });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            if (remember) { forward.Clear(); if (current != null) { back.Add(current); if (back.Count > 100) back.RemoveAt(0); } }
            return result.Data;
        }
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
