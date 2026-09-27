using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class VbeChatReference
    {
        public string Project;
        public string Module;
        public string Name;
        public string Kind { get; set; }
        public int ProcKind;
        public int StartLine;
        public int EndLine;
        public string Sha256;

        public string Token
        {
            get
            {
                string path = (Name == null ? "#" : "@") + Project;
                if (Module != null) path += "." + Module;
                if (Name != null) path += "." + Name;
                if (Kind == "Property") path += ":" + new[] { "", "Let", "Set", "Get" }[ProcKind];
                return path;
            }
        }

        public string Display { get { return Token + "  —  " + Kind; } }
        public override string ToString() { return Display; }
    }

    // VBIDE is read only on the VBE UI thread. Filtering happens on copied strings.
    internal sealed class VbeChatReferences
    {
        private readonly VbeSession session;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly Queue<string> projects = new Queue<string>();
        private readonly Queue<VbeChatReference> pending = new Queue<VbeChatReference>();
        private readonly List<VbeChatReference> entries = new List<VbeChatReference>();
        public event Action Changed;
        public string Error { get; private set; }
        public bool IsLoading { get { return projects.Count > 0 || pending.Count > 0; } }

        public VbeChatReferences(VbeSession session) { this.session = session; }
        public IList<VbeChatReference> Entries { get { return entries; } }

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

        public void Step()
        {
            if (projects.Count > 0)
            {
                string project = projects.Dequeue();
                try
                {
                    foreach (var moduleInfo in Read("list_modules", project, null))
                    {
                        var item = new VbeChatReference { Project = project,
                            Module = Field(moduleInfo, "Name"), Kind = "Module" };
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
                Response response = session.Execute(new Request { Command = "list_procedures",
                    Project = module.Project, Module = module.Module });
                if (!response.Ok) throw new InvalidOperationException(response.Error);
                var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
                module.Sha256 = Field(data, "Sha256");
                object raw;
                var procedures = data != null && data.TryGetValue("Procedures", out raw) ? raw as object[] : null;
                if (procedures != null)
                    foreach (var value in procedures)
                    {
                        var procedure = value as IDictionary<string, object>;
                        if (procedure == null) continue;
                        int kind = Convert.ToInt32(procedure["Kind"]);
                        string declaration = Field(procedure, "Declaration");
                        entries.Add(new VbeChatReference { Project = module.Project, Module = module.Module,
                            Name = Field(procedure, "Name"), Kind = kind == 0
                                ? (declaration.IndexOf("Function", StringComparison.OrdinalIgnoreCase) >= 0 ? "Function" : "Sub")
                                : "Property",
                            ProcKind = kind, StartLine = Convert.ToInt32(procedure["StartLine"]),
                            EndLine = Convert.ToInt32(procedure["EndLine"]), Sha256 = module.Sha256 });
                    }
            }
            catch (Exception ex) { Error = module.Token + " : " + ex.Message; }
            Changed?.Invoke();
        }

        public IEnumerable<VbeChatReference> Match(string query)
        {
            return MatchPrefix(query, '\0');
        }

        public IEnumerable<VbeChatReference> MatchPrefix(string query, char prefix)
        {
            string term = query.TrimStart('#', '@');
            return entries.Where(item => (prefix == '\0' || (prefix == '@' ? item.Name != null : item.Name == null)) &&
                (item.Token.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                (item.Name != null && item.Name.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)))
                .OrderBy(item => item.Token.Substring(1).StartsWith(term, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(item => item.Token, StringComparer.OrdinalIgnoreCase).Take(40);
        }

        public Response Navigate(VbeChatReference item)
        {
            if (item.Module == null)
                return Response.Failure("Sélectionnez un module ou une procédure de ce projet pour ouvrir le code.");
            Response response = session.Execute(new Request { Command = "read_module", Project = item.Project, Module = item.Module });
            if (!response.Ok) return response;
            var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
            return session.Execute(new Request {
                Command = item.Name == null ? "select_code" : "select_procedure",
                Project = item.Project, Module = item.Module, Procedure = item.Name,
                ProcKind = item.ProcKind, StartLine = 1, ExpectedSha256 = Field(data, "Sha256")
            });
        }

        public string Resolve(VbeChatReference item)
        {
            if (item.Module == null)
            {
                var modules = Read("list_modules", item.Project, null);
                return item.Token + " (projet vivant ; " + modules.Length + " composants)\n" +
                    string.Join("\n", modules.Select(value => Field(value, "Name") +
                        " (type " + Field(value, "Type") + ", " + Field(value, "Lines") + " lignes)"));
            }
            Response response = session.Execute(new Request { Command = "read_module",
                Project = item.Project, Module = item.Module });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
            string code = Field(data, "Code");
            string sha = Field(data, "Sha256");
            if (item.Name == null) { item.Sha256 = sha; return item.Token + " (SHA-256 " + sha + ")\n" + code; }
            if (!string.Equals(sha, item.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(item.Token + " a changé depuis sa sélection. Supprimez le jeton et sélectionnez-le à nouveau avec #.");
            string[] lines = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            int start = Math.Max(1, item.StartLine);
            int end = Math.Min(lines.Length, item.EndLine);
            if (end < start) throw new InvalidOperationException(item.Token + " n'a plus de plage valide.");
            return item.Token + " (lignes " + start + "-" + end + ", SHA-256 " + sha + ")\n" +
                string.Join("\n", lines.Skip(start - 1).Take(end - start + 1));
        }

        private IDictionary<string, object>[] Read(string command, string project, string module)
        {
            Response response = session.Execute(new Request { Command = command, Project = project, Module = module });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var array = json.DeserializeObject(json.Serialize(response.Data)) as object[];
            return array == null ? new IDictionary<string, object>[0] :
                array.OfType<IDictionary<string, object>>().ToArray();
        }

        private static string Field(IDictionary<string, object> value, string name)
        {
            object raw;
            return value != null && value.TryGetValue(name, out raw) ? Convert.ToString(raw) : "";
        }
    }
}
