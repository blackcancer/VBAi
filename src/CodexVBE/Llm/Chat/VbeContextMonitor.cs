using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // Incremental STA polling complements native reference events, including code changes
    // without a documented notification. One module is read per tick to bound UI work.
    internal sealed class VbeContextMonitor
    {
        private readonly Func<Request, Response> execute;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
        private readonly Queue<Request> pending = new Queue<Request>();
        private readonly Dictionary<string, string> fingerprints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private string scope;
        internal event Action Changed;
        internal VbeContextMonitor(Func<Request, Response> execute) { this.execute = execute; }
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
        private void Observe(string key, Response response)
        { if (response.Ok) ObserveValue(key, json.Serialize(response.Data)); }
        private void ObserveValue(string key, string value)
        {
            string previous;
            bool changed = fingerprints.TryGetValue(key, out previous) && previous != value;
            fingerprints[key] = value;
            if (changed) Changed?.Invoke();
        }
    }
}
