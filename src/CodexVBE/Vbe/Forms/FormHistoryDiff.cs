using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal static class FormHistoryDiff
    {
        internal sealed class Change
        {
            public string Path { get; set; }
            public string Property { get; set; }
            public object Before { get; set; }
            public object After { get; set; }
        }
        private sealed class PropertyValue
        {
            public object Value { get; set; }
            public object Digest { get; set; }
            public object Members { get; set; }
            public string Error { get; set; }
        }
        private static bool Readable(object value) { return !(value is PropertyValue p) || string.IsNullOrEmpty(p.Error); }
        internal static int ReadErrorCount(object tree)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            return Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(tree))).Values.Count(x => !Readable(x));
        }
        internal static Change[] Compare(object before, object after)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var left = Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(before)));
            var right = Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(after)));
            return left.Keys.Union(right.Keys).OrderBy(x => x, StringComparer.Ordinal)
                .Where(k => (!left.ContainsKey(k) || Readable(left[k])) && (!right.ContainsKey(k) || Readable(right[k])))
                .Where(k => !left.ContainsKey(k) || !right.ContainsKey(k) || json.Serialize(left[k]) != json.Serialize(right[k]))
                .Select(k => new Change { Path = k.Substring(0, k.IndexOf('	')), Property = k.Substring(k.IndexOf('	') + 1),
                    Before = left.ContainsKey(k) ? left[k] : null, After = right.ContainsKey(k) ? right[k] : null }).ToArray();
        }
        private static Dictionary<string, object> Flatten(IDictionary<string, object> tree)
        {
            var result = new Dictionary<string, object>(StringComparer.Ordinal);
            ReadProperties(tree, "UserForm", result);
            ReadNodes(tree, "Controls", result);
            return result;
        }
        private static void ReadNodes(IDictionary<string, object> owner, string key, Dictionary<string, object> result)
        {
            if (!owner.TryGetValue(key, out object raw) || !(raw is object[] nodes)) return;
            foreach (IDictionary<string, object> node in nodes)
            {
                string path = Convert.ToString(node["Path"]);
                result[path + "\t$exists"] = true;
                if (node.TryGetValue("Type", out object type)) result[path + "\t$type"] = type;
                ReadProperties(node, path, result);
                ReadNodes(node, "Children", result);
            }
        }
        private static void ReadProperties(IDictionary<string, object> owner, string path, Dictionary<string, object> result)
        {
            if (!owner.TryGetValue("Properties", out object raw) || !(raw is object[] properties)) return;
            foreach (IDictionary<string, object> property in properties)
            {
                string name = Convert.ToString(property["Name"]);
                // These are history/clipboard availability, not Designer edits.
                if (name == "CanUndo" || name == "CanRedo" || name == "CanPaste") continue;
                property.TryGetValue("Value", out object value);
                property.TryGetValue("Digest", out object digest);
                property.TryGetValue("Members", out object members);
                property.TryGetValue("Error", out object error);
                result[path + "\t" + name] = new PropertyValue { Value = value, Digest = digest, Members = members, Error = Convert.ToString(error) };
            }
        }
    }
}
