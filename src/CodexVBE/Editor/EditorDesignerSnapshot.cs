using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    // Compare stable Designer content, not OLE timestamps or uninitialized FRX padding.
    internal static class EditorDesignerSnapshot
    {
        internal static string Capture(object tree)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var root = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(tree));
            var rows = new SortedDictionary<string, object>(StringComparer.Ordinal);
            Properties(root, "UserForm", rows);
            Nodes(root, "Controls", rows);
            return json.Serialize(rows);
        }
        private static void Nodes(IDictionary<string, object> owner, string key, SortedDictionary<string, object> rows)
        {
            if (!owner.TryGetValue(key, out var raw) || !(raw is object[] nodes)) return;
            foreach (IDictionary<string, object> node in nodes)
            {
                string path = Convert.ToString(node["Path"]), type = Convert.ToString(node["Type"]);
                if (!new[] { "CheckBox", "ComboBox", "CommandButton", "Frame", "Image", "Label", "ListBox", "MultiPage", "OptionButton", "ScrollBar", "SpinButton", "TabStrip", "TextBox", "ToggleButton", "Page", "Tab" }.Contains(type))
                    throw new InvalidOperationException("UserForm replacement cannot verify this Designer control type: " + type);
                rows[path + "/$type"] = type;
                Properties(node, path, rows); Nodes(node, "Children", rows);
            }
        }
        private static void Properties(IDictionary<string, object> owner, string path, SortedDictionary<string, object> rows)
        {
            if (!owner.TryGetValue("Properties", out var raw) || !(raw is object[] properties)) throw new InvalidOperationException("Designer properties are unavailable.");
            foreach (IDictionary<string, object> property in properties)
            {
                string name = Convert.ToString(property["Name"]);
                // Runtime links, native Undo state and collection wrappers are not persisted content.
                if (new[] { "CanUndo", "CanRedo", "CanPaste", "ActiveControl", "Selected", "Controls", "Pages", "Tabs", "Parent", "Object", "_Font_Reserved" }.Contains(name)) continue;
                property.TryGetValue("Value", out var value); property.TryGetValue("Digest", out var digest);
                property.TryGetValue("Members", out var members); property.TryGetValue("Error", out var error);
                property.TryGetValue("Kind", out var kind); property.TryGetValue("Display", out var display);
                string failure = Convert.ToString(error);
                // MSForms reports E_UNEXPECTED when these optional images are absent.
                if (!string.IsNullOrEmpty(failure) && !((name == "Picture" || name == "MouseIcon") && failure.IndexOf("0x8000FFFF", StringComparison.OrdinalIgnoreCase) >= 0))
                    throw new InvalidOperationException("UserForm replacement cannot verify Designer property: " + path + "." + name);
                if (Convert.ToString(kind) == "object" && digest == null && members == null && string.IsNullOrEmpty(failure) && Convert.ToString(display) != "(empty)")
                    throw new InvalidOperationException("UserForm replacement cannot verify an opaque Designer property: " + path + "." + name);
                if (path == "UserForm" && name == "Name") value = "$form";
                rows[path + "/" + name] = new { Value = value, Digest = digest, Members = members, Error = failure };
            }
        }
    }
}
