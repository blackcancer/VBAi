using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{
    // Compare stable Designer content, not OLE timestamps or uninitialized FRX padding.
    /// <summary>Crée une empreinte sérialisée des propriétés stables du concepteur UserForm et de ses contrôles.</summary>
    internal static class EditorDesignerSnapshot
    {
        /// <summary>Sérialise les propriétés de formulaire et de contrôles dans un ordre déterministe.</summary>
        /// <param name="tree">Arbre de Designer produit par le lecteur des formulaires.</param>
        /// <returns>JSON des propriétés pertinentes, triées par chemin et nom.</returns>
        internal static string Capture(object tree)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var root = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(tree));
            var rows = new SortedDictionary<string, object>(StringComparer.Ordinal);
            Properties(root, "UserForm", rows);
            Nodes(root, "Controls", rows);
            return json.Serialize(rows);
        }
        /// <summary>Parcourt récursivement les contrôles connus et ajoute leurs propriétés à l’empreinte.</summary>
        /// <param name="owner">Nœud contenant éventuellement une collection d’enfants.</param>
        /// <param name="key">Nom de la propriété qui contient cette collection.</param>
        /// <param name="rows">Dictionnaire trié recevant les valeurs observées.</param>
        /// <exception cref="InvalidOperationException">Le formulaire contient un type de contrôle non vérifiable.</exception>
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
        /// <summary>Ajoute les propriétés persistées d’un nœud et refuse les valeurs opaques qui ne peuvent pas être vérifiées.</summary>
        /// <param name="owner">Nœud contenant la collection de propriétés.</param>
        /// <param name="path">Chemin stable du formulaire ou du contrôle.</param>
        /// <param name="rows">Dictionnaire trié qui reçoit les propriétés.</param>
        /// <exception cref="InvalidOperationException">Les propriétés sont absentes ou une valeur de Designer est invérifiable.</exception>
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
