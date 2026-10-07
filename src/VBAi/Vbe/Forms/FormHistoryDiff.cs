using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Aplatit deux instantanés Designer et calcule les différences dont les deux valeurs sont lisibles.</summary>
    internal static class FormHistoryDiff
    {

        /// <summary>Changement d’une propriété ou de l’existence d’un nœud entre deux arbres Designer.</summary>
        internal sealed class Change
        {

            /// <summary>Gets or sets the path.</summary>
            /// <value>Current path exposed by change.</value>
            public string Path { get; set; }

            /// <summary>Gets or sets the property.</summary>
            /// <value>Current property exposed by change.</value>
            public string Property { get; set; }

            /// <summary>Gets or sets the before.</summary>
            /// <value>Current before exposed by change.</value>
            public object Before { get; set; }

            /// <summary>Gets or sets the after.</summary>
            /// <value>Current after exposed by change.</value>
            public object After { get; set; }
        }

        /// <summary>Valeur sérialisée avec les données qui permettent de comparer une propriété Designer.</summary>
        private sealed class PropertyValue
        {

            /// <summary>Gets or sets the value.</summary>
            /// <value>Current value exposed by property value.</value>
            public object Value { get; set; }

            /// <summary>Gets or sets the digest.</summary>
            /// <value>Current digest exposed by property value.</value>
            public object Digest { get; set; }

            /// <summary>Gets or sets the members.</summary>
            /// <value>Current members exposed by property value.</value>
            public object Members { get; set; }

            /// <summary>Gets or sets the error.</summary>
            /// <value>Current error exposed by property value.</value>
            public string Error { get; set; }
        }

        /// <summary>Indique si la valeur peut participer à une comparaison.</summary>
        /// <param name="value">Valeur de propriété aplatie.</param>
        /// <returns><see langword="true"/> si elle n’est pas affectée par une erreur de lecture.</returns>
        private static bool Readable(object value) { return !(value is PropertyValue p) || string.IsNullOrEmpty(p.Error); }

        /// <summary>Compte les propriétés illisibles dans un instantané Designer.</summary>
        /// <param name="tree">Arbre sérialisable du formulaire et de ses contrôles.</param>
        /// <returns>Nombre de propriétés contenant une erreur de lecture.</returns>
        internal static int ReadErrorCount(object tree)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            return Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(tree))).Values.Count(x => !Readable(x));
        }

        /// <summary>Compare deux instantanés et retourne les différences structurelles ou de propriété vérifiables.</summary>
        /// <param name="before">Premier instantané.</param>
        /// <param name="after">Second instantané.</param>
        /// <returns>Changements triés par chemin et propriété; les propriétés illisibles sont exclues.</returns>
        internal static Change[] Compare(object before, object after)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var left = Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(before)));
            var right = Flatten((IDictionary<string, object>)json.DeserializeObject(json.Serialize(after)));
            return left.Keys.Union(right.Keys).OrderBy(x => x, StringComparer.Ordinal)
                .Where(k => (!left.ContainsKey(k) || Readable(left[k])) && (!right.ContainsKey(k) || Readable(right[k])))
                .Where(k => !left.ContainsKey(k) || !right.ContainsKey(k) || json.Serialize(left[k]) != json.Serialize(right[k]))
                .Select(k => new Change
                {
                    Path = k.Substring(0, k.IndexOf('	')),
                    Property = k.Substring(k.IndexOf('	') + 1),
                    Before = left.ContainsKey(k) ? left[k] : null,
                    After = right.ContainsKey(k) ? right[k] : null
                }).ToArray();
        }

        /// <summary>Convertit l’arbre JSON en clés chemin/propriété pour comparer ses instantanés.</summary>
        /// <param name="tree">Racine désérialisée du formulaire.</param>
        /// <returns>Valeurs indexées par chemin de contrôle et nom de propriété.</returns>
        private static Dictionary<string, object> Flatten(IDictionary<string, object> tree)
        {
            var result = new Dictionary<string, object>(StringComparer.Ordinal);
            ReadProperties(tree, "UserForm", result);
            ReadNodes(tree, "Controls", result);
            return result;
        }

        /// <summary>Parcourt les contrôles et leurs enfants en enregistrant présence, type et propriétés.</summary>
        /// <param name="owner">Nœud contenant éventuellement une collection enfant.</param>
        /// <param name="key">Nom de la collection à lire.</param>
        /// <param name="result">Index de différences en cours de construction.</param>
        private static void ReadNodes(IDictionary<string, object> owner, string key, Dictionary<string, object> result)
        {
            if (!owner.TryGetValue(key, out object raw) || !(raw is object[] nodes)) return;
            foreach (IDictionary<string, object> node in nodes.Cast<IDictionary<string, object>>())
            {
                string path = Convert.ToString(node["Path"]);
                result[path + "\t$exists"] = true;
                if (node.TryGetValue("Type", out object type)) result[path + "\t$type"] = type;
                ReadProperties(node, path, result);
                ReadNodes(node, "Children", result);
            }
        }

        /// <summary>Ajoute les propriétés sérialisées d’un nœud, en ignorant les indicateurs d’historique non persistés.</summary>
        /// <param name="owner">Nœud contenant les propriétés.</param>
        /// <param name="path">Chemin de formulaire ou de contrôle.</param>
        /// <param name="result">Index de différences enrichi.</param>
        private static void ReadProperties(IDictionary<string, object> owner, string path, Dictionary<string, object> result)
        {
            if (!owner.TryGetValue("Properties", out object raw) || !(raw is object[] properties)) return;
            foreach (IDictionary<string, object> property in properties.Cast<IDictionary<string, object>>())
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
