using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Implements project component operations, including post-save publication verification.</summary>
    internal sealed partial class VbeProjectComponents
    {
        // Raw revisions are deliberately unchanged. Only this verified-save content oracle
        // permits the documented edited-since-save flag to settle from false to true.
        /// <summary>Checks whether a verified save changed only the allowed Saved flag and volatile version fields.</summary>
        /// <param name="before">Immutable component snapshot captured before the save.</param>
        /// <param name="after">Component snapshot read back after the save.</param>
        /// <returns><see langword="true"/> when all stable descriptors match and Saved changed only from false to true.</returns>
        internal static bool PublicationPostSaveMetadataEquals(PublicationComponent before, PublicationComponent after)
        {
            if (before == null || after == null || before.ComponentSnapshotJson == null || after.ComponentSnapshotJson == null)
                return false;
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var left = serializer.Deserialize<Dictionary<string, object>>(before.ComponentSnapshotJson);
            var right = serializer.Deserialize<Dictionary<string, object>>(after.ComponentSnapshotJson);
            RequirePublicationDesignerReadable(left); RequirePublicationDesignerReadable(right);
            if (!left.TryGetValue("Properties", out var leftProperties) || !right.TryGetValue("Properties", out var rightProperties)) return false;
            var l = PublicationPostSaveDescriptors(leftProperties);
            var r = PublicationPostSaveDescriptors(rightProperties);
            if (l == null || r == null || !l.Keys.SequenceEqual(r.Keys)) return false;
            if (l.TryGetValue("Saved", out var oldSaved))
            {
                var newSaved = r["Saved"];
                if (!oldSaved.TryGetValue("Kind", out var oldKind) || !Equals(oldKind, "scalar") ||
                    !newSaved.TryGetValue("Kind", out var newKind) || !Equals(newKind, "scalar") ||
                    !oldSaved.TryGetValue("Value", out var oldValue) || !(oldValue is bool v) ||
                    !newSaved.TryGetValue("Value", out var newValue) || !(newValue is bool v1) ||
                    v && !v1) return false;
                // The complete descriptor, including Type/ReadOnly/Errors, remains equal.
                newSaved["Value"] = oldValue;
            }
            if (serializer.Serialize(l) != serializer.Serialize(r)) return false;
            left.Remove("Version"); right.Remove("Version");
            if (before.Type == 3 && after.Type == 3)
            {
                // Persisted designer/header/attributes are independently verified by the caller.
                // TreeVersion remains available in the immutable observation and source guards.
                left.Remove("FormVersion"); right.Remove("FormVersion");
            }
            left["Properties"] = l; right["Properties"] = r;
            return serializer.Serialize(left) == serializer.Serialize(right);
        }

        /// <summary>Indexes property descriptors by unique, nonblank name for order-independent comparison.</summary>
        /// <param name="descriptors">JSON-decoded property list; each item must be a dictionary with a string Name.</param>
        /// <returns>Ordinally sorted name-to-descriptor map, or <see langword="null"/> for malformed or duplicate entries.</returns>
        private static SortedDictionary<string, Dictionary<string, object>> PublicationPostSaveDescriptors(object descriptors)
        {
            if (!(descriptors is IList array)) return null;
            var result = new SortedDictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);
            foreach (object item in array)
            {
                if (!(item is Dictionary<string, object> row) || !row.TryGetValue("Name", out var name) || !(name is string v) ||
                    string.IsNullOrWhiteSpace(v) || result.ContainsKey(v)) return null;
                result.Add((string)name, row);
            }
            return result;
        }
    }
}
