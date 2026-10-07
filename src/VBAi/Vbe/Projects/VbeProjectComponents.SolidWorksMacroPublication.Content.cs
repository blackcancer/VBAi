using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Implements project component operations, including persisted publication verification.</summary>
    internal sealed partial class VbeProjectComponents
    {
        // A transport oracle, separate from source-local revision/TreeVersion guards.
        /// <summary>Compares exported UserForm headers, persisted root properties, and the complete control inventories.</summary>
        /// <param name="expected">Component metadata and Designer JSON from the approved source snapshot.</param>
        /// <param name="actual">Component metadata and Designer JSON read back from the published project.</param>
        /// <param name="expectedExport">Path to the source UserForm text export and its FRX companion.</param>
        /// <param name="actualExport">Path to the published UserForm text export and its FRX companion.</param>
        /// <returns><see langword="true"/> only when normalized headers, root settings, and all controls match.</returns>
        /// <exception cref="IOException">An export is oversized or cannot be decoded losslessly.</exception>
        /// <exception cref="InvalidOperationException">A required form header, FRX binding, or complete Designer inventory is invalid.</exception>
        internal static bool PublicationDesignerContentEquals(PublicationComponent expected,
            PublicationComponent actual, string expectedExport, string actualExport)
        {
            string left = PublicationDesignerHeader(expectedExport, expected.Name);
            string right = PublicationDesignerHeader(actualExport, actual.Name);
            if (!string.Equals(left, right, StringComparison.Ordinal)) return false;
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var a = serializer.Deserialize<Dictionary<string, object>>(expected.DesignerJson);
            var b = serializer.Deserialize<Dictionary<string, object>>(actual.DesignerJson);
            RequirePublicationDesignerReadable(a); RequirePublicationDesignerReadable(b);
            var x = PublicationDesignerRootProperties(a, left);
            var y = PublicationDesignerRootProperties(b, right);
            if (!x.Keys.SequenceEqual(y.Keys) || x.Any(p => p.Value != y[p.Key])) return false;
            if (!a.TryGetValue("Controls", out object ac) || !b.TryGetValue("Controls", out object bc) ||
                !(ac is IEnumerable) || ac is string || !(bc is IEnumerable) || bc is string)
                throw new InvalidOperationException("Complete persisted control inventories are required.");
            return serializer.Serialize(ac) == serializer.Serialize(bc);
        }

        /// <summary>Reads a bounded UserForm export and normalizes only its exact FRX filename binding.</summary>
        /// <param name="path">Absolute or relative path to the exported .frm file; its sibling .frx must exist.</param>
        /// <param name="name">Expected VB_Name used in the canonical MSForms export header.</param>
        /// <returns>Header text with the owned FRX filename replaced by a stable comparison token.</returns>
        /// <exception cref="IOException">The export exceeds 16 MiB or its system-code-page encoding is not lossless.</exception>
        /// <exception cref="InvalidOperationException">The canonical header or unique sibling FRX binding is missing or differs.</exception>
        private static string PublicationDesignerHeader(string path, string name)
        {
            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length > 16 * 1024 * 1024) throw new IOException("Form export size limit exceeded.");
            string text = System.Text.Encoding.Default.GetString(bytes);
            if (!bytes.SequenceEqual(System.Text.Encoding.Default.GetBytes(text))) throw new IOException("Lossless form export encoding is required.");
            var attribute = Regex.Match(text, @"(?m)^Attribute VB_Name\s*=");
            if (!attribute.Success) throw new InvalidOperationException("A complete UserForm export header is required.");
            string header = text.Substring(0, attribute.Index);
            if (!Regex.IsMatch(header, @"\AVERSION 5\.00\r\nBegin \{C62A69F0-16DC-11CE-9E98-00AA00574A4F\} " + Regex.Escape(name) + @"\s*\r\n", RegexOptions.CultureInvariant))
                throw new InvalidOperationException("The canonical MSForm export header differs.");
            var blobs = Regex.Matches(header, @"(?m)^(\s*OleObjectBlob\s*=\s*)""([^""\r\n]+)""(:[0-9A-Fa-f]+[^\r\n]*)(\r?)$");
            if (blobs.Count != 1) throw new InvalidOperationException("A unique owned form resource binding is required.");
            string file = blobs[0].Groups[2].Value;
            string expectedFile = Path.GetFileName(Path.ChangeExtension(path, ".frx"));
            if (!string.Equals(file, expectedFile, StringComparison.Ordinal) || Path.GetFileName(file) != file ||
                !File.Exists(Path.ChangeExtension(path, ".frx")))
                throw new InvalidOperationException("The exact exported FRX companion is unavailable.");
            return header.Substring(0, blobs[0].Groups[2].Index) + "$owned-frx" +
                header.Substring(blobs[0].Groups[2].Index + blobs[0].Groups[2].Length);
        }

        /// <summary>Compares persisted root Designer settings, using the export header only for documented defaults.</summary>
        /// <param name="tree">Deserialized Designer JSON whose root property getters may be incomplete.</param>
        /// <param name="header">Export header used to verify stored values when root getters are absent or normalized.</param>
        /// <returns>Sorted property names mapped to serialized, verified values for deterministic comparison.</returns>
        /// <exception cref="InvalidOperationException">A required property is unreadable or disagrees with its persisted setting.</exception>
        private static SortedDictionary<string, string> PublicationDesignerRootProperties(
            Dictionary<string, object> tree, string header)
        {
            if (!tree.TryGetValue("Properties", out object raw) || !(raw is IEnumerable) || raw is string)
                throw new InvalidOperationException("Complete persisted root properties are required.");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var defaults = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["HelpContextID"] = 0,
                ["ShowModal"] = true,
                ["WhatsThisButton"] = false,
                ["WhatsThisHelp"] = false,
                ["Visible"] = true
            };
            foreach (object item in (IEnumerable)raw)
            {
                if (!(item is IDictionary<string, object> property) || !property.TryGetValue("Name", out object propertyName) || !(propertyName is string v) ||
                    !seen.Add(v)) throw new InvalidOperationException("Root property identity is incomplete or duplicated.");
                string name = (string)propertyName;
                // These public API properties describe the live undo/clipboard state, never persisted form contents.
                if (name == "CanUndo" || name == "CanRedo" || name == "CanPaste") continue;
                if (defaults.ContainsKey(name))
                {
                    if (!property.TryGetValue("Value", out object value) || !property.TryGetValue("Kind", out object kind) ||
                        (string)kind != "scalar" || !property.TryGetValue("Error", out object error) || error != null ||
                        (name == "HelpContextID" ? !(value is int) : !(value is bool)))
                        throw new InvalidOperationException("A normative persisted root property is unreadable.");
                    object stored = PublicationRootHeaderValue(header, name, defaults[name]);
                    if (!Equals(value, stored)) throw new InvalidOperationException("A root getter differs from its persisted export setting.");
                    result[name] = serializer.Serialize(stored);
                }
                else result[name] = serializer.Serialize(property);
            }
            foreach (var pair in defaults)
                if (!result.ContainsKey(pair.Key))
                {
                    object stored = PublicationRootHeaderValue(header, pair.Key, pair.Value);
                    if (!Equals(stored, pair.Value)) throw new InvalidOperationException("A missing root getter cannot verify a nondefault persisted setting.");
                    result[pair.Key] = serializer.Serialize(stored);
                }
            return result;
        }

        /// <summary>Parses one persisted root setting, or returns its documented default when the export omits it.</summary>
        /// <param name="header">UserForm export header containing persisted root settings.</param>
        /// <param name="name">Exact setting name to find once in the header.</param>
        /// <param name="fallback">Documented default type and value; booleans and nonnegative integers are supported.</param>
        /// <returns>The parsed Boolean or integer value, or <paramref name="fallback"/> if no setting is present.</returns>
        /// <exception cref="InvalidOperationException">The setting is duplicated or has an unsupported persisted representation.</exception>
        private static object PublicationRootHeaderValue(string header, string name, object fallback)
        {
            var matches = Regex.Matches(header, @"(?m)^\s*" + Regex.Escape(name) + @"\s*=\s*([^\r\n]*)\r?$");
            if (matches.Count == 0) return fallback; // Only five MS-OVBA/VBA documented root defaults, in a proved export.
            if (matches.Count != 1) throw new InvalidOperationException("A persisted root setting is duplicated.");
            string token = matches[0].Groups[1].Value.Split('\'')[0].Trim();
            if (fallback is bool)
            {
                if (token == "-1" || token == "True") return true;
                if (token == "0" || token == "False") return false;
            }
            else
            {
                if (int.TryParse(token, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int value) && value >= 0) return value;
            }
            throw new InvalidOperationException("A persisted root setting has an unsupported encoding.");
        }

        /// <summary>Counts an extra prefix of CRLF pairs inserted before an otherwise identical imported source string.</summary>
        /// <param name="expected">Expected source text before the importer's added line-ending prefix.</param>
        /// <param name="actual">Observed text that may contain the expected text after the prefix.</param>
        /// <returns>Number of leading CRLF pairs, from 0 through 256; returns 0 for any other difference.</returns>
        internal static int PublicationSurplusImportedCrLfPrefix(string expected, string actual)
        {
            if (expected == null || actual == null || actual.Length <= expected.Length ||
                !actual.EndsWith(expected, StringComparison.Ordinal)) return 0;
            int length = actual.Length - expected.Length;
            if (length % 2 != 0 || length > 512) return 0;
            for (int i = 0; i < length; i += 2) if (actual[i] != '\r' || actual[i + 1] != '\n') return 0;
            return length / 2;
        }
    }
}
