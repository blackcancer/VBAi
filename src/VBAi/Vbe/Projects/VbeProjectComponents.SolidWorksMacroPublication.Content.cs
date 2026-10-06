using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {
        // A transport oracle, separate from source-local revision/TreeVersion guards.
        /// <summary>Handles publication designer content equals for vbe project components.</summary>
        /// <param name="expected">publication component that supplies the expected for this operation.</param>
        /// <param name="actual">publication component that supplies the actual for this operation.</param>
        /// <param name="expectedExport">Text that supplies the expected export value. Use the format required by the calling operation.</param>
        /// <param name="actualExport">Text that supplies the actual export value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for publication designer content equals on vbe project components.</returns>
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
            object ac, bc;
            if (!a.TryGetValue("Controls", out ac) || !b.TryGetValue("Controls", out bc) ||
                !(ac is IEnumerable) || ac is string || !(bc is IEnumerable) || bc is string)
                throw new InvalidOperationException("Complete persisted control inventories are required.");
            return serializer.Serialize(ac) == serializer.Serialize(bc);
        }

        /// <summary>Handles publication designer header for vbe project components.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for publication designer header on vbe project components.</returns>
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

        /// <summary>Handles publication designer root properties for vbe project components.</summary>
        /// <param name="tree">dictionary&lt;string, object&gt; that supplies the tree for this operation.</param>
        /// <param name="header">Text that supplies the header value. Use the format required by the calling operation.</param>
        /// <returns>sorted dictionary&lt;string, string&gt; produced by the operation for publication designer root properties on vbe project components.</returns>
        private static SortedDictionary<string, string> PublicationDesignerRootProperties(
            Dictionary<string, object> tree, string header)
        {
            object raw;
            if (!tree.TryGetValue("Properties", out raw) || !(raw is IEnumerable) || raw is string)
                throw new InvalidOperationException("Complete persisted root properties are required.");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var defaults = new Dictionary<string, object>(StringComparer.Ordinal) {
                ["HelpContextID"] = 0, ["ShowModal"] = true, ["WhatsThisButton"] = false,
                ["WhatsThisHelp"] = false, ["Visible"] = true };
            foreach (object item in (IEnumerable)raw)
            {
                var property = item as IDictionary<string, object>; object propertyName;
                if (property == null || !property.TryGetValue("Name", out propertyName) || !(propertyName is string) ||
                    !seen.Add((string)propertyName)) throw new InvalidOperationException("Root property identity is incomplete or duplicated.");
                string name = (string)propertyName;
                // These public API properties describe the live undo/clipboard state, never persisted form contents.
                if (name == "CanUndo" || name == "CanRedo" || name == "CanPaste") continue;
                if (defaults.ContainsKey(name))
                {
                    object value, kind, error;
                    if (!property.TryGetValue("Value", out value) || !property.TryGetValue("Kind", out kind) ||
                        (string)kind != "scalar" || !property.TryGetValue("Error", out error) || error != null ||
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

        /// <summary>Handles publication root header value for vbe project components.</summary>
        /// <param name="header">Text that supplies the header value. Use the format required by the calling operation.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="fallback">object that supplies the fallback for this operation.</param>
        /// <returns>object produced by the operation for publication root header value on vbe project components.</returns>
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
                int value;
                if (int.TryParse(token, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out value) && value >= 0) return value;
            }
            throw new InvalidOperationException("A persisted root setting has an unsupported encoding.");
        }

        /// <summary>Handles publication surplus imported cr lf prefix for vbe project components.</summary>
        /// <param name="expected">Text that supplies the expected value. Use the format required by the calling operation.</param>
        /// <param name="actual">Text that supplies the actual value. Use the format required by the calling operation.</param>
        /// <returns>int produced by the operation for publication surplus imported cr lf prefix on vbe project components.</returns>
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