using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class VbaGitComponent
    {
        public string Name { get; set; }
        public int Type { get; set; }
        public bool HasResources { get; set; }
        public string FileName { get { return Name + Extension(Type); } }
        internal static string Extension(int type)
        {
            switch (type) { case 1: return ".bas"; case 2: return ".cls"; case 3: return ".frm"; case 100: return ".vba"; }
            throw new InvalidOperationException("Type de composant VBA non pris en charge : " + type);
        }
    }

    internal sealed class VbaGitManifest
    {
        public int Format { get; set; } = 1;
        public string References { get; set; }
        public VbaGitComponent[] Components { get; set; }
    }

    // A snapshot is independent of the host file format and contains no local paths.
    internal sealed class VbaGitSnapshot
    {
        internal static readonly Encoding Utf8 = new UTF8Encoding(false, true);
        internal readonly VbaGitManifest Manifest;
        internal readonly SortedDictionary<string, byte[]> Files;
        internal const int MaxBytes = 32 * 1024 * 1024;

        internal VbaGitSnapshot(VbaGitManifest manifest, IDictionary<string, byte[]> files)
        {
            Manifest = manifest;
            Files = new SortedDictionary<string, byte[]>(files, StringComparer.Ordinal);
            foreach (string file in Files.Keys.ToArray())
                if (!file.EndsWith(".frx", StringComparison.Ordinal) && Files[file] != null)
                    Files[file] = Utf8.GetBytes(Utf8.GetString(Files[file]).Replace("\r\n", "\n").Replace("\r", "\n"));
            Validate();
        }

        internal static VbaGitSnapshot Read(IDictionary<string, byte[]> files)
        {
            if (!files.ContainsKey("manifest.json")) throw new InvalidOperationException(UiText.Get("No CodexVBA manifest in this repository."));
            var manifest = new JavaScriptSerializer().Deserialize<VbaGitManifest>(Utf8.GetString(files["manifest.json"]));
            var content = new Dictionary<string, byte[]>(files);
            content.Remove("manifest.json");
            return new VbaGitSnapshot(manifest, content);
        }

        internal SortedDictionary<string, byte[]> Serialize()
        {
            var result = new SortedDictionary<string, byte[]>(Files, StringComparer.Ordinal);
            result.Add("manifest.json", Utf8.GetBytes(new JavaScriptSerializer().Serialize(Manifest) + "\n"));
            return result;
        }

        internal bool SameAs(VbaGitSnapshot other)
        {
            if (other == null) return false;
            var left = Serialize(); var right = other.Serialize();
            return left.Count == right.Count && left.All(x => right.ContainsKey(x.Key) && x.Value.SequenceEqual(right[x.Key]));
        }

        internal string[] Changes(VbaGitSnapshot previous)
        {
            var before = previous?.Serialize() ?? new SortedDictionary<string, byte[]>();
            var after = Serialize();
            return before.Keys.Union(after.Keys).OrderBy(x => x, StringComparer.Ordinal).Where(x =>
                !before.ContainsKey(x) || !after.ContainsKey(x) || !before[x].SequenceEqual(after[x]))
                .Select(x => (!before.ContainsKey(x) ? "+ " : !after.ContainsKey(x) ? "− " : "~ ") + x).ToArray();
        }

        // Components, including their form resources, are the smallest commit/import unit.
        internal static VbaGitSnapshot Select(VbaGitSnapshot baseline, VbaGitSnapshot source, IEnumerable<string> names, bool references = false)
        {
            if (source == null) throw new ArgumentException(UiText.Get("The target contains no VBA sources."));
            var selected = new HashSet<string>(names ?? new string[0], StringComparer.Ordinal);
            var known = (baseline?.Manifest.Components ?? new VbaGitComponent[0]).Concat(source.Manifest.Components).Select(x => x.Name);
            if (selected.Any(x => !known.Contains(x))) throw new ArgumentException(UiText.Get("Unknown VBA module."));
            var components = (baseline?.Manifest.Components ?? new VbaGitComponent[0]).Where(x => !selected.Contains(x.Name))
                .Concat(source.Manifest.Components.Where(x => selected.Contains(x.Name))).OrderBy(x => x.Name, StringComparer.Ordinal).ToArray();
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            foreach (var component in components)
            {
                var origin = selected.Contains(component.Name) ? source : baseline;
                files.Add(component.FileName, origin.Files[component.FileName]);
                if (component.HasResources) files.Add(component.Name + ".frx", origin.Files[component.Name + ".frx"]);
            }
            return new VbaGitSnapshot(new VbaGitManifest { Components = components,
                References = references || baseline == null ? source.Manifest.References : baseline.Manifest.References }, files);
        }

        internal string ImportSummary(VbaGitSnapshot previous)
        {
            return string.Join(Environment.NewLine, Changes(previous)) + Environment.NewLine +
                UiText.Get("Required VBA references") + ": " + Manifest.References + Environment.NewLine +
                (previous != null && previous.Manifest.References != Manifest.References ? UiText.Get("References differ: align them in the VBE before importing.") : UiText.Get("A checkpoint protects this import."));
        }

        internal static void ValidateName(string name)
        {
            if (!Regex.IsMatch(name ?? "", @"^[\p{L}][\p{L}\p{N}_]{0,39}$") ||
                Regex.IsMatch(name, @"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$", RegexOptions.IgnoreCase))
                throw new InvalidOperationException("Nom de composant non exportable : " + name);
        }

        private void Validate()
        {
            if (Manifest == null || Manifest.Format != 1 || Manifest.Components == null || Manifest.Components.Length > 1024 || Manifest.References == null)
                throw new InvalidOperationException("Manifeste VBA invalide ou version non prise en charge.");
            var expected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var component in Manifest.Components)
            {
                if (component == null) throw new InvalidOperationException(UiText.Get("Component missing from the manifest."));
                ValidateName(component.Name);
                if (!expected.Add(component.FileName)) throw new InvalidOperationException(UiText.Get("Duplicate component name."));
                if (component.HasResources && (component.Type != 3 || !expected.Add(component.Name + ".frx")))
                    throw new InvalidOperationException("Ressources de formulaire invalides.");
            }
            if (Files.Count != expected.Count || Files.Any(x => !expected.Contains(x.Key) || x.Value == null) ||
                Files.Sum(x => (long)x.Value.Length) > MaxBytes)
                throw new InvalidOperationException(UiText.Get("VBA sources are incomplete, unexpected or too large (32 MB maximum)."));
            foreach (var component in Manifest.Components)
            {
                if (!Files.ContainsKey(component.FileName)) throw new InvalidOperationException(UiText.Get("Incorrect filename casing."));
                string text = Utf8.GetString(Files[component.FileName]);
                if (text.IndexOf('\0') >= 0 || text.Contains("<<<<<<< ") || text.Contains(">>>>>>> "))
                    throw new InvalidOperationException(UiText.Get("Binary source or unresolved conflict: ") + component.FileName);
                if (component.Type != 100 && !Regex.IsMatch(text, "^Attribute VB_Name = \"" + Regex.Escape(component.Name) + "\"\\r?$", RegexOptions.Multiline))
                    throw new InvalidOperationException(UiText.Get("Exported identity does not match the manifest: ") + component.Name);
                if (component.Type == 100 && Regex.IsMatch(text, @"^\s*Attribute\s", RegexOptions.Multiline | RegexOptions.IgnoreCase))
                    throw new InvalidOperationException(UiText.Get(".vba files contain only the visible code of the host module."));
                if (component.Type == 3)
                {
                    // A form must never address a companion file outside its snapshot.
                    foreach (Match match in Regex.Matches(text, "\"([^\"\\r\\n]+\\.frx)\"", RegexOptions.IgnoreCase))
                        if (!component.HasResources || !string.Equals(match.Groups[1].Value, component.Name + ".frx", StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Chemin FRX externe interdit : " + component.Name);
                    foreach (Match match in Regex.Matches(text, "=\\s*\"([^\"\\r\\n]+)\"\\s*:[0-9A-Fa-f]+"))
                        if (!component.HasResources || match.Groups[1].Value != component.Name + ".frx")
                            throw new InvalidOperationException("Ressource externe interdite : " + component.Name);
                }
            }
        }
    }
}
