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
            if (!files.ContainsKey("manifest.json")) throw new InvalidOperationException("Aucun manifeste CodexVBA dans ce dépôt.");
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
                if (component == null) throw new InvalidOperationException("Composant absent du manifeste.");
                ValidateName(component.Name);
                if (!expected.Add(component.FileName)) throw new InvalidOperationException("Nom de composant dupliqué.");
                if (component.HasResources && (component.Type != 3 || !expected.Add(component.Name + ".frx")))
                    throw new InvalidOperationException("Ressources de formulaire invalides.");
            }
            if (Files.Count != expected.Count || Files.Any(x => !expected.Contains(x.Key) || x.Value == null) ||
                Files.Sum(x => (long)x.Value.Length) > MaxBytes)
                throw new InvalidOperationException("Sources VBA incomplètes, inattendues ou trop volumineuses (32 Mo maximum).");
            foreach (var component in Manifest.Components)
            {
                if (!Files.ContainsKey(component.FileName)) throw new InvalidOperationException("Casse du nom de fichier incorrecte.");
                string text = Utf8.GetString(Files[component.FileName]);
                if (text.IndexOf('\0') >= 0 || text.Contains("<<<<<<< ") || text.Contains(">>>>>>> "))
                    throw new InvalidOperationException("Source binaire ou conflit non résolu : " + component.FileName);
                if (component.Type != 100 && !Regex.IsMatch(text, "^Attribute VB_Name = \"" + Regex.Escape(component.Name) + "\"\\r?$", RegexOptions.Multiline))
                    throw new InvalidOperationException("L’identité exportée ne correspond pas au manifeste : " + component.Name);
                if (component.Type == 100 && Regex.IsMatch(text, @"^\s*Attribute\s", RegexOptions.Multiline | RegexOptions.IgnoreCase))
                    throw new InvalidOperationException("Les fichiers .vba contiennent uniquement le code visible du module hôte.");
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
