using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Composant VBA inclus dans un snapshot Git.</summary>
    internal sealed class VbaGitComponent
    {

        /// <summary>Nom du composant dans le projet VBA.</summary>
        /// <value>Nom du composant dans le projet VBA.</value>
        public string Name { get; set; }

        /// <summary>Type de composant selon les constantes VBE.</summary>
        /// <value>Type de composant selon les constantes VBE.</value>
        public int Type { get; set; }

        /// <summary>Indique si le UserForm possède un fichier de ressources FRX.</summary>
        /// <value>Indique si le UserForm possède un fichier de ressources FRX.</value>
        public bool HasResources { get; set; }

        /// <summary>Nom de fichier produit pour ce type de composant.</summary>
        /// <value>Nom de fichier produit pour ce type de composant.</value>
        public string FileName { get { return Name + Extension(Type); } }

        /// <summary>Retourne l’extension correspondant au type de composant VBA.</summary>
        /// <param name="type">Type du composant d’après les constantes VBE.</param>
        /// <returns>Extension de fichier associée au type, précédée d’un point.</returns>
        internal static string Extension(int type)
        {
            switch (type) { case 1: return ".bas"; case 2: return ".cls"; case 3: return ".frm"; case 100: return ".vba"; }
            throw new InvalidOperationException("Type de composant VBA non pris en charge : " + type);
        }
    }

    /// <summary>Manifeste versionné du contenu VBA du dépôt.</summary>
    internal sealed class VbaGitManifest
    {

        /// <summary>Version du format de manifeste.</summary>
        /// <value>Version du format de manifeste.</value>
        public int Format { get; set; } = 1;

        /// <summary>Empreinte des références requises par le projet.</summary>
        /// <value>Empreinte des références requises par le projet.</value>
        public string References { get; set; }

        /// <summary>Composants et ressources décrits dans le snapshot.</summary>
        /// <value>Composants et ressources décrits dans le snapshot.</value>
        public VbaGitComponent[] Components { get; set; }
    }

    // A snapshot is independent of the host file format and contains no local paths.
    /// <summary>Normalise les fins de ligne et valide le manifeste ainsi que les fichiers.</summary>
    internal sealed class VbaGitSnapshot
    {

        /// <summary>Encodage UTF-8 strict, sans marqueur BOM.</summary>
        internal static readonly Encoding Utf8 = new UTF8Encoding(false, true);

        /// <summary>Manifeste des composants VBA et des références.</summary>
        internal readonly VbaGitManifest Manifest;

        /// <summary>Fichiers du snapshot, rangés par nom ordinal.</summary>
        internal readonly SortedDictionary<string, byte[]> Files;

        /// <summary>Taille maximale cumulée des fichiers du snapshot, en octets.</summary>
        internal const int MaxBytes = 32 * 1024 * 1024;

        /// <summary>Construit un snapshot, normalise ses sources textuelles et valide son contenu.</summary>
        /// <param name="manifest">Manifeste des composants à valider.</param>
        /// <param name="files">Fichiers nommés inclus dans le snapshot.</param>
        internal VbaGitSnapshot(VbaGitManifest manifest, IDictionary<string, byte[]> files)
        {
            Manifest = manifest;
            Files = new SortedDictionary<string, byte[]>(files, StringComparer.Ordinal);
            foreach (string file in Files.Keys.ToArray())
                if (!file.EndsWith(".frx", StringComparison.Ordinal) && Files[file] != null)
                    Files[file] = Utf8.GetBytes(Utf8.GetString(Files[file]).Replace("\r\n", "\n").Replace("\r", "\n"));
            Validate();
        }

        /// <summary>Charge le manifeste puis construit un snapshot validé à partir des autres fichiers.</summary>
        /// <param name="files">Fichiers nommés inclus dans le snapshot.</param>
        /// <returns>Snapshot créé après validation des fichiers et du manifeste.</returns>
        internal static VbaGitSnapshot Read(IDictionary<string, byte[]> files)
        {
            if (!files.ContainsKey("manifest.json")) throw new InvalidOperationException(UiText.Get("No CodexVBA manifest in this repository."));
            var manifest = new JavaScriptSerializer().Deserialize<VbaGitManifest>(Utf8.GetString(files["manifest.json"]));
            var content = new Dictionary<string, byte[]>(files);
            content.Remove("manifest.json");
            return new VbaGitSnapshot(manifest, content);
        }

        /// <summary>Sérialise les composants et le manifeste dans un dictionnaire trié.</summary>
        /// <returns>Dictionnaire contenant les fichiers VBA et manifest.json.</returns>
        internal SortedDictionary<string, byte[]> Serialize()
        {
            var result = new SortedDictionary<string, byte[]>(Files, StringComparer.Ordinal);
            result.Add("manifest.json", Utf8.GetBytes(new JavaScriptSerializer().Serialize(Manifest) + "\n"));
            return result;
        }

        /// <summary>Returns comparison data without altering serialized/exported resources.</summary>
        /// <returns>sorted dictionary&lt;string, byte[]&gt; produced by the operation for comparison files on vba git snapshot.</returns>
        internal SortedDictionary<string, byte[]> ComparisonFiles()
        {
            var result = Serialize();
            foreach (var component in Manifest.Components.Where(x => x.Type == 3 && x.HasResources))
            {
                string text = Utf8.GetString(Files[component.FileName]);
                string metadata = text.Substring(0, Regex.Match(text, "^Attribute VB_Name = ", RegexOptions.Multiline).Index);
                var blobDeclarations = OleBlobs(metadata).Cast<Match>().ToArray();
                var blobs = blobDeclarations.Select(x =>
                    (int)uint.Parse(x.Groups[1].Value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture)).Distinct().OrderBy(x => x).ToArray();
                if (blobs.Length == 0) continue;
                byte[] resources = Files[component.Name + ".frx"];
                // Opaque resources have unknown extents. Their starting offsets cannot
                // prove non-overlap, so any declaration outside a recognized OLE blob
                // keeps the entire companion byte-exact, including multiline syntax.
                bool opaque = ResourceReferences(metadata).Cast<Match>().Any(reference =>
                    !blobDeclarations.Any(blob => reference.Index >= blob.Index &&
                        reference.Index + reference.Length <= blob.Index + blob.Length));
                bool overlap = blobs.Where((offset, index) => index > 0 && offset < blobs[index - 1] + 24 + BitConverter.ToUInt32(resources, blobs[index - 1] + 4)).Any();
                if (!opaque && !overlap) result[component.Name + ".frx"] = FormResourcePreflight.ComparisonBytes(resources, blobs);
            }
            return result;
        }

        /// <summary>Compares one component file under the same rules used by revision guards.</summary>
        /// <param name="other">vba git snapshot that supplies the other for this operation.</param>
        /// <param name="file">Text that supplies the file value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for same file on vba git snapshot.</returns>
        internal bool SameFile(VbaGitSnapshot other, string file)
        {
            return other != null && ComparisonFiles().TryGetValue(file, out var left) &&
                other.ComparisonFiles().TryGetValue(file, out var right) && left.SequenceEqual(right);
        }

        /// <summary>Handles ole blobs for vba git snapshot.</summary>
        /// <param name="metadata">Text that supplies the metadata value. Use the format required by the calling operation.</param>
        /// <returns>match collection produced by the operation for ole blobs on vba git snapshot.</returns>
        private static MatchCollection OleBlobs(string metadata)
        {
            return Regex.Matches(metadata, "^[ \\t]*OleObjectBlob[ \\t]*=[ \\t]*\"[^\"\\r\\n]+\"[ \\t]*:[ \\t]*([0-9a-f]+)[ \\t]*$",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
        }

        /// <summary>Prepares a bounded form-font plan only for one unambiguous supported native container.</summary>
        /// <param name="component">vba git component that supplies the component for this operation.</param>
        /// <returns>form font binding[] produced by the operation for form fonts on vba git snapshot.</returns>
        internal FormStreamPadding.FormFontBinding[] FormFonts(VbaGitComponent component)
        {
            if (component.Type != 3 || !component.HasResources) return null;
            string text = Utf8.GetString(Files[component.FileName]);
            string metadata = text.Substring(0, Regex.Match(text, "^Attribute VB_Name = ", RegexOptions.Multiline).Index);
            var declarations = OleBlobs(metadata).Cast<Match>().ToArray();
            if (declarations.Length != 1 || ResourceReferences(metadata).Cast<Match>().Any(reference =>
                reference.Index < declarations[0].Index || reference.Index + reference.Length > declarations[0].Index + declarations[0].Length)) return null;
            int offset = checked((int)uint.Parse(declarations[0].Groups[1].Value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture));
            return FormResourcePreflight.ReadFontBindings(Files[component.Name + ".frx"], offset);
        }

        /// <summary>Finds resource references with the same grammar for validation and comparison.</summary>
        /// <param name="metadata">Text that supplies the metadata value. Use the format required by the calling operation.</param>
        /// <returns>match collection produced by the operation for resource references on vba git snapshot.</returns>
        private static MatchCollection ResourceReferences(string metadata)
        {
            return Regex.Matches(metadata, "=\\s*\"([^\"\\r\\n]+)\"[ \\t]*:[ \\t]*([^\\r\\n]*)");
        }

        /// <summary>Compares exact text/manifest data and logical form resources without rewriting serialized files.</summary>
        /// <param name="other">Snapshot comparé à l’instance courante.</param>
        /// <returns>True when both snapshots describe the same sources and resource contents.</returns>
        internal bool SameAs(VbaGitSnapshot other)
        {
            if (other == null) return false;
            var left = ComparisonFiles(); var right = other.ComparisonFiles();
            return left.Count == right.Count && left.All(x => right.ContainsKey(x.Key) && x.Value.SequenceEqual(right[x.Key]));
        }

        /// <summary>Liste les fichiers ajoutés, supprimés ou modifiés par rapport au snapshot précédent.</summary>
        /// <param name="previous">État de référence utilisé pour calculer les modifications.</param>
        /// <returns>Libellés préfixés par +, − ou ~ pour chaque différence.</returns>
        internal string[] Changes(VbaGitSnapshot previous)
        {
            var before = previous?.ComparisonFiles() ?? new SortedDictionary<string, byte[]>();
            var after = ComparisonFiles();
            return before.Keys.Union(after.Keys).OrderBy(x => x, StringComparer.Ordinal).Where(x =>
                !before.ContainsKey(x) || !after.ContainsKey(x) || !before[x].SequenceEqual(after[x]))
                .Select(x => (!before.ContainsKey(x) ? "+ " : !after.ContainsKey(x) ? "− " : "~ ") + x).ToArray();
        }

        // Components, including their form resources, are the smallest commit/import unit.
        /// <summary>Construit un snapshot avec les seuls composants sélectionnés et leurs ressources.</summary>
        /// <param name="baseline">Snapshot local servant de base aux composants non sélectionnés.</param>
        /// <param name="source">Snapshot source fournissant les composants choisis.</param>
        /// <param name="names">Noms exacts des composants à inclure depuis la source.</param>
        /// <param name="references">Indique si les références doivent également venir de la source.</param>
        /// <returns>Snapshot contenant les composants choisis et références retenues.</returns>
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

        /// <summary>Produit le résumé des changements et des références avant import.</summary>
        /// <param name="previous">État de référence utilisé pour calculer les modifications.</param>
        /// <returns>Résumé textuel des différences et des références à aligner.</returns>
        internal string ImportSummary(VbaGitSnapshot previous)
        {
            return string.Join(Environment.NewLine, Changes(previous)) + Environment.NewLine +
                UiText.Get("Required VBA references") + ": " + Manifest.References + Environment.NewLine +
                (previous != null && previous.Manifest.References != Manifest.References ? UiText.Get("References differ: align them in the VBE before importing.") : UiText.Get("A checkpoint protects this import."));
        }

        /// <summary>Vérifie qu’un nom de composant peut être exporté comme fichier.</summary>
        /// <param name="name">Nom de composant à vérifier.</param>
        internal static void ValidateName(string name)
        {
            if (!Regex.IsMatch(name ?? "", @"^[\p{L}][\p{L}\p{N}_]{0,39}$") ||
                Regex.IsMatch(name, @"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$", RegexOptions.IgnoreCase))
                throw new InvalidOperationException("Nom de composant non exportable : " + name);
        }

        /// <summary>Vérifie l’intégrité du manifeste, des noms, tailles, encodages et contenus exportés.</summary>
        private void Validate()
        {
            if (Manifest == null || Manifest.Format != 1 || Manifest.Components == null || Manifest.Components.Length > 1024 || Manifest.References == null)
                throw new InvalidOperationException("Manifeste VBA invalide ou version non prise en charge.");
            var expected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var component in Manifest.Components)
            {
                if (component == null) throw new InvalidOperationException(UiText.Get("Component missing from the manifest."));
                ValidateName(component.Name);
                if (!names.Add(component.Name)) throw new InvalidOperationException(UiText.Get("Duplicate component name."));
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
                    byte[] resources = null;
                    if (component.HasResources && (!Files.TryGetValue(component.Name + ".frx", out resources) || resources.Length == 0))
                        throw new InvalidOperationException("Form resources are missing, empty or incorrectly cased: " + component.Name);
                    // Only designer metadata contains resource offsets. VBA code and
                    // comments following VB_Name are not resource declarations.
                    int metadataLength = Regex.Match(text, "^Attribute VB_Name = ", RegexOptions.Multiline).Index;
                    string metadata = text.Substring(0, metadataLength);
                    foreach (Match resource in ResourceReferences(metadata))
                    {
                        uint offset;
                        if (!component.HasResources || resource.Groups[1].Value != component.Name + ".frx" ||
                            !uint.TryParse(resource.Groups[2].Value.Trim(), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out offset) ||
                            offset >= resources.Length)
                            throw new InvalidOperationException("Invalid or out-of-range form resource offset: " + component.Name);
                    }
                    foreach (Match blob in OleBlobs(metadata))
                        FormResourcePreflight.ValidateOleObjectBlob(resources,
                            (int)uint.Parse(blob.Groups[1].Value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture));
                    // These bounds are necessary, not a complete MS-OFORMS parser.
                    // Preserve resource bytes and the existing import/recovery guards.
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
