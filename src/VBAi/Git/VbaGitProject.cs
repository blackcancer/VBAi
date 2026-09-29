using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VBAi
{
    // All methods run on the VBE UI thread. Never send COM objects to the Git worker.
    /// <summary>Associe un résolveur de projet et son chemin hôte attendu.</summary>
    internal sealed class VbaGitProject
    {
        /// <summary>Fournit le projet COM courant sans déplacer les appels hors du thread VBE.</summary>
        private readonly Func<object> resolve;
        /// <summary>Chemin absolu du document hôte auquel le cache Git est lié.</summary>
        private readonly string hostPath;
        /// <summary>Retourne la page de codes ANSI du système Windows.</summary>
        /// <returns>Identifiant numérique de la page de codes ANSI active.</returns>
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetACP();
        /// <summary>Encodage natif strict utilisé pour lire et écrire les exports COM.</summary>
        /// <value>Encodage natif strict utilisé pour lire et écrire les exports COM.</value>
        private static Encoding NativeEncoding { get { return Encoding.GetEncoding(
            (int)GetACP(), EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); } }

        /// <summary>Crée l’adaptateur associé au résolveur et au chemin hôte attendus.</summary>
        /// <param name="resolve">Fonction qui résout le projet COM au moment de l’opération.</param>
        /// <param name="hostPath">Chemin absolu attendu du document hôte.</param>
        internal VbaGitProject(Func<object> resolve, string hostPath) { this.resolve = resolve; this.hostPath = hostPath; }

        /// <summary>Ouvre un module VBA au point fourni.</summary>
        /// <param name="name">Nom de l’outil Git à invoquer.</param>
        /// <param name="line">Numéro de ligne initial, ramené au minimum à 1.</param>
        internal void OpenModule(string name, int line = 1)
        {
            VbaGitSnapshot.ValidateName(name);
            dynamic project = CheckedProject();
            dynamic pane = project.VBComponents.Item(name).CodeModule.CodePane;
            pane.Show();
            pane.SetSelection(Math.Max(1, line), 1, Math.Max(1, line), 1);
        }

        /// <summary>Résout le projet et vérifie son chemin, son déverrouillage et le mode conception.</summary>
        /// <returns>Projet COM correspondant au document toujours lié et modifiable.</returns>
        private object CheckedProject()
        {
            dynamic project = resolve();
            if (!string.Equals(Path.GetFullPath((string)project.FileName), hostPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(UiText.Get("The linked document changed. Reopen GitHub integration."));
            if ((int)project.Mode != 2 || (int)project.Protection != 0)
                throw new InvalidOperationException(UiText.Get("The VBA project must be unlocked and in design mode."));
            return project;
        }

        /// <summary>Exporte les composants et références en snapshot validé sans envoyer d’objets COM au worker Git.</summary>
        /// <returns>Snapshot validé des composants, ressources et références.</returns>
        internal VbaGitSnapshot Capture()
        {
            dynamic project = CheckedProject();
            var components = new List<VbaGitComponent>();
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using (var scratch = new Scratch())
            {
                foreach (dynamic item in project.VBComponents)
                {
                    var component = new VbaGitComponent { Name = (string)item.Name, Type = (int)item.Type };
                    // Validate before using a COM-supplied name as a path.
                    VbaGitSnapshot.ValidateName(component.Name);
                    string file = Path.Combine(scratch.Path, component.FileName);
                    if (component.Type == 100)
                        files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(Code(item.CodeModule)));
                    else
                    {
                        item.Export(file);
                        files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(Normalize(File.ReadAllText(file, NativeEncoding))));
                        string frx = Path.Combine(scratch.Path, component.Name + ".frx");
                        component.HasResources = component.Type == 3 && File.Exists(frx);
                        if (component.HasResources) files.Add(component.Name + ".frx", File.ReadAllBytes(frx));
                    }
                    components.Add(component);
                }
            }
            var references = new List<string>();
            foreach (dynamic reference in project.References)
            {
                if ((bool)reference.IsBroken) throw new InvalidOperationException(UiText.Get("Missing VBA reference: fix it before synchronizing."));
                references.Add(((string)reference.GUID).ToUpperInvariant() + ":" + (int)reference.Major + ":" + (int)reference.Minor);
            }
            return new VbaGitSnapshot(new VbaGitManifest {
                Components = components.OrderBy(x => x.Name, StringComparer.Ordinal).ToArray(),
                References = string.Join(";", references.OrderBy(x => x, StringComparer.Ordinal))
            }, files);
        }

        /// <summary>Applique le snapshot si l’état avant mutation correspond à expected.</summary>
        /// <param name="target">Snapshot à importer.</param>
        /// <param name="expected">Snapshot attendu avant mutation.</param>
        /// <param name="beforeMutation">Action facultative exécutée après validation et avant la première modification.</param>
        internal void Apply(VbaGitSnapshot target, VbaGitSnapshot expected, Action beforeMutation = null)
        {
            if (!Capture().SameAs(expected)) throw new InvalidOperationException(UiText.Get("VBA changed during synchronization. No import performed."));
            if (target.Manifest.References != expected.Manifest.References)
                throw new InvalidOperationException(UiText.Get("VBA references differ. Align them in Tools > References before importing."));
            var beforeDocs = expected.Manifest.Components.Where(x => x.Type == 100).Select(x => x.Name);
            var afterDocs = target.Manifest.Components.Where(x => x.Type == 100).Select(x => x.Name);
            if (!beforeDocs.SequenceEqual(afterDocs))
                throw new InvalidOperationException(UiText.Get("Document modules do not match. Sheets and host modules must already exist with the same names."));

            using (var scratch = new Scratch())
            {
                // Encode and materialize the entire import before touching the live project.
                foreach (var file in target.Files)
                    File.WriteAllBytes(Path.Combine(scratch.Path, file.Key), file.Key.EndsWith(".frx", StringComparison.Ordinal) ? file.Value :
                        NativeEncoding.GetBytes(VbaGitSnapshot.Utf8.GetString(file.Value).Replace("\n", "\r\n")));
                dynamic project = CheckedProject();
                beforeMutation?.Invoke();
                var changed = new HashSet<string>(StringComparer.Ordinal);
                foreach (var old in expected.Manifest.Components)
                {
                    var next = target.Manifest.Components.FirstOrDefault(x => x.Name == old.Name);
                    bool same = next != null && next.Type == old.Type && next.HasResources == old.HasResources &&
                        expected.Files[old.FileName].SequenceEqual(target.Files[next.FileName]) &&
                        (!old.HasResources || expected.Files[old.Name + ".frx"].SequenceEqual(target.Files[next.Name + ".frx"]));
                    if (same) continue;
                    changed.Add(old.Name);
                    if (old.Type != 100) project.VBComponents.Remove(project.VBComponents.Item(old.Name));
                }
                foreach (var next in target.Manifest.Components)
                {
                    if (!changed.Contains(next.Name) && expected.Manifest.Components.Any(x => x.Name == next.Name)) continue;
                    if (next.Type == 100)
                    {
                        dynamic module = project.VBComponents.Item(next.Name).CodeModule;
                        int count = (int)module.CountOfLines;
                        if (count > 0) module.DeleteLines(1, count);
                        string code = VbaGitSnapshot.Utf8.GetString(target.Files[next.FileName]);
                        if (code.Length > 0) module.InsertLines(1, code.Replace("\n", "\r\n"));
                    }
                    else
                    {
                        // A failing COM call can still have applied. Do not retry or automatically re-import.
                        dynamic imported = project.VBComponents.Import(Path.Combine(scratch.Path, next.FileName));
                        if ((string)imported.Name != next.Name || (int)imported.Type != next.Type)
                            throw new InvalidOperationException(UiText.Get("Unexpected identity after import: ") + next.Name + UiText.Get(". Use Restore."));
                    }
                }
            }
            if (!Capture().SameAs(target)) throw new InvalidOperationException(UiText.Get("The VBE did not preserve the imported sources exactly. Use Restore or check the project."));
        }

        /// <summary>Lit le texte visible d’un module et normalise ses fins de ligne.</summary>
        /// <param name="module">Module de code dont les lignes sont lues.</param>
        /// <returns>Code source normalisé avec LF.</returns>
        private static string Code(dynamic module)
        {
            int count = (int)module.CountOfLines;
            return count == 0 ? "" : Normalize((string)module.Lines[1, count]);
        }
        /// <summary>Remplace CRLF et CR par LF.</summary>
        /// <param name="text">Texte source dont les fins de ligne sont normalisées.</param>
        /// <returns>Texte dont les séparateurs de ligne sont des LF.</returns>
        private static string Normalize(string text) { return text.Replace("\r\n", "\n").Replace("\r", "\n"); }

        /// <summary>Répertoire temporaire privé utilisé pour un transfert de fichiers.</summary>
        private sealed class Scratch : IDisposable
        {
        /// <summary>Chemin unique du répertoire temporaire de l’opération.</summary>
            internal readonly string Path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VBAi", "GitTemporary", Guid.NewGuid().ToString("N"));
            /// <summary>Crée le répertoire temporaire unique.</summary>
            internal Scratch() { Directory.CreateDirectory(Path); }
        /// <summary>Supprime les fichiers générés dans le répertoire temporaire.</summary>
            public void Dispose()
            {
                // Only our freshly generated, private flat directory is cleaned up.
                try { foreach (string file in Directory.GetFiles(Path)) File.Delete(file); Directory.Delete(Path); }
                catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
    }
}
