using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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

        /// <summary>Project identity captured when this integration is opened.</summary>
        private readonly object boundProject;

        /// <summary>Reads the current persisted document path on the VBE thread.</summary>
        private readonly Func<object, string> readHostPath;

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
        /// <param name="readHostPath">Optional owning-thread path reader for isolated host contracts.</param>
        internal VbaGitProject(Func<object> resolve, string hostPath, Func<object, string> readHostPath = null)
        {
            this.resolve = resolve ?? throw new ArgumentNullException(nameof(resolve));
            if (string.IsNullOrWhiteSpace(hostPath) || !Path.IsPathRooted(hostPath))
                throw new InvalidOperationException("A saved absolute host document path is required for Git.");
            this.hostPath = Path.GetFullPath(hostPath);
            this.readHostPath = readHostPath ?? (project => VbeProjectHostPath.Read(project));
            boundProject = resolve();
        }

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
            if (!VbeProjectHostPath.SameProject(boundProject, (object)project) ||
                !string.Equals(readHostPath((object)project), hostPath, StringComparison.OrdinalIgnoreCase))
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

            FormFontObservation observation = null;
            object deferredImported = null;
            FormStreamPadding.FormFontBinding[] deferredBindings = null;
            Exception deferredPrimary = null;
            try
            {
                using (var scratch = new Scratch())
                {
                    // Encode and materialize the entire import before touching the live project.
                    var formFonts = target.Manifest.Components.Where(component => component.Type == 3)
                        .ToDictionary(component => component.Name, target.FormFonts, StringComparer.Ordinal);
                    foreach (var file in target.Files)
                        File.WriteAllBytes(Path.Combine(scratch.Path, file.Key), file.Key.EndsWith(".frx", StringComparison.Ordinal) ? file.Value :
                            NativeEncoding.GetBytes(VbaGitSnapshot.Utf8.GetString(file.Value).Replace("\n", "\r\n")));
                    dynamic project = CheckedProject();
                    var expectedFiles = expected.ComparisonFiles();
                    var targetFiles = target.ComparisonFiles();
                    var changed = new HashSet<string>(StringComparer.Ordinal);
                    foreach (var old in expected.Manifest.Components)
                    {
                        var next = target.Manifest.Components.FirstOrDefault(x => x.Name == old.Name);
                        bool same = next != null && next.Type == old.Type && next.HasResources == old.HasResources &&
                            expectedFiles[old.FileName].SequenceEqual(targetFiles[next.FileName]) &&
                            (!old.HasResources || expectedFiles[old.Name + ".frx"].SequenceEqual(targetFiles[next.Name + ".frx"]));
                        if (same) continue;
                        changed.Add(old.Name);
                    }
                    if (System.Runtime.InteropServices.Marshal.IsComObject((object)project) && formFonts.Any(pair =>
                        pair.Value != null && pair.Value.Length != 0 && (changed.Contains(pair.Key) || !expected.Manifest.Components.Any(old => old.Name == pair.Key))))
                        FormFontRestoration.RequireOwner((object)project);
                    observation = FormFontObservation.TryBegin(hostPath, target, changed, (object)project);
                    try
                    {
                        beforeMutation?.Invoke();
                        foreach (var old in expected.Manifest.Components.Where(old => changed.Contains(old.Name) && old.Type != 100))
                            project.VBComponents.Remove(project.VBComponents.Item(old.Name));
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
                                if (next.Type == 3 && observation != null && observation.IsAfterInitialCapture &&
                                    observation.FormName == next.Name && System.Runtime.InteropServices.Marshal.IsComObject((object)imported))
                                {
                                    deferredImported = (object)imported;
                                    deferredBindings = formFonts[next.Name];
                                }
                                if ((string)imported.Name != next.Name || (int)imported.Type != next.Type)
                                    throw new InvalidOperationException(UiText.Get("Unexpected identity after import: ") + next.Name + UiText.Get(". Use Restore."));
                                if (next.Type == 3)
                                {
                                    RestoreFormImportCode((object)imported.CodeModule, VbaGitSnapshot.Utf8.GetString(target.Files[next.FileName]), () => {
                                        dynamic current = CheckedProject();
                                        if (!VbeProjectHostPath.SameProject((object)current.VBComponents.Item(next.Name), (object)imported))
                                            throw new InvalidOperationException("The imported form identity changed before code readback.");
                                    });
                                    if (System.Runtime.InteropServices.Marshal.IsComObject((object)imported) &&
                                        (observation == null || !observation.IsAfterInitialCapture || observation.FormName != next.Name))
                                    {
                                        // Explicit diagnostic observations preserve their declared original
                                        // order. Ordinary imports retain fonts that already match and first
                                        // initialize the actual native designer when its resources differ.
                                        var bindings = formFonts[next.Name];
                                        if (observation == null || observation.FormName != next.Name)
                                            bindings = ImportedFormMaterialization.Prepare(target, next, bindings,
                                                () => CaptureImportedForm(next.Name, (object)imported),
                                                probe => ImportedFormMaterialization.Materialize(CheckedProject(), (object)imported, probe,
                                                    () => RequireImportedForm(next.Name, (object)imported)),
                                                () => RequireImportedForm(next.Name, (object)imported));
                                        FormFontRestoration.Restore((object)imported, bindings,
                                            () => RequireImportedForm(next.Name, (object)imported),
                                            observation != null && observation.FormName == next.Name ? observation : null);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception error)
                    {
                        try { observation?.Failure(error); }
                        catch (Exception receiptError) { throw new AggregateException("Root font observation failed while recording the native failure.", error, receiptError); }
                        throw;
                    }
                }
                if (observation != null && observation.IsAfterInitialCapture)
                {
                    try
                    {
                        // This is the first post-import native snapshot. A mismatch in
                        // the selected FRX is expected evidence, not a mutation failure.
                        observation.RunAfterInitialCapture(Capture, target, () => {
                            if (deferredImported == null || deferredBindings == null)
                                throw new InvalidOperationException("The declared imported form has no deferred font binding.");
                            RestoreDeferredFormFonts(observation, deferredImported, deferredBindings);
                        });
                    }
                    catch (Exception error)
                    {
                        try { observation.Failure(error); }
                        catch (Exception receiptError)
                        {
                            throw new AggregateException("Deferred font transfer failed while recording its native outcome.", error, receiptError);
                        }
                        throw;
                    }
                }
                bool exact;
                try { exact = Capture().SameAs(target); observation?.Complete(exact); }
                catch (Exception error)
                {
                    try { observation?.Failure(error); }
                    catch (Exception receiptError) { throw new AggregateException("Root font observation failed while recording the final comparison failure.", error, receiptError); }
                    throw;
                }
                if (!exact)
                {
                    var error = new InvalidOperationException(UiText.Get("The VBE did not preserve the imported sources exactly. Use Restore or check the project."));
                    try { observation?.Failure(error); }
                    catch (Exception receiptError) { throw new AggregateException("Root font observation failed while recording the exact comparison refusal.", error, receiptError); }
                    throw error;
                }
            }
            catch (Exception error) { deferredPrimary = error; throw; }
            finally
            {
                if (deferredImported != null)
                    FormFontRestoration.ReleaseOwnedReferences(new[] { deferredImported }, value => {
                        if (System.Runtime.InteropServices.Marshal.IsComObject(value))
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
                    }, deferredPrimary);
            }
        }

        /// <summary>Reacquires the one imported component after the first post-import capture.</summary>
        /// <param name="observation">form font observation that supplies the observation for this operation.</param>
        /// <param name="imported">object that supplies the imported for this operation.</param>
        /// <param name="bindings">form font binding[] that supplies the bindings for this operation.</param>
        private void RestoreDeferredFormFonts(FormFontObservation observation, object imported,
            FormStreamPadding.FormFontBinding[] bindings)
        {
            dynamic project = CheckedProject();
            FormFontRestoration.RequireOwner((object)project);
            object components = null, current = null;
            Exception primary = null;
            try
            {
                components = project.VBComponents;
                current = ((dynamic)components).Item(observation.FormName);
                if (!VbeProjectHostPath.SameProject(current, imported))
                    throw new InvalidOperationException("The deferred imported UserForm identity changed after initial capture.");
                FormFontRestoration.Restore(current, bindings,
                    () => RequireImportedForm(observation.FormName, imported), observation);
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                FormFontRestoration.ReleaseOwnedReferences(new object[] { components, current }, value => {
                    if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
                }, primary);
            }
        }

        /// <summary>Exports only the verified imported component without reading or assigning font properties.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="imported">object that supplies the imported for this operation.</param>
        /// <returns>vba git snapshot produced by the operation for capture imported form on vba git project.</returns>
        private VbaGitSnapshot CaptureImportedForm(string name, object imported)
        {
            RequireImportedForm(name, imported);
            using (var scratch = new Scratch())
            {
                string file = Path.Combine(scratch.Path, name + ".frm");
                ((dynamic)imported).Export(file);
                RequireImportedForm(name, imported);
                string resource = Path.Combine(scratch.Path, name + ".frx");
                var component = new VbaGitComponent { Name = name, Type = 3, HasResources = File.Exists(resource) };
                var files = new Dictionary<string, byte[]> {
                    { component.FileName, VbaGitSnapshot.Utf8.GetBytes(Normalize(File.ReadAllText(file, NativeEncoding))) }
                };
                if (component.HasResources) files.Add(name + ".frx", File.ReadAllBytes(resource));
                return new VbaGitSnapshot(new VbaGitManifest { Components = new[] { component }, References = "" }, files);
            }
        }

        /// <summary>Balances fresh identity readback references without releasing the imported component lease.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="imported">object that supplies the imported for this operation.</param>
        private void RequireImportedForm(string name, object imported)
        {
            object components = null, current = null;
            try
            {
                dynamic project = CheckedProject(); components = project.VBComponents;
                current = ((dynamic)components).Item(name);
                if (!VbeProjectHostPath.SameProject(current, imported))
                    throw new InvalidOperationException("The imported form identity changed before font restoration.");
            }
            finally
            {
                if (current != null && System.Runtime.InteropServices.Marshal.IsComObject(current)) System.Runtime.InteropServices.Marshal.ReleaseComObject(current);
                if (components != null && System.Runtime.InteropServices.Marshal.IsComObject(components)) System.Runtime.InteropServices.Marshal.ReleaseComObject(components);
            }
        }

        /// <summary>Removes only the single leading blank line added by VBIDE's native form import.</summary>
        /// <param name="codeModule">object that supplies the code module for this operation.</param>
        /// <param name="exported">Text that supplies the exported value. Use the format required by the calling operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        internal static void RestoreFormImportCode(object codeModule, string exported, Action revalidate)
        {
            string source = Normalize(exported);
            var identity = Regex.Match(source, "^Attribute VB_Name = ", RegexOptions.Multiline);
            if (!identity.Success) throw new InvalidOperationException("The imported form has no exported identity.");
            // Hidden export attributes are not part of CodeModule.Lines. Preserve
            // every visible character, including intentional blank lines and EOF.
            string expected = string.Join("\n", source.Substring(identity.Index).Split('\n')
                .Where(line => !Regex.IsMatch(line, "^Attribute[ \\t]+[^=\\r\\n]+=[^\\r\\n]*$")));
            string observed = Code((dynamic)codeModule);
            // Native Lines omits the final export line terminator. Select that
            // representation only when the complete observed code proves it.
            // Final snapshot readback still requires the target's exact FRM text.
            if (observed != "\n" + expected && expected.EndsWith("\n", StringComparison.Ordinal) &&
                observed == "\n" + expected.Substring(0, expected.Length - 1))
                expected = expected.Substring(0, expected.Length - 1);
            if (observed != "\n" + expected) return;
            revalidate?.Invoke();
            if (Code((dynamic)codeModule) != observed)
                throw new InvalidOperationException("Imported form code changed before removing the native prefix.");
            // One mutation, never replayed if COM throws. The final capture still
            // checks every source/attribute and logical resource against the target.
            ((dynamic)codeModule).DeleteLines(1, 1);
            if (Code((dynamic)codeModule) != expected)
                throw new InvalidOperationException("The native form import prefix was not removed exactly. Use Restore.");
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
