using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Prévisualise et applique des renommages liés aux membres privés de modules de classe ordinaires.</summary>
    internal sealed partial class VbeCodeEdits
    {

        /// <summary>Factory locale de chemin temporaire ; chaque instance possède sa dépendance.</summary>
        private readonly Func<string> classMemberExportPath = () => Path.Combine(Path.GetTempPath(),
            "VBAi-ClassMembers-" + Guid.NewGuid().ToString("N") + ".cls");

        /// <summary>Injecte le chemin temporaire possédé, sans modifier la factory des autres instances.</summary>
        /// <param name="execute">Transport des commandes VBE.</param>
        /// <param name="exportPath">Factory de chemin temporaire pour les exports de classe.</param>
        internal VbeCodeEdits(Func<Request, Response> execute, Func<string> exportPath) : this(execute)
        { classMemberExportPath = exportPath ?? throw new ArgumentNullException(nameof(exportPath)); }

        /// <summary>Projet complet et export de classe qualifié, sans attribut de membre caché.</summary>
        private sealed class ClassMemberProjectSnapshot
        {

            /// <summary>Snapshot des modules VBA du projet inspecté.</summary>
            internal ProcedureProjectSnapshot Project;

            /// <summary>Empreinte combinée du catalogue de projet et de l’export natif du composant cible.</summary>
            internal string Version, ExportVersion;
        }

        /// <summary>Prévisualise un membre privé de classe et ses références internes réellement liées.</summary>
        /// <param name="request">Projet, module, nom à rechercher et nouveau nom demandé.</param>
        /// <returns>Plan de renommage, empreintes de précondition et description des limites de liaison.</returns>
        internal object PreviewClassMemberRename(Request request)
        {
            var snapshot = CaptureClassMemberProject(request);
            var plan = VbaClassMemberRename.Prepare(snapshot.Project.CanonicalProjectName, snapshot.Project.Modules, request);
            return new
            {
                snapshot.Project.Project,
                snapshot.Project.CanonicalProjectName,
                request.Module,
                request.Query,
                request.NewName,
                ExpectedProjectVersion = snapshot.Version,
                snapshot.ExportVersion,
                plan.Edits,
                Changed = plan.Edits.Count > 0,
                Modules = snapshot.Project.Modules.Select(x => new { x.Name, x.Type, Sha256 = VbaProcedureRename.Digest(x.Source) }).ToArray(),
                Scope = "Explicit Private ordinary-class Sub/Function or complete private Property accessor family. Direct internal references, Me references and return names only. Public/Friend, typed receivers, hidden member attributes, callbacks and dynamic binding are refused."
            };
        }

        /// <summary>Vérifie tout le projet et l'export caché avant écriture, puis relit code et attributs après mutation.</summary>
        /// <param name="request">Requête correspondant à l’aperçu, avec mode et version attendus.</param>
        /// <returns>Résultat vérifié par module, avec état non atomique et limites d’annulation.</returns>
        /// <exception cref="ArgumentException">La requête ne porte pas le mode design ou la version d’aperçu attendue.</exception>
        /// <exception cref="InvalidOperationException">Le projet a changé ou une écriture/readback a échoué.</exception>
        internal object ApplyClassMemberRename(Request request)
        {
            if (request == null || request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("ExpectedMode=2 and ExpectedProjectVersion from preview_class_member_rename are required.");
            var snapshot = CaptureClassMemberProject(request);
            if (!string.Equals(snapshot.Version, request.ExpectedProjectVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The complete project or native class export changed; preview the member rename again.");
            var plan = VbaClassMemberRename.Prepare(snapshot.Project.CanonicalProjectName, snapshot.Project.Modules, request);
            var preflight = CaptureClassMemberProject(request);
            if (!string.Equals(preflight.Version, snapshot.Version, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The complete project changed during preparation; no module was written.");
            var completed = new List<object>();
            string current = null;
            try
            {
                foreach (var edit in plan.Edits)
                {
                    current = edit.Module;
                    Write(snapshot.Project.Project, edit.Module, edit.Before, edit.After);
                    string exportVersion = InspectClassMemberExport(snapshot.Project.Project, edit.Module, edit.After);
                    completed.Add(new
                    {
                        edit.Module,
                        BeforeSha256 = edit.ExpectedSha256,
                        AfterSha256 = VbaProcedureRename.Digest(edit.After),
                        ExportVersion = exportVersion,
                        edit.Replacements,
                        ReadbackVerified = true
                    });
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Class member rename stopped while writing " + current + ". " + completed.Count +
                    " module(s) were already verified. The failing module may also have changed. Read every module before continuing; " +
                    "existing per-module history is retained. No automatic rollback or retry was attempted. " + ex.Message, ex);
            }
            return new
            {
                snapshot.Project.Project,
                request.Module,
                request.Query,
                request.NewName,
                Modules = completed,
                ReadbackVerified = true,
                Atomic = false,
                MultiModuleUndoAvailable = false,
                History = "Existing session history per module; undo only with the current SHA.",
                Limit = "Private internal member binding only; no compilation or host runtime acceptance is implied."
            };
        }

        /// <summary>Ajoute la version de l'export natif au catalogue et SHA de tous les composants.</summary>
        /// <param name="request">Projet et module de classe sélectionné.</param>
        /// <returns>Snapshot combiné du catalogue de projet et des métadonnées de l’export cible.</returns>
        private ClassMemberProjectSnapshot CaptureClassMemberProject(Request request)
        {
            var project = CaptureProcedureProject(request);
            if (project.Mode != 2) throw new InvalidOperationException("Class member refactoring requires native design mode.");
            var target = project.Modules.SingleOrDefault(x => string.Equals(x.Name, request.Module, StringComparison.OrdinalIgnoreCase));
            if (target == null || target.Type != 2) throw new InvalidOperationException("An exact ordinary class module is required.");
            string export = InspectClassMemberExport(project.Project, target.Name, target.Source);
            return new ClassMemberProjectSnapshot
            {
                Project = project,
                ExportVersion = export,
                Version = VbaProcedureRename.Digest(project.Version + ":" + export)
            };
        }

        /// <summary>Exporte en lecture un seul fichier temporaire possédé pour inspecter les attributs omis par CodeModule.</summary>
        /// <param name="project">Projet qui contient le module.</param>
        /// <param name="module">Nom du module de classe à exporter.</param>
        /// <param name="source">Source CodeModule capturée avant l’export.</param>
        /// <returns>Empreinte de l’export après vérification de son code et de ses attributs.</returns>
        private string InspectClassMemberExport(string project, string module, string source)
        {
            var component = execute(new Request { Command = "component_properties", Project = project, Module = module });
            if (!component.Ok) throw new InvalidOperationException(component.Error);
            dynamic state = component.Data;
            if ((int)state.Type != 2 || string.IsNullOrWhiteSpace((string)state.Version))
                throw new InvalidOperationException("The native class component version is unavailable.");
            string path = classMemberExportPath();
            if (File.Exists(path)) throw new InvalidOperationException("The temporary class export path already exists.");
            try
            {
                var exported = execute(new Request
                {
                    Command = "export_component",
                    Project = project,
                    Module = module,
                    Path = path,
                    ExpectedComponentVersion = (string)state.Version
                });
                if (!exported.Ok) throw new InvalidOperationException(exported.Error);
                if (!File.Exists(path) || new FileInfo(path).Length > 4 * 1024 * 1024)
                    throw new InvalidOperationException("The native class export is unavailable or exceeds its bound.");
                var encoding = Encoding.GetEncoding(Encoding.Default.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                string text = encoding.GetString(File.ReadAllBytes(path));
                ValidateClassMemberExport(text, module, source);
                return VbaProcedureRename.Digest(text);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        /// <summary>Refuse les attributs de membres et exige le préambule standard d'une classe non exposée.</summary>
        /// <param name="export">Export natif de classe complet.</param>
        /// <param name="module">Nom attendu pour l’attribut VB_Name.</param>
        /// <param name="source">Source visible capturée par CodeModule.</param>
        /// <exception cref="InvalidOperationException">L’export n’a pas le format attendu ou le code/métadonnées diffère de l’instantané.</exception>
        internal static void ValidateClassMemberExport(string export, string module, string source)
        {
            if (export == null || source == null) throw new InvalidOperationException("The native class export is unreadable.");
            string[] lines = CodeRollback.Lines(export);
            if (lines.Length < 9 || lines[0] != "VERSION 1.0 CLASS" || lines[1] != "BEGIN" ||
                !Regex.IsMatch(lines[2], @"^\s*MultiUse\s*=\s*-1\s*'True\s*$", RegexOptions.IgnoreCase) || lines[3] != "END")
                throw new InvalidOperationException("The native class export header is not the qualified ordinary-class format.");
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int body = 4;
            while (body < lines.Length && lines[body].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                var match = Regex.Match(lines[body], @"^Attribute (VB_Name|VB_GlobalNameSpace|VB_Creatable|VB_PredeclaredId|VB_Exposed) = (.+)$", RegexOptions.IgnoreCase);
                if (!match.Success || attributes.ContainsKey(match.Groups[1].Value))
                    throw new InvalidOperationException("Hidden member or unsupported class attributes require a separate consumer plan.");
                attributes.Add(match.Groups[1].Value, match.Groups[2].Value); body++;
            }
            if (attributes.Count != 5 || attributes["VB_Name"] != "\"" + module + "\"" ||
                attributes.Where(x => !string.Equals(x.Key, "VB_Name", StringComparison.OrdinalIgnoreCase)).Any(x => !string.Equals(x.Value, "False", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Only the exact non-predeclared, non-exposed ordinary class header is qualified.");
            string code = string.Join("\r\n", lines.Skip(body));
            if (Regex.IsMatch(code, @"(?im)^\s*Attribute\b") ||
                !string.Equals(code.TrimEnd('\r', '\n'), source.TrimEnd('\r', '\n'), StringComparison.Ordinal))
                throw new InvalidOperationException("The exported code differs from the snapshot or contains hidden member attributes.");
        }
    }
}
