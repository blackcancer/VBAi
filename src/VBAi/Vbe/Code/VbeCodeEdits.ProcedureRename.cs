using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Prévisualise les renommages de procédures standard et vérifie le projet complet avant leurs écritures multiples.</summary>
    internal sealed partial class VbeCodeEdits
    {

        /// <summary>Instantané complet du projet, métadonnées et sources, relu avant toute mutation.</summary>
        private sealed class ProcedureProjectSnapshot
        {

            /// <summary>Sélecteur du projet accepté par le transport VBE.</summary>
            internal string Project, CanonicalProjectName, Version;

            /// <summary>Mode courant du projet VBA.</summary>
            internal int Mode;

            /// <summary>Sources complètes et types de tous les composants inspectés.</summary>
            internal VbaProcedureRename.ModuleSnapshot[] Modules;
        }

        /// <summary>Prévisualise toutes les éditions et leur SHA avant renommage d'une procédure standard.</summary>
        /// <param name="request">Projet, module, déclaration exacte, Query, NewName et SHA cible inspecté.</param>
        /// <returns>Version complète à fournir dans ExpectedProjectVersion lors de l'application.</returns>
        internal object PreviewProcedureRename(Request request)
        {
            var snapshot = CaptureProcedureProject(request);
            var plan = VbaProcedureRename.Prepare(snapshot.CanonicalProjectName, snapshot.Modules, request);
            return new { snapshot.Project, snapshot.CanonicalProjectName, request.Module, request.Query, request.NewName,
                ExpectedProjectVersion = snapshot.Version, Edits = plan.Edits, Changed = plan.Edits.Count > 0,
                Modules = snapshot.Modules.Select(x => new { x.Name, x.Type, Sha256 = VbaProcedureRename.Digest(x.Source) }).ToArray(),
                Scope = "One standard-module Sub/Function and resolved direct or project/module-qualified calls in this complete project. Other projects, external consumers and strings are not rewritten. Conditional, callback, dynamic and shadowed binding is refused." };
        }

        /// <summary>Vérifie tout le projet, puis écrit et relit chaque module; les éditions validées restent dans l'historique par module.</summary>
        /// <param name="request">Même identité que la prévisualisation, ExpectedProjectVersion et ExpectedMode=2.</param>
        /// <returns>Modules relus; cette opération multi-module ne promet pas une transaction ou un undo global.</returns>
        internal object ApplyProcedureRename(Request request)
        {
            if (request == null || request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("ExpectedMode=2 and ExpectedProjectVersion from preview_procedure_rename are required.");
            var snapshot = CaptureProcedureProject(request);
            if (snapshot.Mode != 2 || !string.Equals(snapshot.Version, request.ExpectedProjectVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project mode, metadata, component catalogue or source changed; preview the rename again.");
            var plan = VbaProcedureRename.Prepare(snapshot.CanonicalProjectName, snapshot.Modules, request);
            var preflight = CaptureProcedureProject(request);
            if (preflight.Mode != 2 || !string.Equals(preflight.Version, snapshot.Version, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The complete project changed during preparation; no module was written.");
            var completed = new List<object>();
            string currentModule = null;
            try
            {
                foreach (var edit in plan.Edits)
                {
                    currentModule = edit.Module;
                    // replace_lines guards this module's SHA again and records its existing session history.
                    Write(snapshot.Project, edit.Module, edit.Before, edit.After);
                    completed.Add(new { edit.Module, BeforeSha256 = edit.ExpectedSha256,
                        AfterSha256 = VbaProcedureRename.Digest(edit.After), edit.Replacements, ReadbackVerified = true });
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Procedure rename stopped while writing " + currentModule + ". " +
                    completed.Count + " module(s) were already verified. The failing module may also have changed. " +
                    "Read every module before continuing; completed edits use the existing per-module history. " +
                    "No automatic rollback or retry was attempted. " + ex.Message, ex);
            }
            return new { snapshot.Project, request.Query, request.NewName, Modules = completed,
                ReadbackVerified = true, Atomic = false, MultiModuleUndoAvailable = false,
                History = "Existing session history per module; undo each module after inspecting its current SHA.",
                Limit = "Other projects and external/string consumers are outside this plan. No compile or host runtime acceptance is implied." };
        }

        /// <summary>Collecte toutes les sources et vérifie la stabilité du catalogue et des références pendant cette lecture.</summary>
        /// <param name="request">Projet et module cible pour lequel capturer l’état complet.</param>
        /// <returns>Snapshot cohérent des métadonnées, mode, composants et sources VBA.</returns>
        private ProcedureProjectSnapshot CaptureProcedureProject(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.Module))
                throw new ArgumentException("An exact Project and Module are required.");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            var metadata = ProcedureMetadata(request.Project, serializer);
            var response = execute(new Request { Command = "list_modules", Project = request.Project });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var rows = serializer.DeserializeObject(serializer.Serialize(response.Data)) as object[];
            if (rows == null || rows.Length < 1 || rows.Length > 1000) throw new InvalidOperationException("The complete component catalogue is unavailable.");
            var modules = new List<VbaProcedureRename.ModuleSnapshot>();
            foreach (object item in rows)
            {
                var row = item as IDictionary<string, object>;
                if (row == null || !row.ContainsKey("Name") || !row.ContainsKey("Type"))
                    throw new InvalidOperationException("A component identity or type is unreadable.");
                string name = Convert.ToString(row["Name"]);
                if (string.IsNullOrWhiteSpace(name) || modules.Any(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("The component identity is empty or ambiguous.");
                modules.Add(new VbaProcedureRename.ModuleSnapshot(name, Convert.ToInt32(row["Type"]), Read(request.Project, name)));
            }
            var observed = ProcedureMetadata(request.Project, serializer);
            if (!string.Equals((string)metadata["Version"], (string)observed["Version"], StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Convert.ToString(metadata["Project"]), Convert.ToString(observed["Project"]), StringComparison.Ordinal))
                throw new InvalidOperationException("Project metadata or component catalogue changed during inspection.");
            string project = Convert.ToString(metadata["Project"]);
            if (string.IsNullOrWhiteSpace(project)) throw new InvalidOperationException("The native project name is unreadable.");
            return new ProcedureProjectSnapshot { Project = request.Project, CanonicalProjectName = project, Mode = Convert.ToInt32(metadata["Mode"]), Modules = modules.ToArray(),
                Version = VbaProcedureRename.Digest((string)metadata["Version"] + ":" + VbaProcedureRename.Version(project, modules)) };
        }

        /// <summary>Lit la version VBIDE des propriétés, références, types et identités de composants.</summary>
        /// <param name="project">Sélecteur du projet VBA.</param>
        /// <param name="serializer">Sérialiseur utilisé pour normaliser le résultat du transport.</param>
        /// <returns>Champs de métadonnées du projet, dont nom, mode et version.</returns>
        private IDictionary<string, object> ProcedureMetadata(string project, JavaScriptSerializer serializer)
        {
            var response = execute(new Request { Command = "project_properties", Project = project });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var fields = serializer.DeserializeObject(serializer.Serialize(response.Data)) as IDictionary<string, object>;
            if (fields == null || !fields.ContainsKey("Project") || !fields.ContainsKey("Mode") || !fields.ContainsKey("Version") ||
                string.IsNullOrWhiteSpace(Convert.ToString(fields["Version"])))
                throw new InvalidOperationException("The project metadata snapshot is incomplete.");
            return fields;
        }
    }
}
