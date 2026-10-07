using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Renomme les projets Excel pris en charge et vérifie la conservation de leur code source.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Renomme un projet Excel enregistré, non protégé et identifié par son chemin stable.</summary>
        /// <param name="request">Requête contenant le chemin exact du classeur, la nouvelle valeur et la version attendue.</param>
        /// <returns>Les propriétés relues et les résultats de vérification du nom, du chemin et des sources.</returns>
        private object RenameSavedExcelProject(Request request)
        {
            if (!host.IsExcel || string.IsNullOrWhiteSpace(request.Project) || !Path.IsPathRooted(request.Project))
                throw new InvalidOperationException("Project rename is disabled for unqualified/unsaved/other-host scopes; select a saved Excel workbook by its exact FileName.");
            if (!(request.Value is string name) || !Regex.IsMatch(name, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("A project name of 1-40 ASCII identifier characters is required.");
            VbaTextEdits.ValidateIdentifier(name);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            int protection;
            try { protection = (int)project.Protection; }
            catch { throw new InvalidOperationException("Project protection could not be read; rename was not attempted."); }
            if (protection != 0) throw new InvalidOperationException("Unlock the project in the native IDE before renaming it; no protection is bypassed.");
            dynamic workbook = MatchExcelWorkbook(project, false);
            if ((bool)workbook.ReadOnly) throw new InvalidOperationException("The workbook is read-only.");
            string originalPath = (string)project.FileName, oldName = (string)project.Name;
            if (!string.Equals(Path.GetFullPath(originalPath), Path.GetFullPath(request.Project), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The workbook path changed before rename.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The project name would collide with one of its components.");
            var beforeSources = ProjectSourceVersions((object)project);
            // Use the typed VBIDE setter, rather than a generic PropertyDescriptor dispatch.
            if (!string.Equals(oldName, name, StringComparison.Ordinal)) project.Name = name;
            var afterSources = ProjectSourceVersions((object)project);
            bool sourcePreserved = beforeSources.Count == afterSources.Count && beforeSources.All(x => afterSources.TryGetValue(x.Key, out string sha) && sha == x.Value);
            bool pathPreserved = string.Equals((string)project.FileName, originalPath, StringComparison.OrdinalIgnoreCase);
            bool nameVerified = string.Equals((string)project.Name, name, StringComparison.Ordinal);
            dynamic snapshot = ProjectProperties(originalPath);
            return new
            {
                Project = (string)snapshot.Project,
                Mode = (int)snapshot.Mode,
                Version = (string)snapshot.Version,
                Properties = (object)snapshot.Properties,
                Components = (object)snapshot.Components,
                References = (object)snapshot.References,
                BeforeName = oldName,
                AfterName = (string)project.Name,
                Verified = nameVerified && pathPreserved && sourcePreserved,
                SourcePreserved = sourcePreserved,
                HostPathPreserved = pathPreserved,
                CompilationVerified = false,
                PersistenceVerified = false,
                NextRead = "project_properties, project_persistence_status",
                Limit = "Project metadata only; qualified source references and external callers are not refactored. Save/reopen separately. Protected, unsaved and other-host projects remain refused."
            };
        }

        /// <summary>Empreintes des modules utilisées pour détecter une mutation inattendue du code durant le renommage.</summary>
        /// <param name="project">Projet dont les modules de code sont lus.</param>
        /// <returns>Une table nom de composant vers empreinte du texte source.</returns>
        private static Dictionary<string, string> ProjectSourceVersions(dynamic project)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            long characters = 0;
            foreach (dynamic component in project.VBComponents)
            {
                dynamic module = component.CodeModule;
                int lines = (int)module.CountOfLines;
                if (lines < 0 || lines > 100000) throw new InvalidOperationException("The project source is too large to verify safely.");
                string source = lines == 0 ? "" : (string)module.Lines(1, lines);
                characters += source.Length;
                if (characters > 4 * 1024 * 1024) throw new InvalidOperationException("The project source is too large to verify safely.");
                result.Add((string)component.Name, Hash(source));
            }
            return result;
        }
    }
}
