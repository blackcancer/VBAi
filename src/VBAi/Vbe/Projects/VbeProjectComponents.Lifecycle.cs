using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBAi
{
    /// <summary>Inventorie les projets VBIDE et contrôle leurs opérations de cycle de vie autonome.</summary>
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Lit la collection native avec identités, sauvegarde et empreintes de source.</summary>
        /// <returns>Version à fournir à la création ou à l'ouverture d'un projet autonome.</returns>
        public object ProjectCollectionState()
        {
            var rows = ReadLifecycleCollection();
            return new { Version = LifecycleVersion(rows), Projects = rows,
                NativeApi = "VBProjects.Add/Open/Remove", RuntimeQualified = false };
        }

        /// <summary>Crée un projet autonome par l'API de l'hôte après vérification de collection.</summary>
        /// <param name="request">Version attendue de la collection dans ExpectedProjectVersion.</param>
        /// <returns>Résultat natif vérifié ou état incertain sans nouvel essai.</returns>
        public object CreateStandaloneProject(Request request)
        {
            var before = RequireLifecycleCollection(request);
            return AddLifecycleProject(before, null);
        }

        /// <summary>Ouvre une macro SWP existante par VBProjects.Open.</summary>
        /// <param name="request">Chemin absolu SWP et version attendue de collection.</param>
        /// <returns>Identité relue et état vérifié, ou résultat incertain après invocation.</returns>
        public object OpenStandaloneProject(Request request)
        {
            if (request == null) throw new ArgumentException("Request is required.");
            // VBProjects.Open terminated the native SOLIDWORKS 2019 host during qualification.
            // Refuse before accessing its collection; never retry through another native API.
            if (SolidWorksSaveProbe().IsSolidWorks)
                throw new NotSupportedException("Opening a standalone SWP through VBProjects.Open is disabled in SOLIDWORKS after an observed host termination. Use SOLIDWORKS Tools > Macro > Edit instead.");
            string path = RequireAbsolutePath(request.Path);
            if (!Path.GetExtension(path).Equals(".swp", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Only an absolute standalone .swp macro is supported.");
            if (!File.Exists(path)) throw new FileNotFoundException("Macro file is absent.", path);
            long bytes = new FileInfo(path).Length;
            if (bytes == 0 || bytes > 64 * 1024 * 1024)
                throw new IOException("Macro size must be between 1 byte and 64 MiB.");
            var before = RequireLifecycleCollection(request);
            if (before.Any(row => string.Equals(row.Path, path, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The macro is already open.");
            return AddLifecycleProject(before, path);
        }

        /// <summary>Ferme uniquement un projet autonome déjà sauvegardé sans supprimer son fichier.</summary>
        /// <param name="request">Sélecteur, version de project_properties et ExpectedHostPath exact.</param>
        /// <returns>Résultat de retrait vérifié ou incertitude après appel natif.</returns>
        public object CloseStandaloneProject(Request request)
        {
            if (request == null) throw new ArgumentException("Request is required.");
            string path = RequireAbsolutePath(request.ExpectedHostPath);
            dynamic project = GetDesignProject(request.Project);
            if ((int)project.Type != 101 || (int)project.Protection != 0)
                throw new InvalidOperationException("Only an unprotected standalone project can be closed.");
            if (!(bool)project.Saved)
                throw new InvalidOperationException("Save the standalone project before closing it.");
            if (!SupportsStandaloneMacro((object)project) ||
                !string.Equals((string)project.FileName, path, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The standalone project path changed.");
            AssertProjectVersion(request, project);
            var before = ReadLifecycleCollection();
            var selected = before.Single(row => row.Name == (string)project.Name &&
                string.Equals(row.Path, path, StringComparison.OrdinalIgnoreCase));
            RequireLifecycleDesign(before);
            try
            {
                vbe.VBProjects.Remove(project);
                var after = ReadLifecycleCollection();
                var remaining = before.Where(row => row.Identity != selected.Identity).ToList();
                if (LifecycleVersion(after) != LifecycleVersion(remaining))
                    throw new InvalidOperationException("Collection readback does not match the single requested removal.");
                return new { Verified = true, MutationInvoked = true, Uncertain = false,
                    Project = selected.Name, HostPath = path, CollectionState = ProjectCollectionState() };
            }
            catch (Exception ex) { return LifecycleUncertain("Remove", ex); }
        }

                /// <summary>Invoque Add ou Open exactement une fois et vérifie tous les projets préexistants.</summary>
                /// <param name="before">Inventaire validé avant la mutation.</param>
                /// <param name="path">Chemin SWP à ouvrir, ou <see langword="null"/> pour créer un projet autonome.</param>
                /// <returns>État vérifié ou résultat indiquant une issue incertaine après l’appel natif.</returns>
        private object AddLifecycleProject(List<LifecycleProject> before, string path)
        {
            try
            {
                dynamic added = path == null ? vbe.VBProjects.Add(101) : vbe.VBProjects.Open(path);
                var after = ReadLifecycleCollection();
                var extra = after.Where(candidate => !before.Any(old => old.Identity == candidate.Identity)).ToList();
                var retained = after.Where(candidate => before.Any(old => old.Identity == candidate.Identity)).ToList();
                if (extra.Count != 1 || after.Count != before.Count + 1 ||
                    LifecycleVersion(retained) != LifecycleVersion(before))
                    throw new InvalidOperationException("Collection readback does not match a single added project.");
                var row = extra[0];
                if (row.Type != 101 || row.Mode != 2 || row.Protection != 0 ||
                    row.Name != (string)added.Name ||
                    !string.Equals(row.Path, StandaloneAwareProjectPath((object)added), StringComparison.OrdinalIgnoreCase) ||
                    (path != null && !string.Equals(row.Path, path, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("The native returned project does not match the standalone readback.");
                return new { Verified = true, MutationInvoked = true, Uncertain = false,
                    Project = row.Name, HostPath = row.Path, CollectionState = ProjectCollectionState() };
            }
            catch (Exception ex) { return LifecycleUncertain(path == null ? "Add" : "Open", ex); }
        }

                /// <summary>Vérifie la version et le mode de chaque projet avant une mutation de collection.</summary>
                /// <param name="request">Requête avec l’empreinte préalablement lue.</param>
                /// <returns>La collection actuelle, si sa version correspond et si tous les projets sont modifiables.</returns>
        private List<LifecycleProject> RequireLifecycleCollection(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("ExpectedProjectVersion from project_collection_state is required.");
            var rows = ReadLifecycleCollection();
            if (!string.Equals(request.ExpectedProjectVersion, LifecycleVersion(rows), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project collection changed since it was read.");
            RequireLifecycleDesign(rows);
            return rows;
        }

                /// <summary>Interdit toute mutation lorsque la source ou le mode d'un projet est indisponible.</summary>
                /// <param name="rows">États des projets qui seront affectés par la mutation de collection.</param>
        private static void RequireLifecycleDesign(List<LifecycleProject> rows)
        {
            if (rows.Any(row => row.Mode != 2 || row.Protection != 0))
                throw new InvalidOperationException("Every open project must be unprotected and in design mode.");
        }

                /// <summary>Capture tous les projets sans dépendre du projet actif.</summary>
                /// <returns>Les projets triés par identité après lecture de leurs propriétés, références et sources accessibles.</returns>
        private List<LifecycleProject> ReadLifecycleCollection()
        {
            var rows = new List<LifecycleProject>();
            foreach (dynamic project in vbe.VBProjects)
            {
                string name = (string)project.Name;
                string path = StandaloneAwareProjectPath((object)project);
                var components = new List<object>();
                if ((int)project.Protection == 0)
                    foreach (dynamic component in project.VBComponents)
                    {
                        dynamic code = component.CodeModule;
                        int count = (int)code.CountOfLines;
                        components.Add(new { Name = (string)component.Name, Type = (int)component.Type,
                            SourceSha256 = Hash(count == 0 ? "" : (string)code.Lines(1, count)),
                            Properties = ReadProperties((object)component) });
                    }
                var references = new List<object>();
                if ((int)project.Protection == 0)
                    foreach (dynamic reference in project.References)
                        references.Add(new { Guid = (string)reference.GUID, Major = (int)reference.Major,
                            Minor = (int)reference.Minor, IsBroken = (bool)reference.IsBroken, BuiltIn = (bool)reference.BuiltIn });
                rows.Add(new LifecycleProject { Name = name, Path = path,
                    Identity = name + "|" + path, Type = (int)project.Type, Mode = (int)project.Mode,
                    Protection = (int)project.Protection, Saved = (bool)project.Saved,
                    SourceSha256 = Hash(json.Serialize(new { Components = components, References = references, Properties = ReadProperties((object)project) })) });
            }
            if (rows.GroupBy(row => row.Identity, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
                throw new InvalidOperationException("Project identities are ambiguous.");
            return rows.OrderBy(row => row.Identity, StringComparer.OrdinalIgnoreCase).ToList();
        }

                /// <summary>Calcule l'empreinte complète et ordonnée de collection.</summary>
                /// <param name="rows">Lignes de projet déjà capturées.</param>
                /// <returns>Empreinte de la représentation sérialisée de la collection.</returns>
        private string LifecycleVersion(List<LifecycleProject> rows) { return Hash(json.Serialize(rows)); }

                /// <summary>Signale qu'une mutation native peut avoir eu lieu sans effectuer de récupération destructive.</summary>
                /// <param name="api">Nom de l’API native invoquée.</param>
                /// <param name="error">Erreur observée pendant l’opération ou sa vérification.</param>
                /// <returns>Résultat sérialisable marquant l’opération comme incertaine et sans nouvel essai autorisé.</returns>
        private static object LifecycleUncertain(string api, Exception error)
        {
            return new { Verified = false, MutationInvoked = true, Uncertain = true, NativeApi = api,
                Reason = error.Message + " (" + error.GetType().FullName + ", HRESULT 0x" + unchecked((uint)error.HResult).ToString("X8") + ")", RetryAllowed = false,
                Limit = "Inspect project_collection_state before deciding on another operation. No retry or rollback was performed." };
        }

        /// <summary>État sérialisable d'une identité de projet natif.</summary>
        internal sealed class LifecycleProject
        {
                        /// <summary>Identité composée du nom et du chemin.</summary>
                        /// <value>Chaîne utilisée pour distinguer ce projet dans la collection.</value>
            public string Identity { get; set; }
                        /// <summary>Nom natif du projet.</summary>
                        /// <value>Nom retourné par VBIDE.</value>
            public string Name { get; set; }
                        /// <summary>Chemin natif, éventuellement vide avant la première sauvegarde.</summary>
                        /// <value>Chemin du fichier hôte, ou chaîne vide si le projet n’est pas enregistré.</value>
            public string Path { get; set; }
                        /// <summary>Type natif du projet.</summary>
                        /// <value>Valeur de type de projet retournée par VBIDE.</value>
            public int Type { get; set; }
                        /// <summary>Mode natif du projet.</summary>
                        /// <value>Valeur de mode retournée par VBIDE.</value>
            public int Mode { get; set; }
                        /// <summary>Protection native du projet.</summary>
                        /// <value>Valeur de protection retournée par VBIDE.</value>
            public int Protection { get; set; }
                        /// <summary>État natif de sauvegarde.</summary>
                        /// <value><see langword="true"/> si VBIDE indique que le projet est enregistré.</value>
            public bool Saved { get; set; }
                        /// <summary>Empreinte des sources accessibles.</summary>
                        /// <value>Empreinte SHA-256 des sources du projet lisibles.</value>
            public string SourceSha256 { get; set; }
        }
    }
}
