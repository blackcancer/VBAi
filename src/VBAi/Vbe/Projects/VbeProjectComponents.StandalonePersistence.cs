using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Gère l’enregistrement natif des projets VBA autonomes de macros SWP.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Normalizes VBIDE's path-not-found state only for an unsaved standalone project.</summary>
        /// <param name="project">dynamic that supplies the project for this operation.</param>
        /// <returns>Text produced by the operation for standalone aware project path on vbe project components.</returns>
        private static string StandaloneAwareProjectPath(dynamic project)
        {
            try { return (string)project.FileName; }
            catch (COMException error) when (error.ErrorCode == unchecked((int)0x800A004C) &&
                (int)project.Type == 101 && !(bool)project.Saved)
            {
                // Newly added standalone projects may not expose FileName before their first SaveAs.
                // Other HRESULTs, saved projects and host-document projects retain their failures.
                return "";
            }
            catch (DirectoryNotFoundException error) when (error.HResult == unchecked((int)0x80070003) &&
                (int)project.Type == 101 && !(bool)project.Saved)
            {
                return "";
            }
        }

        /// <summary>Détecte uniquement un projet VBA autonome de macro SWP, jamais un projet intégré à un document Office.</summary>
        /// <param name="project">Projet VBIDE à examiner.</param>
        /// <returns><see langword="true"/> si le type et le chemin correspondent à une macro autonome SWP.</returns>
        private static bool SupportsStandaloneMacro(dynamic project)
        {
            try
            {
                if ((int)project.Type != 101) return false; // vbext_pt_StandAlone
                string path = StandaloneAwareProjectPath((object)project);
                return string.IsNullOrWhiteSpace(path) || Path.GetExtension(path).Equals(".swp", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>Lit le fichier natif d’une macro autonome et son état de sauvegarde VBIDE.</summary>
        /// <param name="selector">Identifiant du projet à inclure dans le résultat.</param>
        /// <param name="project">Projet VBIDE dont l’état de persistance est lu.</param>
        /// <returns>Un instantané sérialisable des indicateurs de sauvegarde et du fichier hôte.</returns>
        private object StandalonePersistence(string selector, dynamic project)
        {
            bool solidWorksDraft = SolidWorksSaveProbe().IsSolidWorks;
            string path = StandaloneAwareProjectPath((object)project);
            bool hasPath = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
            bool exists = hasPath && File.Exists(path);
            bool? readOnly = exists ? (bool?)((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0) : null;
            return new { Project = selector, ProjectSaved = (bool)project.Saved, HostAvailable = true,
                HostPath = hasPath ? Path.GetFullPath(path) : null, HostSaved = exists ? (bool?)(bool)project.Saved : null,
                HostReadOnly = readOnly, HostHasPath = (bool?)hasPath, FileExists = exists,
                SaveApi = solidWorksDraft ? null : "Standalone VBProject.SaveAs", SaveSupported = !solidWorksDraft,
                Reason = solidWorksDraft ? "A generic Type101 SaveAs does not create a native SOLIDWORKS container. Use publish_solidworks_macro to retain this draft and create a new native macro, or create_solidworks_macro with a path from the start." : null,
                Limit = "Generic standalone VBIDE persistence is distinct from native SOLIDWORKS SWP hosting. File metadata is not proof of reload fidelity or signature trust." };
        }

        /// <summary>Enregistre une macro SWP autonome par VBIDE après contrôle de sa version et de son chemin.</summary>
        /// <param name="request">Requête contenant le projet, sa version attendue et éventuellement le chemin de destination.</param>
        /// <param name="saveAs">Sélectionne une nouvelle destination lorsque la valeur est vraie.</param>
        /// <returns>Le résultat de l’enregistrement et la vérification du chemin, de l’état et du fichier produit.</returns>
        private object SaveStandaloneMacro(Request request, bool saveAs)
        {
            if (SolidWorksSaveProbe().IsSolidWorks)
                throw new InvalidOperationException("Generic Type101 Save/SaveAs is refused in SOLIDWORKS because it does not create a native SWP container. Use publish_solidworks_macro to preserve the draft and create a new native identity, or create_solidworks_macro with an explicit path from the start.");
            dynamic project = GetDesignProject(request.Project);
            if (!SupportsStandaloneMacro((object)project))
                throw new InvalidOperationException("This host project has no supported save API. VBProject.SaveAs is only available for standalone SWP projects (Type=101).");
            AssertProjectVersion(request, project);
            string original = StandaloneAwareProjectPath((object)project);
            string destination;
            if (saveAs)
            {
                if (string.IsNullOrWhiteSpace(request.Path) || !Path.IsPathRooted(request.Path) || !Path.GetExtension(request.Path).Equals(".swp", StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("A user-provided absolute .swp destination is required.");
                destination = Path.GetFullPath(request.Path);
                if (File.Exists(destination)) throw new IOException("SaveAs destination already exists: " + destination);
            }
            else
            {
                if (string.IsNullOrWhiteSpace(original) || !Path.IsPathRooted(original) || string.IsNullOrWhiteSpace(request.ExpectedHostPath) || !Path.IsPathRooted(request.ExpectedHostPath) ||
                    !string.Equals(Path.GetFullPath(original), Path.GetFullPath(request.ExpectedHostPath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The standalone macro path changed or was never saved; use an explicit SaveAs destination.");
                destination = Path.GetFullPath(original);
                if (!File.Exists(destination)) throw new FileNotFoundException("The expected saved macro file is absent.", destination);
                if ((File.GetAttributes(destination) & FileAttributes.ReadOnly) != 0) throw new InvalidOperationException("The macro file is read-only.");
            }
            if (!Directory.Exists(Path.GetDirectoryName(destination))) throw new DirectoryNotFoundException("The destination directory does not exist.");
            bool savedBefore = (bool)project.Saved;
            // Never use this method for host projects: Microsoft documents a runtime error there.
            project.SaveAs(destination);
            string actual = (string)project.FileName;
            bool saved = (bool)project.Saved;
            if (string.IsNullOrWhiteSpace(actual) || !Path.IsPathRooted(actual) || !string.Equals(Path.GetFullPath(actual), destination, StringComparison.OrdinalIgnoreCase) ||
                !saved || !File.Exists(destination) || new FileInfo(destination).Length == 0)
                throw new InvalidOperationException("Native SaveAs returned without matching saved project/file state; inspect the macro before retrying.");
            return new { Project = request.Project, HostPath = destination, SaveInvoked = true, SaveAsInvoked = saveAs,
                ProjectSavedBefore = savedBefore, ProjectSaved = saved, HostSaved = saved, Bytes = new FileInfo(destination).Length,
                Verification = "StandaloneProjectSaveAsAndFileReadback", ReloadVerified = false,
                Limit = "Native SWP reload and signature persistence must be qualified separately. No export/import or document Save API is substituted." };
        }
    }
}
