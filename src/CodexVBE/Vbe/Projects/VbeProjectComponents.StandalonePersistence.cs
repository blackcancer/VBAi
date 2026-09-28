using System;
using System.IO;

namespace CodexVBE
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Détecte uniquement un projet VBA autonome de macro SWP, jamais un projet intégré à un document Office.</summary>
        private static bool SupportsStandaloneMacro(dynamic project)
        {
            try
            {
                if ((int)project.Type != 101) return false; // vbext_pt_StandAlone
                string path = (string)project.FileName;
                return string.IsNullOrWhiteSpace(path) || Path.GetExtension(path).Equals(".swp", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }
        /// <summary>Lit le fichier natif d’une macro autonome et son état de sauvegarde VBIDE.</summary>
        private static object StandalonePersistence(string selector, dynamic project)
        {
            string path = (string)project.FileName;
            bool hasPath = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
            bool exists = hasPath && File.Exists(path);
            bool? readOnly = exists ? (bool?)((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0) : null;
            return new { Project = selector, ProjectSaved = (bool)project.Saved, HostAvailable = true,
                HostPath = hasPath ? Path.GetFullPath(path) : null, HostSaved = exists ? (bool?)(bool)project.Saved : null,
                HostReadOnly = readOnly, HostHasPath = (bool?)hasPath, FileExists = exists,
                SaveApi = "Standalone VBProject.SaveAs", Reason = (string)null,
                Limit = "Native standalone SWP only. A save/readback of flags and file metadata is not proof of reload fidelity or signature trust." };
        }
        /// <summary>Enregistre une macro SWP autonome par VBIDE après contrôle de sa version et de son chemin.</summary>
        private object SaveStandaloneMacro(Request request, bool saveAs)
        {
            dynamic project = GetDesignProject(request.Project);
            if (!SupportsStandaloneMacro((object)project))
                throw new InvalidOperationException("This host project has no supported save API. VBProject.SaveAs is only available for standalone SWP projects (Type=101).");
            AssertProjectVersion(request, project);
            string original = (string)project.FileName;
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
