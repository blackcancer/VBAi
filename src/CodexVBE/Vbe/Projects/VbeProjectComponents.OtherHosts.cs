using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;

namespace CodexVBE
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>État lu du document Word ou PowerPoint associé par identité COM.</summary>
        internal sealed class OtherHostDocumentState
        {
            /// <summary>Chemin complet, ou chaîne vide avant première sauvegarde.</summary>
            internal string Path;
            /// <summary>État Saved natif du document.</summary>
            internal bool Saved;
            /// <summary>État natif de lecture seule; un état indéterminé est traité comme lecture seule.</summary>
            internal bool ReadOnly;
            /// <summary>Format natif Word; null lorsque PowerPoint n'expose pas cette propriété.</summary>
            internal int? Format;
        }

        /// <summary>Frontière injectable d'identité, état et sauvegarde du seul processus hôte courant.</summary>
        internal interface IOtherHostProbe
        {
            /// <summary>Word, PowerPoint ou null pour un processus non reconnu.</summary>
            string HostKind { get; }
            /// <summary>PID du processus hébergeant l'add-in.</summary>
            int CurrentProcessId { get; }
            /// <summary>Obtient une application déjà ouverte, sans lancement d'hôte.</summary>
            object Application();
            /// <summary>Vérifie les handles natifs de cette application et retourne leur PID commun.</summary>
            uint ApplicationProcessId(object application);
            /// <summary>Énumère tous les documents ouverts de l'application vérifiée.</summary>
            IList<object> Documents(object application);
            /// <summary>Lit le VBProject du document, sans supposer un nom unique.</summary>
            object DocumentProject(object document);
            /// <summary>Compare les identités IUnknown COM des deux projets.</summary>
            bool SameProject(object first, object second);
            /// <summary>Lit chemin, Saved, lecture seule et format natif disponible.</summary>
            OtherHostDocumentState State(object document);
            /// <summary>Invoque Save, Word.SaveAs2 ou PowerPoint.SaveAs exactement une fois.</summary>
            void Save(object document, bool saveAs, string destination, int format);
            /// <summary>Indique si un fichier de destination existe.</summary>
            bool FileExists(string path);
            /// <summary>Indique si le dossier parent existe.</summary>
            bool DirectoryExists(string path);
            /// <summary>Lit la taille du fichier après invocation native.</summary>
            long FileLength(string path);
        }

        /// <summary>Sonde native bornée au processus Office courant; transport réel encore à qualifier dans Word/PowerPoint.</summary>
        internal sealed class NativeOtherHostProbe : IOtherHostProbe
        {
            /// <summary>Reads the real process kind by default; isolates host contracts during qualification.</summary>
            internal Func<string> ReadHostKind = CurrentHostKind;
            /// <summary>Resolves an existing ROT application without starting Office.</summary>
            internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;
            /// <summary>Reads the native window owner used by all application PID guards.</summary>
            internal Func<IntPtr, uint> ReadOwner = Owner;
            private static string CurrentHostKind() => RecognizeOtherHost(Process.GetCurrentProcess().ProcessName);
            /// <summary>Reconnaît exclusivement WINWORD.EXE ou POWERPNT.EXE.</summary>
            public string HostKind => ReadHostKind();
            /// <summary>PID du processus de l'add-in.</summary>
            public int CurrentProcessId => Process.GetCurrentProcess().Id;
            /// <summary>Résout le ROT, puis NativeOM dans un document appartenant au PID courant si nécessaire.</summary>
            public object Application()
            {
                if (HostKind == null) throw new InvalidOperationException("Only the current Word or PowerPoint process is supported.");
                try
                {
                    object registered = ReadActiveApplication(HostKind == "Word" ? "Word.Application" : "PowerPoint.Application");
                    if (ApplicationProcessId(registered) == (uint)CurrentProcessId) return registered;
                }
                catch (COMException) { }
                catch (RuntimeBinderException) { }
                var windows = new List<IntPtr>();
                VbeDebugWindows.EnumWindows((parent, parameter) => {
                    uint owner; VbeDebugWindows.GetWindowThreadProcessId(parent, out owner);
                    if (owner != (uint)CurrentProcessId) return true;
                    VbeDebugWindows.EnumChildWindows(parent, (child, childParameter) => {
                        VbeDebugWindows.GetWindowThreadProcessId(child, out owner);
                        var name = new StringBuilder(256); VbeDebugWindows.GetClassName(child, name, name.Capacity);
                        if (owner == (uint)CurrentProcessId && name.ToString() == (HostKind == "Word" ? "_WwG" : "paneClassDC")) windows.Add(child);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
                foreach (var window in windows)
                {
                    try
                    {
                        Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046"); object native;
                        if (VbeDebugWindows.AccessibleObjectFromWindow(window, 0xFFFFFFF0, ref dispatch, out native) != 0 || native == null) continue;
                        object application = ((dynamic)native).Application;
                        if (ApplicationProcessId(application) == (uint)CurrentProcessId) return application;
                    }
                    catch (COMException) { }
                    catch (RuntimeBinderException) { }
                }
                throw new InvalidOperationException("No running Office application was verified as belonging to this VBE PID.");
            }
            /// <summary>Vérifie HWND PowerPoint ou tous les HWND de fenêtres Word, sans sélectionner de document.</summary>
            public uint ApplicationProcessId(object application)
            {
                if (HostKind == "PowerPoint") return ReadOwner(new IntPtr(Convert.ToInt64(((dynamic)application).HWND)));
                if (HostKind != "Word") return 0;
                uint result = 0; int count = 0;
                foreach (dynamic window in ((dynamic)application).Windows)
                {
                    if (++count > 1000) throw new InvalidOperationException("Unexpected Word window count.");
                    uint observed = ReadOwner(new IntPtr(Convert.ToInt64(window.Hwnd)));
                    if (observed == 0 || (result != 0 && result != observed)) return 0;
                    result = observed;
                }
                return result;
            }
            /// <summary>Obtient le propriétaire Win32 d'une fenêtre, ou zéro pour un handle absent.</summary>
            private static uint Owner(IntPtr window)
            { uint owner; GetWindowThreadProcessId(window, out owner); return owner; }
            /// <summary>Énumère les documents sans ignorer une erreur d'accès au catalogue.</summary>
            public IList<object> Documents(object application)
            {
                var result = new List<object>();
                System.Collections.IEnumerable documents = HostKind == "Word" ? (System.Collections.IEnumerable)((dynamic)application).Documents : (System.Collections.IEnumerable)((dynamic)application).Presentations;
                foreach (object document in documents)
                { if (result.Count >= 1000) throw new InvalidOperationException("Unexpected Office document count."); result.Add(document); }
                return result;
            }
            /// <summary>Lit l'objet projet attaché au document.</summary>
            public object DocumentProject(object document) => ((dynamic)document).VBProject;
            /// <summary>Compare IUnknown et libère seulement les références acquises par cette comparaison.</summary>
            public bool SameProject(object first, object second)
            {
                if (first == null || second == null || !Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
                IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
                try { a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (a != IntPtr.Zero) Marshal.Release(a); if (b != IntPtr.Zero) Marshal.Release(b); }
            }
            /// <summary>Lit les états natifs; aucune propriété FileFormat inexistante n'est inventée pour PowerPoint.</summary>
            public OtherHostDocumentState State(object document)
            {
                dynamic item = document; bool hasPath = !string.IsNullOrWhiteSpace((string)item.Path);
                return new OtherHostDocumentState { Path = hasPath ? (string)item.FullName : "",
                    ReadOnly = HostKind == "Word" ? (bool)item.ReadOnly : Convert.ToInt32(item.ReadOnly) != 0,
                    Saved = HostKind == "Word" ? (bool)item.Saved : Convert.ToInt32(item.Saved) == -1,
                    Format = HostKind == "Word" ? (int?)Convert.ToInt32(item.SaveFormat) : null };
            }
            /// <summary>Appelle la méthode native appropriée sans fermer, imprimer ou démarrer un document.</summary>
            public void Save(object document, bool saveAs, string destination, int format)
            {
                dynamic item = document;
                if (!saveAs) item.Save();
                else if (HostKind == "Word") item.SaveAs2(destination, format);
                else item.SaveAs(destination, format);
            }
            /// <summary>Existence du fichier natif.</summary>
            public bool FileExists(string path) => File.Exists(path);
            /// <summary>Existence du dossier natif.</summary>
            public bool DirectoryExists(string path) => Directory.Exists(path);
            /// <summary>Taille native sans lire le contenu.</summary>
            public long FileLength(string path) => new FileInfo(path).Length;
        }

        /// <summary>Reconnaît les deux noms de processus; Excel et les autres hôtes gardent leurs adaptateurs dédiés.</summary>
        internal static string RecognizeOtherHost(string processName) => string.Equals(processName, "WINWORD", StringComparison.OrdinalIgnoreCase) ? "Word" :
            string.Equals(processName, "POWERPNT", StringComparison.OrdinalIgnoreCase) ? "PowerPoint" : null;
        /// <summary>Indique seulement qu'un adaptateur existe; ne constitue pas une qualification native.</summary>
        internal bool SupportsOtherHost => new NativeOtherHostProbe().HostKind != null;
        /// <summary>Lit l'état du document associé au projet par son identité COM.</summary>
        internal object OtherHostPersistence(string selector) => OtherHostPersistence(selector, new NativeOtherHostProbe());
        /// <summary>Lecture injectable sans mutation ni lancement d'application.</summary>
        internal object OtherHostPersistence(string selector, IOtherHostProbe native)
        {
            dynamic project = GetProject(selector);
            try
            {
                object document = MatchOtherHost((object)project, native); var state = native.State(document);
                return new { Project = selector, ProjectSaved = (bool)project.Saved, Host = native.HostKind,
                    HostAvailable = true, HostPath = state.Path, HostSaved = (bool?)state.Saved, HostReadOnly = (bool?)state.ReadOnly,
                    HostHasPath = (bool?)!string.IsNullOrWhiteSpace(state.Path), NativeFileFormat = state.Format,
                    OwnerProcessId = native.CurrentProcessId, IdentityVerified = true, NativeQualification = "NOT_RUN", Reason = (string)null };
            }
            catch (Exception ex)
            {
                return new { Project = selector, ProjectSaved = (bool)project.Saved, HostAvailable = false,
                    HostPath = (string)null, HostSaved = (bool?)null, HostReadOnly = (bool?)null, HostHasPath = (bool?)null,
                    NativeQualification = "NOT_RUN", Reason = ex.Message };
            }
        }

        /// <summary>Sauvegarde Word/PowerPoint après gardes de version, chemin, identité COM et processus.</summary>
        internal object SaveOtherHost(Request request, bool saveAs) => SaveOtherHost(request, saveAs, new NativeOtherHostProbe());
        /// <summary>Sauvegarde injectable; une exception après invocation produit explicitement un résultat incertain.</summary>
        internal object SaveOtherHost(Request request, bool saveAs, IOtherHostProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion)) throw new ArgumentException("ExpectedProjectVersion is required.");
            dynamic project = GetDesignProject(request.Project); AssertProjectVersion(request, project);
            if ((int)project.Protection != 0) throw new InvalidOperationException("The project is protected.");
            object document = MatchOtherHost((object)project, native); var before = native.State(document);
            if (before.ReadOnly) throw new InvalidOperationException("The matched host document is read-only.");
            string path; int format;
            if (saveAs)
            {
                if (!string.IsNullOrEmpty(before.Path) || !string.IsNullOrEmpty(request.ExpectedHostPath) || !string.IsNullOrEmpty(OtherHostProjectPath((object)project)))
                    throw new InvalidOperationException("SaveAs supports only the first save of an unsaved document and project.");
                path = RequireAbsolutePath(request.Path); format = OtherHostFormat(native.HostKind, path);
                if (native.FileExists(path) || native.DirectoryExists(path)) throw new IOException("The SaveAs destination already exists.");
                if (!native.DirectoryExists(Path.GetDirectoryName(path))) throw new DirectoryNotFoundException("The SaveAs parent directory is absent.");
            }
            else
            {
                path = RequireAbsolutePath(request.ExpectedHostPath); format = OtherHostFormat(native.HostKind, path);
                if (!OtherHostSamePath(before.Path, path) || !OtherHostSamePath(OtherHostProjectPath((object)project), path))
                    throw new InvalidOperationException("The document or selected project's path changed since inspection.");
                if (!native.FileExists(path)) throw new FileNotFoundException("The existing host document is absent.");
                if (before.Format.HasValue && before.Format.Value != format) throw new InvalidOperationException("The native Word format does not match its macro-enabled extension.");
            }
            string sourceSha = OtherHostSourceSha((object)project);
            AssertProjectVersion(request, project);
            if (native.ApplicationProcessId(native.Application()) != (uint)native.CurrentProcessId || !native.SameProject((object)project, native.DocumentProject(document)))
                throw new InvalidOperationException("Host process or document/project identity changed before saving.");
            var preflight = native.State(document);
            if (preflight.ReadOnly || !string.Equals(preflight.Path, before.Path, StringComparison.OrdinalIgnoreCase) || preflight.Format != before.Format ||
                OtherHostSourceSha((object)project) != sourceSha)
                throw new InvalidOperationException("Host path, format, read-only state or live source changed during save preparation.");
            try
            {
                native.Save(document, saveAs, path, format);
                var after = native.State(document);
                if (!OtherHostSamePath(after.Path, path) || !OtherHostSamePath(OtherHostProjectPath((object)project), path) ||
                    !after.Saved || !(bool)project.Saved || !native.FileExists(path) || native.FileLength(path) < 1 ||
                    (after.Format.HasValue && after.Format.Value != format) || OtherHostSourceSha((object)project) != sourceSha ||
                    !native.SameProject((object)project, native.DocumentProject(document)))
                    throw new InvalidOperationException("Saved flags, paths, file bytes, format, identity or live VBA source did not match after save.");
                return new { Project = request.Project, Host = native.HostKind, HostPath = path, SaveAsInvoked = saveAs, SaveInvoked = !saveAs,
                    Verified = true, Uncertain = false, MutationInvoked = true, HostSaved = true, ProjectSaved = true, Bytes = native.FileLength(path),
                    SourceSha256 = sourceSha, CodePreserved = true, NativeFileFormatVerified = after.Format.HasValue,
                    NativeQualification = "NOT_RUN", PersistenceReopenVerified = false,
                    Limit = "Live code, host state and file presence were verified. Reopen the native file to verify code persistence; real Word/PowerPoint host qualification remains NOT_RUN." };
            }
            catch (Exception ex)
            {
                return new { Project = request.Project, Host = native.HostKind, HostPath = path, MutationInvoked = true,
                    SaveAsInvoked = saveAs, SaveInvoked = !saveAs, Verified = false, Uncertain = true, NativeQualification = "NOT_RUN",
                    Reason = ex.Message, Next = "Read project_persistence_status and inspect the file; do not retry automatically." };
            }
        }

        /// <summary>Associe un seul document au projet IUnknown et refuse toute application d'un autre PID.</summary>
        private static object MatchOtherHost(object project, IOtherHostProbe native)
        {
            if (native == null || (native.HostKind != "Word" && native.HostKind != "PowerPoint")) throw new InvalidOperationException("This host has no Word/PowerPoint adapter.");
            object application = native.Application();
            if (native.CurrentProcessId <= 0 || native.ApplicationProcessId(application) != (uint)native.CurrentProcessId)
                throw new InvalidOperationException("The application does not belong to this add-in process.");
            var documents = native.Documents(application);
            if (documents == null || documents.Count > 1000) throw new InvalidOperationException("The document collection is unreadable or oversized.");
            var matches = documents.Where(document => native.SameProject(project, native.DocumentProject(document))).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("No unique open document shares the selected VBProject COM identity.");
            return matches[0];
        }
        /// <summary>Associe uniquement les extensions macro explicites aux constantes Microsoft de SaveAs.</summary>
        internal static int OtherHostFormat(string kind, string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (kind == "Word") { if (extension == ".docm") return 13; if (extension == ".dotm") return 15; }
            if (kind == "PowerPoint") { if (extension == ".pptm") return 25; if (extension == ".potm") return 27; if (extension == ".ppsm") return 29; }
            throw new ArgumentException("Word requires .docm/.dotm; PowerPoint requires .pptm/.potm/.ppsm. Other formats are refused.");
        }
        /// <summary>Lit FileName; son absence n'est acceptée que dans les gardes de première sauvegarde.</summary>
        private static string OtherHostProjectPath(object project)
        { try { return (string)((dynamic)project).FileName; } catch (COMException) { return ""; } }
        /// <summary>Compare deux chemins pleinement qualifiés, sans accepter un nom de fichier relatif.</summary>
        private static bool OtherHostSamePath(string first, string second) => !string.IsNullOrWhiteSpace(first) && Path.IsPathRooted(first) &&
            string.Equals(Path.GetFullPath(first), second, StringComparison.OrdinalIgnoreCase);
        /// <summary>Versionne toutes les sources live et identités de composants avant/après sauvegarde.</summary>
        private static string OtherHostSourceSha(object project)
        {
            var rows = new List<object>();
            foreach (dynamic component in ((dynamic)project).VBComponents)
            {
                dynamic code = component.CodeModule; int count = (int)code.CountOfLines;
                rows.Add(new { Name = (string)component.Name, Type = (int)component.Type,
                    SourceSha = Hash(count == 0 ? "" : (string)code.Lines(1, count)) });
            }
            return Hash(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(rows));
        }
    }
}
