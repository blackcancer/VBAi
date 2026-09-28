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
    /// <summary>Adapte la lecture et la sauvegarde sécurisée de composants VBA aux hôtes Office reconnus.</summary>
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
            /// <value>Nom canonique Word/PowerPoint ou null si le processus courant n’est pas pris en charge.</value>
            string HostKind { get; }
                        /// <summary>PID du processus hébergeant l'add-in.</summary>
            /// <value>Identifiant Windows du processus courant.</value>
            int CurrentProcessId { get; }
                        /// <summary>Obtient une application déjà ouverte, sans lancement d'hôte.</summary>
            /// <returns>Objet Application qui a été vérifié comme appartenant au processus courant.</returns>
            object Application();
                        /// <summary>Vérifie les handles natifs de cette application et retourne leur PID commun.</summary>
            /// <param name="application">Application Office à vérifier.</param>
            /// <returns>PID commun, ou zéro si les fenêtres ne peuvent pas être attribuées à un seul processus.</returns>
            uint ApplicationProcessId(object application);
                        /// <summary>Énumère tous les documents ouverts de l'application vérifiée.</summary>
            /// <param name="application">Application dont les documents seront parcourus.</param>
            /// <returns>Documents ouverts sous forme d’objets COM.</returns>
            IList<object> Documents(object application);
                        /// <summary>Lit le VBProject du document, sans supposer un nom unique.</summary>
            /// <param name="document">Document Word ou présentation PowerPoint.</param>
            /// <returns>Projet VBA attaché au document.</returns>
            object DocumentProject(object document);
                        /// <summary>Compare les identités IUnknown COM des deux projets.</summary>
            /// <param name="first">Premier projet à comparer.</param>
            /// <param name="second">Second projet à comparer.</param>
            /// <returns><see langword="true"/> si les projets partagent la même identité COM.</returns>
            bool SameProject(object first, object second);
                        /// <summary>Lit chemin, Saved, lecture seule et format natif disponible.</summary>
            /// <param name="document">Document à inspecter.</param>
            /// <returns>État natif du document à l’instant de lecture.</returns>
            OtherHostDocumentState State(object document);
                        /// <summary>Invoque Save, Word.SaveAs2 ou PowerPoint.SaveAs exactement une fois.</summary>
            /// <param name="document">Document dont l’API de sauvegarde doit être appelée.</param>
            /// <param name="saveAs">Sélectionne l’opération SaveAs lorsqu’il vaut true.</param>
            /// <param name="destination">Chemin de destination de SaveAs.</param>
            /// <param name="format">Format natif transmis à l’hôte.</param>
            void Save(object document, bool saveAs, string destination, int format);
                        /// <summary>Indique si un fichier de destination existe.</summary>
            /// <param name="path">Chemin à tester.</param>
            /// <returns><see langword="true"/> si le fichier existe.</returns>
            bool FileExists(string path);
                        /// <summary>Indique si le dossier parent existe.</summary>
            /// <param name="path">Chemin du dossier.</param>
            /// <returns><see langword="true"/> si le dossier existe.</returns>
            bool DirectoryExists(string path);
                        /// <summary>Lit la taille du fichier après invocation native.</summary>
            /// <param name="path">Fichier dont la taille doit être lue.</param>
            /// <returns>Taille en octets.</returns>
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
            /// <summary>Performs the current host kind operation for NativeOtherHostProbe.</summary>
/// <returns>The result produced by this operation.</returns>
            private static string CurrentHostKind() => RecognizeOtherHost(Process.GetCurrentProcess().ProcessName);
                        /// <summary>Reconnaît exclusivement WINWORD.EXE ou POWERPNT.EXE.</summary>
            /// <value>Nom Word ou PowerPoint, ou null pour tout autre processus.</value>
            public string HostKind => ReadHostKind();
                        /// <summary>PID du processus de l'add-in.</summary>
            /// <value>Identifiant du processus courant.</value>
            public int CurrentProcessId => Process.GetCurrentProcess().Id;
                        /// <summary>Résout le ROT, puis NativeOM dans un document appartenant au PID courant si nécessaire.</summary>
            /// <returns>Application Office existante vérifiée comme appartenant au processus courant.</returns>
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
            /// <param name="application">Application Office à inspecter.</param>
            /// <returns>PID propriétaire commun, ou zéro lorsque la vérification échoue.</returns>
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
            /// <param name="window">Handle natif à vérifier.</param>
            /// <returns>PID propriétaire de la fenêtre.</returns>
            private static uint Owner(IntPtr window)
            { uint owner; GetWindowThreadProcessId(window, out owner); return owner; }
                        /// <summary>Énumère les documents sans ignorer une erreur d'accès au catalogue.</summary>
            /// <param name="application">Application Office vérifiée.</param>
            /// <returns>Documents ou présentations ouverts, dans leur ordre natif.</returns>
            public IList<object> Documents(object application)
            {
                var result = new List<object>();
                System.Collections.IEnumerable documents = HostKind == "Word" ? (System.Collections.IEnumerable)((dynamic)application).Documents : (System.Collections.IEnumerable)((dynamic)application).Presentations;
                foreach (object document in documents)
                { if (result.Count >= 1000) throw new InvalidOperationException("Unexpected Office document count."); result.Add(document); }
                return result;
            }
                        /// <summary>Lit l'objet projet attaché au document.</summary>
            /// <param name="document">Document hôte à inspecter.</param>
            /// <returns>VBProject associé.</returns>
            public object DocumentProject(object document) => ((dynamic)document).VBProject;
                        /// <summary>Compare IUnknown et libère seulement les références acquises par cette comparaison.</summary>
            /// <param name="first">Première référence de projet.</param>
            /// <param name="second">Seconde référence de projet.</param>
            /// <returns><see langword="true"/> si les deux références désignent le même projet COM.</returns>
            public bool SameProject(object first, object second)
            {
                if (first == null || second == null || !Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
                IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
                try { a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (a != IntPtr.Zero) Marshal.Release(a); if (b != IntPtr.Zero) Marshal.Release(b); }
            }
                        /// <summary>Lit les états natifs; aucune propriété FileFormat inexistante n'est inventée pour PowerPoint.</summary>
            /// <param name="document">Document hôte observé.</param>
            /// <returns>Chemin, indicateurs natifs et format lorsqu’il est disponible.</returns>
            public OtherHostDocumentState State(object document)
            {
                dynamic item = document; bool hasPath = !string.IsNullOrWhiteSpace((string)item.Path);
                return new OtherHostDocumentState { Path = hasPath ? (string)item.FullName : "",
                    ReadOnly = HostKind == "Word" ? (bool)item.ReadOnly : Convert.ToInt32(item.ReadOnly) != 0,
                    Saved = HostKind == "Word" ? (bool)item.Saved : Convert.ToInt32(item.Saved) == -1,
                    Format = HostKind == "Word" ? (int?)Convert.ToInt32(item.SaveFormat) : null };
            }
                        /// <summary>Appelle la méthode native appropriée sans fermer, imprimer ou démarrer un document.</summary>
            /// <param name="document">Document qui sera sauvegardé.</param>
            /// <param name="saveAs">Indique si l’opération utilise SaveAs.</param>
            /// <param name="destination">Chemin demandé pour SaveAs.</param>
            /// <param name="format">Format macro transmis à Word ou PowerPoint.</param>
            public void Save(object document, bool saveAs, string destination, int format)
            {
                dynamic item = document;
                if (!saveAs) item.Save();
                else if (HostKind == "Word") item.SaveAs2(destination, format);
                else item.SaveAs(destination, format);
            }
                        /// <summary>Existence du fichier natif.</summary>
            /// <param name="path">Chemin du fichier.</param>
            /// <returns><see langword="true"/> si le chemin désigne un fichier existant.</returns>
            public bool FileExists(string path) => File.Exists(path);
                        /// <summary>Existence du dossier natif.</summary>
            /// <param name="path">Chemin du dossier.</param>
            /// <returns><see langword="true"/> si le chemin désigne un dossier existant.</returns>
            public bool DirectoryExists(string path) => Directory.Exists(path);
                        /// <summary>Taille native sans lire le contenu.</summary>
            /// <param name="path">Chemin du fichier sauvegardé.</param>
            /// <returns>Taille du fichier en octets.</returns>
            public long FileLength(string path) => new FileInfo(path).Length;
        }

                /// <summary>Reconnaît les deux noms de processus; Excel et les autres hôtes gardent leurs adaptateurs dédiés.</summary>
        /// <param name="processName">Nom de processus sans extension.</param>
        /// <returns>Word, PowerPoint ou null si le nom ne correspond pas à ces hôtes.</returns>
        internal static string RecognizeOtherHost(string processName) => string.Equals(processName, "WINWORD", StringComparison.OrdinalIgnoreCase) ? "Word" :
            string.Equals(processName, "POWERPNT", StringComparison.OrdinalIgnoreCase) ? "PowerPoint" : null;
                /// <summary>Indique seulement qu'un adaptateur existe; ne constitue pas une qualification native.</summary>
        /// <value><see langword="true"/> lorsque le processus courant est Word ou PowerPoint.</value>
        internal bool SupportsOtherHost => new NativeOtherHostProbe().HostKind != null;
                /// <summary>Lit l'état du document associé au projet par son identité COM.</summary>
        /// <param name="selector">Sélecteur du projet VBA.</param>
        /// <returns>État de persistance exposé à l’appelant; les échecs sont renvoyés comme indisponibilité.</returns>
        internal object OtherHostPersistence(string selector) => OtherHostPersistence(selector, new NativeOtherHostProbe());
                /// <summary>Lecture injectable sans mutation ni lancement d'application.</summary>
        /// <param name="selector">Sélecteur du projet VBA.</param>
        /// <param name="native">Sonde utilisée pour l’application hôte, les documents et les états.</param>
        /// <returns>État de persistance natif ou résultat décrivant pourquoi il n’est pas disponible.</returns>
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
        /// <param name="request">Projet, version attendue et chemin éventuel fournis par l’appelant.</param>
        /// <param name="saveAs">Choisit une première sauvegarde sous un nouveau chemin.</param>
        /// <returns>Résultat de la sauvegarde native et état des vérifications.</returns>
        internal object SaveOtherHost(Request request, bool saveAs) => SaveOtherHost(request, saveAs, new NativeOtherHostProbe());
                /// <summary>Sauvegarde injectable; une exception après invocation produit explicitement un résultat incertain.</summary>
        /// <param name="request">Projet et préconditions de sauvegarde.</param>
        /// <param name="saveAs">Indique si une opération SaveAs est demandée.</param>
        /// <param name="native">Sonde des opérations hôte et du document associé.</param>
        /// <returns>Résultat vérifié ou résultat incertain si une mutation a été appelée mais non confirmée.</returns>
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
        /// <param name="project">VBProject sélectionné par l’appelant.</param>
        /// <param name="native">Sonde qui fournit l’application, ses documents et leurs identités.</param>
        /// <returns>Document unique dont le VBProject partage l’identité COM avec le projet demandé.</returns>
        /// <exception cref="InvalidOperationException">L’hôte, le processus ou l’identité du document ne peut pas être vérifié de façon unique.</exception>
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
        /// <param name="kind">Word ou PowerPoint.</param>
        /// <param name="path">Chemin dont l’extension détermine le format.</param>
        /// <returns>Constante SaveAs native associée au format macro-enabled.</returns>
        /// <exception cref="ArgumentException">L’hôte ou l’extension n’est pas prise en charge.</exception>
        internal static int OtherHostFormat(string kind, string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (kind == "Word") { if (extension == ".docm") return 13; if (extension == ".dotm") return 15; }
            if (kind == "PowerPoint") { if (extension == ".pptm") return 25; if (extension == ".potm") return 27; if (extension == ".ppsm") return 29; }
            throw new ArgumentException("Word requires .docm/.dotm; PowerPoint requires .pptm/.potm/.ppsm. Other formats are refused.");
        }
                /// <summary>Lit FileName; son absence n'est acceptée que dans les gardes de première sauvegarde.</summary>
        /// <param name="project">Projet VBA du document.</param>
        /// <returns>Chemin du projet ou chaîne vide si l’hôte ne fournit pas FileName.</returns>
        private static string OtherHostProjectPath(object project)
        { try { return (string)((dynamic)project).FileName; } catch (COMException) { return ""; } }
                /// <summary>Compare deux chemins pleinement qualifiés, sans accepter un nom de fichier relatif.</summary>
        /// <param name="first">Premier chemin à comparer.</param>
        /// <param name="second">Second chemin attendu.</param>
        /// <returns><see langword="true"/> si le premier est absolu et correspond au second sans tenir compte de la casse.</returns>
        private static bool OtherHostSamePath(string first, string second) => !string.IsNullOrWhiteSpace(first) && Path.IsPathRooted(first) &&
            string.Equals(Path.GetFullPath(first), second, StringComparison.OrdinalIgnoreCase);
                /// <summary>Versionne toutes les sources live et identités de composants avant/après sauvegarde.</summary>
        /// <param name="project">VBProject dont les composants doivent être empreintés.</param>
        /// <returns>Empreinte SHA-256 des noms, types et sources des composants.</returns>
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
