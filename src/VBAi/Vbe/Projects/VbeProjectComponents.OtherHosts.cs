using Microsoft.CSharp.RuntimeBinder;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Adapte la lecture et la sauvegarde sécurisée de composants VBA aux hôtes Office reconnus.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Reads the identity-matched Office document state; Access has no document Saved flag.</summary>
        internal sealed class OtherHostDocumentState
        {

            /// <summary>Chemin complet, ou chaîne vide avant première sauvegarde.</summary>
            internal string Path;

            /// <summary>Native document Saved state, or null when the host exposes no document flag.</summary>
            internal bool? Saved;

            /// <summary>État natif de lecture seule; un état indéterminé est traité comme lecture seule.</summary>
            internal bool ReadOnly;

            /// <summary>Native Word, Access or Publisher format; null for PowerPoint.</summary>
            internal int? Format;
        }

        /// <summary>Frontière injectable d'identité, état et sauvegarde du seul processus hôte courant.</summary>
        internal interface IOtherHostProbe
        {

            /// <summary>Word, PowerPoint, Access, Publisher, or null for an unrecognized process.</summary>
            /// <value>Canonical Office host name, or null when the current process has no adapter.</value>
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

            /// <summary>Resolves the project by COM identity or its unique path within the bound VBE.</summary>
            /// <param name="document">Native Office document or Access CurrentProject.</param>
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

            /// <summary>Invokes one native document save or the guarded Access VBE Save command.</summary>
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

        /// <summary>Native probe restricted to the current Office process; each operation requires native qualification.</summary>
        internal sealed class NativeOtherHostProbe : IOtherHostProbe
        {

            /// <summary>VBE editor, canonical Access/Publisher project and document, and prepared Access Save control.</summary>
            private object editor, boundProject, boundDocument, accessSaveControl;

            /// <summary>Exact code pane and component approved for the one guarded Access save.</summary>
            private object accessExpectedPane, accessExpectedComponent;

            /// <summary>COM identity comparer used to re-resolve projects and documents without name-only matching.</summary>
            internal Func<object, object, bool> ReadIdentity = SameComIdentity;

            /// <summary>Gets or sets the save invocation started.</summary>
            /// <value>Current save invocation started exposed by native other host probe.</value>
            internal bool SaveInvocationStarted { get; private set; }
            // Bound only by the original owner-thread Access async save; never approves an existing prompt.
            /// <summary>Final Access authority/revision callback invoked before the native Save command.</summary>
            internal Action AccessBeforeSave;

            /// <summary>Binds the actual selected VBIDE project, never a host-specific invented VBProject property.</summary>
            /// <param name="project">Selected Access or Publisher VBProject whose VBE must belong to the current process.</param>
            internal void BindProject(object project)
            {
                if (HostKind != "Access" && HostKind != "Publisher") return;
                editor = ((dynamic)project).VBE;
                if (editor == null || ReadOwner(new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd))) != (uint)CurrentProcessId)
                    throw new InvalidOperationException("The selected project's VBE does not belong to this add-in process.");
                boundProject = project;
                boundDocument = null;
            }

            /// <summary>Retains the exact document identity found during project association.</summary>
            /// <param name="document">Exact host document matched to the selected project during association.</param>
            internal void BindDocument(object document)
            {
                if (HostKind == "Access" || HostKind == "Publisher") boundDocument = document;
            }

            /// <summary>Retains the exact async-approved Access selection for final pre-command revalidation.</summary>
            /// <param name="pane">Active Access code pane approved for the save operation.</param>
            /// <param name="component">Component identity that the approved pane must continue to display.</param>
            internal void BindAccessSaveSelection(object pane, object component)
            {
                if (HostKind != "Access" || pane == null || component == null)
                    throw new InvalidOperationException("An exact Access code selection is required.");
                accessExpectedPane = pane; accessExpectedComponent = component;
            }

            /// <summary>Prepares the existing native Save command without invoking it or compiling VBA.</summary>
            /// <param name="document">Matched Access document whose project must equal the bound VBE project.</param>
            internal void PrepareSave(object document)
            {
                if (HostKind != "Access") return;
                var commands = new NativeSolidWorksSaveProbe();
                accessSaveControl = commands.SaveControl(editor);
                if (Convert.ToInt32(((dynamic)accessSaveControl).Type) != 1 ||
                    !SameProject(boundProject, DocumentProject(document)) ||
                    !SameProject(boundProject, (object)((dynamic)editor).ActiveVBProject))
                    throw new InvalidOperationException("The built-in VBE Save command must target the matched Access project.");
                if (accessExpectedPane != null)
                {
                    object currentPane = ((dynamic)editor).ActiveCodePane;
                    if (currentPane == null || !SameProject(currentPane, accessExpectedPane) ||
                        !SameProject((object)((dynamic)currentPane).CodeModule.Parent, accessExpectedComponent))
                        throw new InvalidOperationException("The approved Access code selection changed before Save.");
                }
            }

            /// <summary>Reads the real process kind by default; isolates host contracts during qualification.</summary>
            internal Func<string> ReadHostKind = CurrentHostKind;

            /// <summary>Resolves an existing ROT application without starting Office.</summary>
            internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;

            /// <summary>Reads the running executable's Office major version without choosing an installed instance.</summary>
            internal Func<int> ReadHostMajorVersion = () =>
            {
                using (var process = Process.GetCurrentProcess())
                    return FileVersionInfo.GetVersionInfo(process.MainModule.FileName).FileMajorPart;
            };

            /// <summary>Reads the native window owner used by all application PID guards.</summary>
            internal Func<IntPtr, uint> ReadOwner = Owner;

            /// <summary>Reads the existing PowerPoint window HWND for current-process ownership checks.</summary>
            internal Func<object, IntPtr> ReadPowerPointWindow = PowerPointWindow.Read;

            /// <summary>Classifies the current executable without attaching to or launching another host.</summary>
            /// <returns>Canonical host name for a recognized Office process, or null otherwise.</returns>
            private static string CurrentHostKind() => RecognizeOtherHost(Process.GetCurrentProcess().ProcessName);

            /// <summary>Recognizes WINWORD, POWERPNT, MSACCESS and MSPUB process names.</summary>
            /// <value>Word, PowerPoint, Access, Publisher, or null for another process.</value>
            public string HostKind => ReadHostKind();

            /// <summary>PID du processus de l'add-in.</summary>
            /// <value>Identifiant du processus courant.</value>
            public int CurrentProcessId => Process.GetCurrentProcess().Id;

            /// <summary>Résout le ROT, puis NativeOM dans un document appartenant au PID courant si nécessaire.</summary>
            /// <returns>Application Office existante vérifiée comme appartenant au processus courant.</returns>
            public object Application()
            {
                if (HostKind == null) throw new InvalidOperationException("Only the current Word, PowerPoint, Access or Publisher process is supported.");
                string rotObservation = "No verified ROT entry.";
                var progIds = new List<string> { HostKind + ".Application" };
                if (HostKind == "Access")
                {
                    int major = ReadHostMajorVersion();
                    if (major > 0 && major <= 100) progIds.Add("Access.Application." + major);
                }
                foreach (string progId in progIds)
                    try
                    {
                        object registered = ReadActiveApplication(progId);
                        uint owner = ApplicationProcessId(registered);
                        if (owner == (uint)CurrentProcessId) return registered;
                        rotObservation = "ROT entry " + progId + " has owner PID " + owner + "; expected " + CurrentProcessId + ".";
                    }
                    catch (COMException error) { rotObservation = "ROT lookup " + progId + " failed with HRESULT 0x" + unchecked((uint)error.ErrorCode).ToString("X8") + "."; }
                    catch (RuntimeBinderException) { rotObservation = "ROT entry " + progId + " did not expose the required native owner window."; }
                if (HostKind == "Access" || HostKind == "Publisher")
                    throw new InvalidOperationException("No running " + HostKind + " application was verified as belonging to this VBE PID. " + rotObservation);
                var windows = new List<IntPtr>();
                VbeDebugWindows.EnumWindows((parent, parameter) =>
                {
                    VbeDebugWindows.GetWindowThreadProcessId(parent, out uint owner);
                    if (owner != (uint)CurrentProcessId) return true;
                    VbeDebugWindows.EnumChildWindows(parent, (child, childParameter) =>
                    {
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
                        Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                        if (VbeDebugWindows.AccessibleObjectFromWindow(window, 0xFFFFFFF0, ref dispatch, out object native) != 0 || native == null) continue;
                        object application = ((dynamic)native).Application;
                        if (ApplicationProcessId(application) == (uint)CurrentProcessId) return application;
                    }
                    catch (COMException) { }
                    catch (RuntimeBinderException) { }
                }
                throw new InvalidOperationException("No running Office application was verified as belonging to this VBE PID.");
            }

            /// <summary>Verifies the native application window owner without selecting a document.</summary>
            /// <param name="application">Application Office à inspecter.</param>
            /// <returns>PID propriétaire commun, ou zéro lorsque la vérification échoue.</returns>
            public uint ApplicationProcessId(object application)
            {
                if (HostKind == "PowerPoint") return ReadOwner(ReadPowerPointWindow(application));
                if (HostKind == "Access") return ReadOwner(new IntPtr(Convert.ToInt64(((dynamic)application).hWndAccessApp())));
                if (HostKind == "Publisher") return ReadOwner(new IntPtr(Convert.ToInt64(((dynamic)application).ActiveWindow.hWnd)));
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
            { GetWindowThreadProcessId(window, out uint owner); return owner; }

            /// <summary>Énumère les documents sans ignorer une erreur d'accès au catalogue.</summary>
            /// <param name="application">Application Office vérifiée.</param>
            /// <returns>Documents ou présentations ouverts, dans leur ordre natif.</returns>
            public IList<object> Documents(object application)
            {
                var result = new List<object>();
                if (HostKind == "Access") { result.Add((object)((dynamic)application).CurrentProject); return result; }
                System.Collections.IEnumerable documents = HostKind == "PowerPoint" ? (System.Collections.IEnumerable)((dynamic)application).Presentations : (System.Collections.IEnumerable)((dynamic)application).Documents;
                foreach (object document in documents)
                { if (result.Count >= 1000) throw new InvalidOperationException("Unexpected Office document count."); result.Add(document); }
                return result;
            }

            /// <summary>Lit l'objet projet attaché au document.</summary>
            /// <param name="document">Document hôte à inspecter.</param>
            /// <returns>VBProject associé.</returns>
            public object DocumentProject(object document)
            {
                if (HostKind != "Access" && HostKind != "Publisher") return ((dynamic)document).VBProject;
                if (editor == null) throw new InvalidOperationException("The selected VBIDE project must be bound first.");
                string path = State(document).Path;
                if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
                    throw new InvalidOperationException("Access/Publisher require an existing absolute document path for project association.");
                if (HostKind == "Publisher" && string.IsNullOrWhiteSpace(OtherHostProjectPath(boundProject)))
                    return SolePublisherProject(document);
                object match = null; int count = 0;
                foreach (object project in ((dynamic)editor).VBProjects)
                {
                    if (++count > 1000) throw new InvalidOperationException("Unexpected VBIDE project count.");
                    if (!OtherHostSamePath(OtherHostProjectPath(project), Path.GetFullPath(path))) continue;
                    if (match != null) throw new InvalidOperationException("More than one VBIDE project has the matched document path.");
                    if (!SameProject((object)((dynamic)project).VBE, editor))
                        throw new InvalidOperationException("The matched project belongs to another VBE.");
                    match = project;
                }
                if (match == null || !SameProject(match, boundProject))
                    throw new InvalidOperationException("The document path does not identify the selected VBIDE project.");
                return match;
            }

            /// <summary>Associates a pathless Publisher project only within one exact native document/VBE pair.</summary>
            /// <param name="document">Exact open Publisher document associated with the selected project.</param>
            /// <returns>The single VBIDE project only when one document and one project share the bound COM identities.</returns>
            private object SolePublisherProject(object document)
            {
                object application = Application();
                if (ApplicationProcessId(application) != (uint)CurrentProcessId ||
                    ReadOwner(new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd))) != (uint)CurrentProcessId)
                    throw new InvalidOperationException("The Publisher application or bound VBE changed process.");
                var documents = Documents(application);
                if (documents.Count != 1 || !SameProject(documents[0], document))
                    throw new InvalidOperationException("A pathless Publisher project requires exactly one identified open document.");
                object sole = null;
                foreach (object project in ((dynamic)editor).VBProjects)
                {
                    if (sole != null) throw new InvalidOperationException("A pathless Publisher project requires exactly one VBIDE project.");
                    sole = project;
                }
                if (sole == null || !SameProject(sole, boundProject) ||
                    !SameProject((object)((dynamic)sole).VBE, editor))
                    throw new InvalidOperationException("The sole Publisher project is not the selected project's COM identity and VBE.");
                return sole;
            }

            /// <summary>Compare IUnknown et libère seulement les références acquises par cette comparaison.</summary>
            /// <param name="first">Première référence de projet.</param>
            /// <param name="second">Seconde référence de projet.</param>
            /// <returns><see langword="true"/> si les deux références désignent le même projet COM.</returns>
            public bool SameProject(object first, object second) => ReadIdentity(first, second);

            /// <summary>Compares two COM objects by IUnknown identity and releases only the acquired pointers.</summary>
            /// <param name="first">First candidate COM object; non-COM or null inputs return false.</param>
            /// <param name="second">Second candidate COM object; non-COM or null inputs return false.</param>
            /// <returns>True when both objects expose the same canonical IUnknown pointer.</returns>
            private static bool SameComIdentity(object first, object second)
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
                if (HostKind == "Access")
                {
                    dynamic current = document; dynamic database = current.Application.CurrentDb();
                    try
                    {
                        return new OtherHostDocumentState
                        {
                            Path = (string)current.FullName,
                            Saved = null,
                            ReadOnly = !(bool)database.Updatable,
                            Format = Convert.ToInt32(current.FileFormat)
                        };
                    }
                    finally { if (database != null && Marshal.IsComObject(database)) Marshal.ReleaseComObject(database); }
                }
                dynamic item = document; bool hasPath = !string.IsNullOrWhiteSpace((string)item.Path);
                return new OtherHostDocumentState
                {
                    Path = hasPath ? (string)item.FullName : "",
                    ReadOnly = HostKind == "PowerPoint" ? Convert.ToInt32(item.ReadOnly) != 0 : (bool)item.ReadOnly,
                    Saved = HostKind == "PowerPoint" ? Convert.ToInt32(item.Saved) == -1 : (bool)item.Saved,
                    Format = HostKind == "PowerPoint" ? null : (int?)Convert.ToInt32(item.SaveFormat)
                };
            }

            /// <summary>Appelle la méthode native appropriée sans fermer, imprimer ou démarrer un document.</summary>
            /// <param name="document">Document qui sera sauvegardé.</param>
            /// <param name="saveAs">Indique si l’opération utilise SaveAs.</param>
            /// <param name="destination">Chemin demandé pour SaveAs.</param>
            /// <param name="format">Verified native Office file format.</param>
            public void Save(object document, bool saveAs, string destination, int format)
            {
                SaveInvocationStarted = false;
                dynamic item = document;
                if (HostKind == "Access")
                {
                    if (saveAs || accessSaveControl == null) throw new InvalidOperationException("Access requires the prepared existing-document VBE Save command.");
                    PrepareSave(document);
                    if (ApplicationProcessId(Application()) != (uint)CurrentProcessId || State(document).ReadOnly ||
                        !OtherHostSamePath(State(document).Path, destination) || !(bool)((dynamic)accessSaveControl).Enabled)
                        throw new InvalidOperationException("Access save identity or writable state changed before invocation.");
                    PrepareSave(document);
                    AccessBeforeSave?.Invoke();
                    SaveInvocationStarted = true;
                    new NativeSolidWorksSaveProbe().Save(accessSaveControl);
                    return;
                }
                if (HostKind == "Publisher")
                {
                    if (saveAs || boundDocument == null || !SameProject(document, boundDocument))
                        throw new InvalidOperationException("Publisher requires the exact previously matched existing document.");
                    object application = Application();
                    if (ApplicationProcessId(application) != (uint)CurrentProcessId ||
                        Documents(application).Count(itemDocument => SameProject(itemDocument, boundDocument)) != 1 ||
                        !SameProject(boundProject, DocumentProject(document)))
                        throw new InvalidOperationException("Publisher process, document or project identity changed before saving.");
                    var state = State(document);
                    if (state.ReadOnly || !OtherHostSamePath(state.Path, destination) || state.Format != format)
                        throw new InvalidOperationException("Publisher path, format or writable state changed before invocation.");
                }
                SaveInvocationStarted = true;
                if (!saveAs) item.Save();
                else if (HostKind == "Word") item.SaveAs2(destination, format);
                else if (HostKind == "Publisher") item.SaveAs(destination, format, false);
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

        /// <summary>Recognizes the four Office processes using this adapter.</summary>
        /// <param name="processName">Nom de processus sans extension.</param>
        /// <returns>Word, PowerPoint, Access, Publisher, or null for another process.</returns>
        internal static string RecognizeOtherHost(string processName) => string.Equals(processName, "WINWORD", StringComparison.OrdinalIgnoreCase) ? "Word" :
            string.Equals(processName, "POWERPNT", StringComparison.OrdinalIgnoreCase) ? "PowerPoint" :
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ? "Access" :
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase) ? "Publisher" : null;

        /// <summary>Indique seulement qu'un adaptateur existe; ne constitue pas une qualification native.</summary>
        /// <value><see langword="true"/> when the current process has an Office document adapter.</value>
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
                return new
                {
                    Project = selector,
                    ProjectSaved = (bool)project.Saved,
                    Host = native.HostKind,
                    HostAvailable = true,
                    HostPath = state.Path,
                    HostSaved = (bool?)state.Saved,
                    HostReadOnly = (bool?)state.ReadOnly,
                    HostHasPath = (bool?)!string.IsNullOrWhiteSpace(state.Path),
                    NativeFileFormat = state.Format,
                    OwnerProcessId = native.CurrentProcessId,
                    IdentityVerified = true,
                    NativeQualification = "NOT_RUN",
                    Reason = (string)null
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    Project = selector,
                    ProjectSaved = (bool)project.Saved,
                    HostAvailable = false,
                    HostPath = (string)null,
                    HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null,
                    HostHasPath = (bool?)null,
                    NativeQualification = "NOT_RUN",
                    Reason = ex.Message
                };
            }
        }

        /// <summary>Saves a matched Office document after version, path, COM identity and process guards.</summary>
        /// <param name="request">Projet, version attendue et chemin éventuel fournis par l’appelant.</param>
        /// <param name="saveAs">Choisit une première sauvegarde sous un nouveau chemin.</param>
        /// <returns>Résultat de la sauvegarde native et état des vérifications.</returns>
        internal object SaveOtherHost(Request request, bool saveAs) => SaveOtherHost(request, saveAs, new NativeOtherHostProbe());

        /// <summary>Sauvegarde injectable; une exception après invocation produit explicitement un résultat incertain.</summary>
        /// <param name="request">Projet et préconditions de sauvegarde.</param>
        /// <param name="saveAs">Indique si une opération SaveAs est demandée.</param>
        /// <param name="native">Sonde des opérations hôte et du document associé.</param>
        /// <param name="approvedApplication">Optional retained Access application required throughout save preflight.</param>
        /// <returns>Résultat vérifié ou résultat incertain si une mutation a été appelée mais non confirmée.</returns>
        internal object SaveOtherHost(Request request, bool saveAs, IOtherHostProbe native, object approvedApplication = null)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion)) throw new ArgumentException("ExpectedProjectVersion is required.");
            dynamic project = GetDesignProject(request.Project); AssertProjectVersion(request, project);
            if ((int)project.Protection != 0) throw new InvalidOperationException("The project is protected.");
            object document = MatchOtherHost((object)project, native, approvedApplication);
            try
            {
                var before = native.State(document);
                if (saveAs && (native.HostKind == "Access" || native.HostKind == "Publisher"))
                    throw new InvalidOperationException("Access/Publisher project association requires an existing saved document; first SaveAs is unavailable.");
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
                    if (!OtherHostSamePath(before.Path, path) || !OtherHostProjectPathMatches((object)project, path, native.HostKind))
                        throw new InvalidOperationException("The document or selected project's path changed since inspection.");
                    if (!native.FileExists(path)) throw new FileNotFoundException("The existing host document is absent.");
                    if (before.Format.HasValue && before.Format.Value != format) throw new InvalidOperationException("The native Office format does not match the supported document extension.");
                }
                string sourceSha = OtherHostSourceSha((object)project);
                AssertProjectVersion(request, project);
                object preflightApplication = native.Application();
                try
                {
                    if (native.ApplicationProcessId(preflightApplication) != (uint)native.CurrentProcessId ||
                        (approvedApplication != null && !native.SameProject(approvedApplication, preflightApplication)) ||
                        !native.SameProject((object)project, native.DocumentProject(document)))
                        throw new InvalidOperationException("Host process or document/project identity changed before saving.");
                }
                finally { if (approvedApplication != null) ReleaseAccessObservation(preflightApplication); }
                var preflight = native.State(document);
                if (preflight.ReadOnly || !string.Equals(preflight.Path, before.Path, StringComparison.OrdinalIgnoreCase) || preflight.Format != before.Format ||
                    OtherHostSourceSha((object)project) != sourceSha)
                    throw new InvalidOperationException("Host path, format, read-only state or live source changed during save preparation.");
                if (native is NativeOtherHostProbe nativeProbe) nativeProbe.PrepareSave(document);
                try
                {
                    native.Save(document, saveAs, path, format);
                    var after = native.State(document);
                    if ((native.HostKind == "Access" || native.HostKind == "Publisher") &&
                        native.ApplicationProcessId(native.Application()) != (uint)native.CurrentProcessId)
                        throw new InvalidOperationException("The host process changed after saving.");
                    // Preserve short-circuit ordering: an uncertain earlier observation must not
                    // trigger additional COM reads. Report the exact failed check without source/path data.
                    if (!OtherHostSamePath(after.Path, path))
                        throw new InvalidOperationException("Save verification failed: HostPath.");
                    if (!OtherHostProjectPathMatches((object)project, path, native.HostKind))
                        throw new InvalidOperationException("Save verification failed: ProjectPath.");
                    if (native.HostKind != "Access" && after.Saved != true)
                        throw new InvalidOperationException("Save verification failed: HostSaved.");
                    if (!(bool)project.Saved)
                        throw new InvalidOperationException("Save verification failed: ProjectSaved.");
                    if (!native.FileExists(path))
                        throw new InvalidOperationException("Save verification failed: FileExists.");
                    if (native.FileLength(path) < 1)
                        throw new InvalidOperationException("Save verification failed: FileLength.");
                    if (after.Format.HasValue && after.Format.Value != format)
                        throw new InvalidOperationException("Save verification failed: FileFormat.");
                    if (OtherHostSourceSha((object)project) != sourceSha)
                        throw new InvalidOperationException("Save verification failed: SourceSha256.");
                    if (!native.SameProject((object)project, native.DocumentProject(document)))
                        throw new InvalidOperationException("Save verification failed: ProjectIdentity.");
                    return new
                    {
                        request.Project,
                        Host = native.HostKind,
                        HostPath = path,
                        SaveAsInvoked = saveAs,
                        SaveInvoked = !saveAs,
                        Verified = true,
                        Uncertain = false,
                        MutationInvoked = true,
                        HostSaved = after.Saved,
                        ProjectSaved = true,
                        Bytes = native.FileLength(path),
                        SourceSha256 = sourceSha,
                        CodePreserved = true,
                        NativeFileFormatVerified = after.Format.HasValue,
                        NativeQualification = "NOT_RUN",
                        PersistenceReopenVerified = false,
                        Limit = "Live code, available host state and file presence were verified. Access has no document Saved property. Reopen the native file to verify code persistence; native qualification remains NOT_RUN."
                    };
                }
                catch (Exception ex)
                {
                    if (native is NativeOtherHostProbe invokedProbe && !invokedProbe.SaveInvocationStarted) throw;
                    return new
                    {
                        request.Project,
                        Host = native.HostKind,
                        HostPath = path,
                        MutationInvoked = true,
                        SaveAsInvoked = saveAs,
                        SaveInvoked = !saveAs,
                        Verified = false,
                        Uncertain = true,
                        NativeQualification = "NOT_RUN",
                        Reason = ex.Message,
                        Next = "Read project_persistence_status and inspect the file; do not retry automatically."
                    };
                }
            }
            finally { if (approvedApplication != null) ReleaseAccessObservation(document); }
        }

        /// <summary>Associe un seul document au projet IUnknown et refuse toute application d'un autre PID.</summary>
        /// <param name="project">VBProject sélectionné par l’appelant.</param>
        /// <param name="native">Sonde qui fournit l’application, ses documents et leurs identités.</param>
        /// <param name="approvedApplication">Optional retained Access application whose COM identity must remain exact.</param>
        /// <returns>Document unique dont le VBProject partage l’identité COM avec le projet demandé.</returns>
        /// <exception cref="InvalidOperationException">L’hôte, le processus ou l’identité du document ne peut pas être vérifié de façon unique.</exception>
        internal static object MatchOtherHost(object project, IOtherHostProbe native, object approvedApplication = null)
        {
            if (native == null || !new[] { "Word", "PowerPoint", "Access", "Publisher" }.Contains(native.HostKind)) throw new InvalidOperationException("This host has no Office document save adapter.");
            if (native is NativeOtherHostProbe nativeProbe) nativeProbe.BindProject(project);
            object application = native.Application();
            IList<object> documents = null;
            object matchedDocument = null;
            try
            {
                if (approvedApplication != null && (native.HostKind != "Access" || !native.SameProject(approvedApplication, application)))
                    throw new InvalidOperationException("Access Save verification failed: owning application identity changed.");
                if (native.CurrentProcessId <= 0 || native.ApplicationProcessId(application) != (uint)native.CurrentProcessId)
                    throw new InvalidOperationException("The application does not belong to this add-in process.");
                documents = native.Documents(application);
                if (documents == null || documents.Count > 1000) throw new InvalidOperationException("The document collection is unreadable or oversized.");
                var matches = documents.Where(document => native.SameProject(project, native.DocumentProject(document))).ToArray();
                if (matches.Length != 1) throw new InvalidOperationException("No unique open document shares the selected VBProject COM identity.");
                if (native is NativeOtherHostProbe matchedProbe) matchedProbe.BindDocument(matches[0]);
                matchedDocument = matches[0];
                return matchedDocument;
            }
            finally
            {
                // Only the deferred Access route retains its own approved application acquisition.
                if (approvedApplication != null)
                {
                    if (documents != null)
                        foreach (object document in documents)
                            if (!ReferenceEquals(document, matchedDocument)) ReleaseAccessObservation(document);
                    ReleaseAccessObservation(application);
                }
            }
        }

        /// <summary>Maps supported native document extensions to exact Microsoft file format constants.</summary>
        /// <param name="kind">Word, PowerPoint, Access or Publisher.</param>
        /// <param name="path">Chemin dont l’extension détermine le format.</param>
        /// <returns>Constante SaveAs native associée au format macro-enabled.</returns>
        /// <exception cref="ArgumentException">L’hôte ou l’extension n’est pas prise en charge.</exception>
        internal static int OtherHostFormat(string kind, string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (kind == "Word") { if (extension == ".docm") return 13; if (extension == ".dotm") return 15; }
            if (kind == "PowerPoint") { if (extension == ".pptm") return 25; if (extension == ".potm") return 27; if (extension == ".ppsm") return 29; }
            if (kind == "Publisher" && extension == ".pub") return 1;
            if (kind == "Access" && extension == ".accdb") return 12;
            throw new ArgumentException("Word requires .docm/.dotm; PowerPoint requires .pptm/.potm/.ppsm; existing Access documents require .accdb; existing Publisher documents require .pub. Other formats are refused.");
        }

        /// <summary>Lit FileName et distingue une absence de chemin d'une autre erreur native.</summary>
        /// <param name="project">Projet VBA du document.</param>
        /// <returns>Chemin du projet ou chaîne vide si l’hôte ne fournit pas FileName.</returns>
        private static string OtherHostProjectPath(object project)
        {
            try { return (string)((dynamic)project).FileName; }
            catch (COMException error) when (error.ErrorCode == unchecked((int)0x800A004C)) { return ""; }
            catch (DirectoryNotFoundException error) when (error.HResult == unchecked((int)0x80070003)) { return ""; }
        }

        /// <summary>Vérifie le chemin VBIDE seulement lorsque ce chemin représente le document hôte.</summary>
        /// <param name="project">Matched VBProject whose path is compared for hosts where FileName identifies the document.</param>
        /// <param name="expectedPath">Identity-matched native document path being verified.</param>
        /// <param name="hostKind">Canonical host name; Word uses document FullName and Publisher may use its strict sole-project identity rule.</param>
        /// <returns>True when host-specific path rules confirm the same document; otherwise the canonical paths must match.</returns>
        private static bool OtherHostProjectPathMatches(object project, string expectedPath, string hostKind)
        {
            // MatchOtherHost and the final readback still require the same document/project COM identity,
            // owning PID, native FullName, macro format, saved flags and file bytes. Word's FileName
            // identifies VBA backing storage: it may throw before Save and become ~WRLxxxx.tmp after
            // Save. Only the identity-matched Word Document.FullName identifies its persisted file.
            if (hostKind == "Word") return true;
            string projectPath = OtherHostProjectPath(project);
            // Publisher can also omit FileName. Its native probe then requires the sole document,
            // sole VBIDE project, exact COM identities and current PID on every association/readback.
            if (hostKind == "Publisher" && string.IsNullOrWhiteSpace(projectPath)) return true;
            return OtherHostSamePath(projectPath, expectedPath);
        }

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
                rows.Add(new
                {
                    Name = (string)component.Name,
                    Type = (int)component.Type,
                    SourceSha = Hash(count == 0 ? "" : (string)code.Lines(1, count))
                });
            }
            return Hash(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(rows));
        }
    }
}
