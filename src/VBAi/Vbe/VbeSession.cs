using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Route les commandes vers les services VBE partagés et les adaptateurs propres à chaque application hôte.</summary>
    internal sealed partial class VbeSession
    {

        /// <summary>Fournit la fenêtre d’éditeur moderne, en pouvant la créer à la demande.</summary>
        internal Func<bool, ModernEditorWindow> ModernEditor;

        /// <summary>Session-owned testing service shared with the native explorer.</summary>
        internal VbeTestExplorerService TestExplorer;

        /// <summary>Per-dispatch assistant authorization captured by an asynchronous test batch.</summary>
        internal Action TestExecutionGuard;

        /// <summary>Résout un composant du projet vers l’adaptateur de module VBE.</summary>
        /// <param name="projectName">Sélecteur du projet dans la session.</param>
        /// <param name="moduleName">Nom exact du composant à résoudre.</param>
        /// <returns>Adaptateur lié aux objets COM du projet et du composant trouvé.</returns>
        internal IEditorModule ResolveEditorModule(string projectName, string moduleName)
        {
            dynamic project = GetProject(projectName);
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module is required.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                    return new EditorVbeModule((object)vbe, (object)project, (object)component);
            throw new InvalidOperationException("Module not found: " + moduleName);
        }

        /// <summary>Instance VBE cible utilisée pour résoudre projets et modules.</summary>
        private readonly dynamic vbe;

        /// <summary>Service des opérations de débogage et des boîtes de dialogue natives.</summary>
        private readonly VbeDebug debugger;

        /// <summary>Service de lecture et de modification des UserForms.</summary>
        private readonly VbeForms forms;

        /// <summary>Service des propriétés, composants et références des projets VBA.</summary>
        private readonly VbeProjectComponents components;

        /// <summary>Service d’inspection et de contrôle des fenêtres de l’éditeur VBE.</summary>
        private readonly VbeEditorWindows editorWindows;

        /// <summary>Service de recherche et de modification des procédures et fichiers de code.</summary>
        private readonly VbeCodeNavigation codeNavigation;

        /// <summary>Service d’inspection des bibliothèques et types exposés par les références.</summary>
        private readonly VbeReferenceTypes referenceTypes;

        /// <summary>Service d’édition du code source par opérations préparées.</summary>
        private readonly VbeCodeEdits codeEdits;

        /// <summary>Service de lecture et d’écriture du presse-papiers de code.</summary>
        private readonly VbeCodeClipboard codeClipboard;

        /// <summary>Service de navigation et d’historique de positions du code.</summary>
        private readonly VbeNavigationHistory navigationHistory;

        /// <summary>Abstraction du magasin de certificats utilisée pour lire les certificats de signature.</summary>
        internal interface ISigningStore : IDisposable
        {

            /// <summary>Certificats présents dans le magasin ouvert.</summary>
            /// <value>Collection fournie par le magasin actuellement ouvert.</value>
            X509Certificate2Collection Certificates { get; }

            /// <summary>Ouvre le magasin selon les droits indiqués.</summary>
            /// <param name="flags">Options d’ouverture du magasin de certificats.</param>
            void Open(OpenFlags flags);
        }

        /// <summary>Adaptateur vers le magasin de certificats Windows.</summary>
        private sealed class NativeSigningStore : ISigningStore
        {

            /// <summary>Magasin Windows encapsulé.</summary>
            private readonly X509Store store;

            /// <summary>Crée un accès au magasin personnel de l’emplacement indiqué.</summary>
            /// <param name="location">Emplacement Windows du magasin personnel à ouvrir.</param>
            public NativeSigningStore(StoreLocation location) { store = new X509Store(StoreName.My, location); }

            /// <summary>Expose les certificats du magasin natif.</summary>
            /// <value>Collection de certificats exposée par le magasin Windows.</value>
            public X509Certificate2Collection Certificates => store.Certificates;

            /// <summary>Ouvre le magasin natif avec les options demandées.</summary>
            /// <param name="flags">Options d’ouverture du magasin de certificats.</param>
            public void Open(OpenFlags flags) { store.Open(flags); }

            /// <summary>Libère le magasin natif.</summary>
            public void Dispose() { store.Dispose(); }
        }

        // Keep OS reads and native scheduling injectable without changing the signing checks.
        /// <summary>Fabrique injectable de magasins de certificats, initialisée avec l’implémentation Windows.</summary>
        internal Func<StoreLocation, ISigningStore> SigningStore = location => new NativeSigningStore(location);

        /// <summary>Fournit le nom du processus hôte pour appliquer les validations Excel.</summary>
        internal Func<string> SigningProcessName = () => System.Diagnostics.Process.GetCurrentProcess().ProcessName;

        /// <summary>Horloge utilisée pour vérifier la période de validité du certificat.</summary>
        internal Func<DateTime> SigningClock = () => DateTime.Now;

        /// <summary>Planifie l’affichage de la boîte de signature native pour une requête validée.</summary>
        internal Func<Request, object> SignatureScheduler;

        /// <summary>Crée une session pour le VBE fourni.</summary>
        /// <param name="vbe">Objet VBE auquel rattacher la session.</param>
        public VbeSession(object vbe) : this(vbe, null, null) { }

        /// <summary>Crée une session avec un stockage explicite des signets.</summary>
        /// <param name="vbe">Objet VBE de l’hôte.</param>
        /// <param name="bookmarkDatabase">Base SQLite des signets.</param>
        public VbeSession(object vbe, string bookmarkDatabase) : this(vbe, null, bookmarkDatabase) { }

        /// <summary>Crée les services de session et permet d’injecter la sonde d’hôte Excel.</summary>
        /// <param name="vbe">Objet VBE auquel rattacher la session.</param>
        /// <param name="host">Sonde utilisée pour distinguer les hôtes Excel lors des opérations concernées.</param>
        internal VbeSession(object vbe, VbeProjectComponents.IExcelHostProbe host) : this(vbe, host, null) { }

        /// <summary>Initialise les services VBE, la sonde d’hôte et le stockage des signets.</summary>
        /// <param name="vbe">Objet VBE de l’hôte.</param>
        /// <param name="host">Sonde d’hôte injectable.</param>
        /// <param name="bookmarkDatabase">Base SQLite des signets ou null pour le stockage par défaut.</param>
        private VbeSession(object vbe, VbeProjectComponents.IExcelHostProbe host, string bookmarkDatabase)
            : this(vbe, host, bookmarkDatabase, () => System.Diagnostics.Process.GetCurrentProcess().ProcessName,
                () => Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)) { }

        /// <summary>Injecte uniquement l'identité du processus et la racine locale avant la restauration native des barres.</summary>
        /// <param name="vbe">Objet COM du VBE utilisé par les services de session.</param>
        /// <param name="host">Sonde facultative de l’hôte Excel.</param>
        /// <param name="bookmarkDatabase">Chemin SQLite des signets, ou nul pour le stockage par défaut.</param>
        /// <param name="toolbarProcessName">Fournit le nom de processus utilisé pour choisir le profil natif.</param>
        /// <param name="localApplicationData">Fournit le dossier de données locales pour le profil des barres.</param>
        internal VbeSession(object vbe, VbeProjectComponents.IExcelHostProbe host, string bookmarkDatabase,
            Func<string> toolbarProcessName, Func<string> localApplicationData) { this.vbe = vbe; debugger = new VbeDebug(vbe);
            forms = new VbeForms(vbe); components = host == null
                ? new VbeProjectComponents(vbe, forms) : new VbeProjectComponents(vbe, forms, host);
            editorWindows = new VbeEditorWindows(vbe);
            string toolbarHost = toolbarProcessName().ToUpperInvariant();
            if (toolbarHost == "EXCEL" || toolbarHost == "SLDWORKS")
            {
                editorWindows.ToolbarProfiles = new VbeToolbarProfiles(Path.Combine(localApplicationData(), "VBAi", "VbeToolbars", toolbarHost + ".sqlite"));
                editorWindows.RestoreToolbarProfiles();
            }
            codeNavigation = new VbeCodeNavigation(vbe, forms);
            referenceTypes = new VbeReferenceTypes(vbe);
            codeEdits = new VbeCodeEdits(Execute);
            codeClipboard = new VbeCodeClipboard(Execute, new WindowsCodeClipboard());
            navigationHistory = new VbeNavigationHistory(vbe, Execute, bookmarkDatabase);
            SignatureScheduler = request => debugger.QueueSignatureDialog(request); }

        /// <summary>Vérifie qu’une coupe de formulaire peut être récupérée.</summary>
        /// <param name="request">Identité et version de la sauvegarde.</param>
        /// <returns>Disponibilité de la récupération.</returns>
        internal bool CanRecoverFormCut(Request request) { return forms.CanRecoverCut(request); }

        /// <summary>Obtient la collection émettant les événements de projet.</summary>
        /// <returns>Collection VBProjects native.</returns>
        internal object ProjectsEventSource() { return vbe.VBProjects; }

        /// <summary>Resolves a borrowed live project for private owning-thread conversation identity checks.</summary>
        /// <param name="selector">Text that supplies the selector value. Use the format required by the calling operation.</param>
        /// <returns>object produced by the operation for project scope source on vbe session.</returns>
        internal object ProjectScopeSource(string selector) => (object)GetProject(selector);

        /// <summary>Obtient la collection émettant les événements de composants.</summary>
        /// <param name="project">Projet ciblé.</param>
        /// <returns>Collection VBComponents ou null sans projet.</returns>
        internal object ComponentsEventSource(string project)
        {
            if (string.IsNullOrEmpty(project)) return null;
            return VbeProjectResolver.Resolve(vbe, project).VBComponents;
        }

        /// <summary>Obtient les événements de références du projet.</summary>
        /// <param name="project">Projet ciblé.</param>
        /// <returns>Source d’événements native ou null sans projet.</returns>
        internal object ReferenceEventSource(string project)
        {
            if (string.IsNullOrEmpty(project)) return null;
            return vbe.Events.ReferencesEvents[VbeProjectResolver.Resolve(vbe, project)];
        }

        /// <summary>Runs an explicit Immediate capture on the caller's VBE UI context.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for read immediate async on vbe session.</returns>
        internal System.Threading.Tasks.Task<object> ReadImmediateAsync(Request request)
        {
            RequireGeneralSettled();
            return debugger.ReadImmediateAsync(request);
        }

        /// <summary>Inspects declared scalar locals through the asynchronous native debugger route.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for inspect local scalars async on vbe session.</returns>
        internal System.Threading.Tasks.Task<object> InspectLocalScalarsAsync(Request request)
        {
            RequireGeneralSettled();
            return debugger.InspectLocalScalarsAsync(request);
        }

        /// <summary>Allows queued native saves to finish without blocking the VBE message loop.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save host document async on vbe session.</returns>
        internal System.Threading.Tasks.Task<object> SaveHostDocumentAsync(Request request)
        {
            RequireGeneralSettled();
            return components.SaveHostDocumentAsync(request);
        }

        /// <summary>Maintains the general in flight and general quarantined state for vbe session.</summary>
        private bool generalInFlight, generalQuarantined;

        /// <summary>Maintains the general authorization depth state for vbe session.</summary>
        private int generalAuthorizationDepth;

        /// <summary>Maintains the bridge operations in flight state for vbe session.</summary>
        private int bridgeOperationsInFlight;

        /// <summary>Requires general settled for vbe session.</summary>
        internal void RequireGeneralSettled()
        {
            RequireMacroSettled();
            if (generalInFlight || generalQuarantined)
                throw new InvalidOperationException("An original General operation is pending or uncertain. No further native operation is permitted in this session.");
        }

        // Claimed and released on the bridge's owning STA. The worker keeps its
        // existing native thread while General cannot enter during its dispatch.
        /// <summary>Handles admit bridge operation for vbe session.</summary>
        /// <returns>action produced by the operation for admit bridge operation on vbe session.</returns>
        internal Action AdmitBridgeOperation()
        {
            RequireGeneralSettled();
            int owner = System.Threading.Thread.CurrentThread.ManagedThreadId;
            if (System.Threading.Thread.CurrentThread.GetApartmentState() != System.Threading.ApartmentState.STA)
                throw new InvalidOperationException("Bridge admission requires the owning STA.");
            bridgeOperationsInFlight++;
            bool released = false;
            return () => {
                if (System.Threading.Thread.CurrentThread.ManagedThreadId != owner)
                    throw new InvalidOperationException("Bridge admission must be released on its owning STA.");
                if (released) return;
                released = true;
                bridgeOperationsInFlight--;
            };
        }

        /// <summary>Reads or edits the native General page on the original VBE UI thread.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="write">Indicates whether write is enabled.</param>
        /// <returns>task&lt;object&gt; produced by the operation for project general async on vbe session.</returns>
        internal async System.Threading.Tasks.Task<object> ProjectGeneralAsync(Request request, bool write)
        {
            RequireGeneralSettled();
            if (bridgeOperationsInFlight != 0)
                throw new InvalidOperationException("A bridge operation is pending. General cannot enter until it settles.");
            int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            Action requireContext = () => {
                if (System.Threading.Thread.CurrentThread.ManagedThreadId != ownerThread ||
                    System.Threading.Thread.CurrentThread.GetApartmentState() != System.Threading.ApartmentState.STA)
                    throw new InvalidOperationException("General left its original owning STA.");
            };
            Action<VbeProjectGeneralOperation.Result> journal = result => {
                requireContext();
                // Claims are durable before dispatch and deliberately omit project data and file paths.
                LoadLog.AppendText(LoadLog.PathName, DateTime.UtcNow.ToString("o") +
                    " General claim Open=" + result.OpenAttempts + " Field=" + result.FieldAttempts +
                    " OK=" + result.OkAttempts + " Cancel=" + result.CancelAttempts +
                    " Terminal=" + result.Terminal + " Mutation=" + result.MutationInvoked +
                    " Uncertain=" + result.Uncertain + " Closed=" + result.DialogClosed +
                    " ExecuteReturned=" + result.OriginalExecuteReturned + Environment.NewLine);
            };
            generalInFlight = true;
            Action<bool> originalAuthorization = request?.RevalidateProjectPropertyAuthorization;
            Action<bool> scopedAuthorization = live => {
                requireContext();
                if (originalAuthorization == null) throw new InvalidOperationException("Original General authorization is required.");
                if (live) generalAuthorizationDepth++;
                try { originalAuthorization(live); }
                finally { if (live) generalAuthorizationDepth--; }
            };
            if (request != null) request.RevalidateProjectPropertyAuthorization = scopedAuthorization;
            try
            {
                var result = (VbeProjectGeneralOperation.Result)await components.ProjectGeneralAsync(
                    request, write, debugger.CaptureGeneralCommand, journal, requireContext);
                generalQuarantined = result.Uncertain;
                return result;
            }
            finally
            {
                if (request != null && request.RevalidateProjectPropertyAuthorization == scopedAuthorization)
                    request.RevalidateProjectPropertyAuthorization = originalAuthorization;
                generalInFlight = false;
            }
        }

        /// <summary>Exécute la commande demandée et encapsule son résultat dans une réponse.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>Réponse contenant le résultat de la commande ou son erreur de validation.</returns>
        public Response Execute(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Command))
                return Response.Failure("A command is required.");

            // The live conversation authorization reads the project inventory. Permit
            // that read during the operation, while retaining all mutation exclusions.
            if (request.Command != "status" && (MacroDispatchBlocked(request.Command) || generalQuarantined ||
                (generalInFlight && !(generalAuthorizationDepth > 0 && request.Command == "list_projects"))))
                return Response.Failure("An original General operation is pending or uncertain. No further session operation is permitted.");

            switch (request.Command)
            {
                case "create_solidworks_macro":
                case "publish_solidworks_macro":
                    return Response.Failure("Explicit native macro creation/publication requires InvokeAsync or the bridge worker.");
                case "read_project_general":
                case "set_project_general":
                    return Response.Failure("The native General command requires InvokeAsync or the bridge worker.");
                case "discover_vba_tests":
                case "preview_vba_test_support":
                case "install_vba_test_support":
                case "run_vba_tests":
                case "vba_test_run_status":
                case "stop_vba_tests":
                case "navigate_vba_test":
                case "vba_test_coverage":
                case "show_vba_test_explorer":
                    if (TestExplorer == null) return Response.Failure("The session test explorer is unavailable.");
                    return Response.Success(TestExplorer.Command(request, TestExecutionGuard));
                case "status":
                    return Response.Success(new { Version = "0.1.0", Connected = true,
                        AssemblyPath = typeof(VbeSession).Assembly.Location,
                        AssemblyModuleVersionId = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                        HostProcessId = System.Diagnostics.Process.GetCurrentProcess().Id,
                        ProcessBitness = IntPtr.Size * 8 });
                case "preview_procedure_rename": return Response.Success(codeEdits.PreviewProcedureRename(request));
                case "apply_procedure_rename": return Response.Success(codeEdits.ApplyProcedureRename(request));
                case "preview_class_member_rename": return Response.Success(codeEdits.PreviewClassMemberRename(request));
                case "apply_class_member_rename": return Response.Success(codeEdits.ApplyClassMemberRename(request));
                case "open_native_ide_dialog": return Response.Success(debugger.QueueNativeIdeDialog(request));
                case "read_project_protection":
                case "set_project_protection":
                    if (request.ExpectedMode != 2) return Response.Failure("ExpectedMode=2 is required for project properties.");
                    return Response.Success(debugger.QueueProjectPropertiesDialog(request, candidate =>
                        string.Equals((string)((dynamic)components.ProjectProperties(candidate.Project)).Version, candidate.ExpectedProjectVersion, StringComparison.OrdinalIgnoreCase)));
                case "project_collection_state": return Response.Success(components.ProjectCollectionState());
                case "create_standalone_project": return Response.Success(components.CreateStandaloneProject(request));
                case "open_standalone_project": return Response.Success(components.OpenStandaloneProject(request));
                case "close_standalone_project": return Response.Success(components.CloseStandaloneProject(request));
                case "open_project_help": return Response.Success(components.OpenProjectHelp(request));
                case "list_macros": return Response.Success(debugger.ListMacros(request));
                case "read_navigation_surface":
                case "change_navigation_surface": return Response.Failure("This native navigation command requires InvokeAsync or the bridge worker.");
                case "list_projects":
                    return Response.Success(ListProjects());
                case "list_modules":
                    return Response.Success(ListModules(request.Project));
                case "vbe_windows":
                    return Response.Success(editorWindows.Windows());
                case "vbe_environment":
                    return Response.Success(editorWindows.Environment());
                case "set_addin_connection": return Response.Success(editorWindows.SetAddInConnection(request));
                case "list_addins":
                    return Response.Success(editorWindows.AddIns());
                case "focus_vbe_window":
                    return Response.Success(editorWindows.FocusWindow(request.WindowCaption, request.WindowType));
                case "show_vbe_window":
                    return Response.Success(editorWindows.ShowWindow(request.WindowCaption, request.WindowType));
                case "window_layout": return Response.Success(editorWindows.WindowLayout(request.WindowCaption, request.WindowType));
                case "set_window_state": return Response.Success(editorWindows.SetWindowState(request));
                case "set_window_bounds": return Response.Success(editorWindows.SetWindowBounds(request));
                case "link_vbe_window": return Response.Success(editorWindows.LinkWindow(request));
                case "window_linkage":
                    return Response.Success(editorWindows.WindowLinkage(request.WindowCaption, request.WindowType));
                case "close_vbe_window":
                    return Response.Success(editorWindows.CloseWindow(request.WindowCaption, request.WindowType));
                case "set_code_view": return Response.Success(debugger.SetCodePaneView(request));
                case "code_pane_layout": return Response.Success(debugger.CodePaneLayout(request));
                case "scroll_code_pane": return Response.Success(debugger.ScrollCodePane(request));
                case "editor_layout": return Response.Success(debugger.EditorLayout());
                case "arrange_editor_windows": return Response.Success(debugger.ArrangeEditorWindows(request));
                case "native_code_navigation": return Response.Success(debugger.NativeNavigation(request));
                case "set_code_split": return Response.Success(debugger.SetCodeSplit(request));
                case "code_panes":
                    return Response.Success(editorWindows.CodePanes());
                case "open_object_browser":
                    return Response.Success(debugger.OpenObjectBrowser(editorWindows));
                case "project_symbols": return Response.Success(codeNavigation.ProjectSymbols(request));
                case "navigate_code": return Response.Success(navigationHistory.Go(request));
                case "code_bookmark": return Response.Success(navigationHistory.Bookmark(request));
                case "list_procedures":
                    return Response.Success(codeNavigation.Procedures(request.Project, request.Module));
                case "find_code":
                    return Response.Success(codeNavigation.Find(request));
                case "select_procedure":
                    return Response.Success(codeNavigation.SelectProcedure(request, debugger));
                case "create_event_procedure":
                    return Response.Success(codeNavigation.CreateEventProcedure(request));
                case "create_procedure":
                    return Response.Success(codeNavigation.CreateProcedure(request));
                case "replace_procedure":
                    return Response.Success(codeNavigation.ReplaceProcedure(request));
                case "remove_procedure":
                    return Response.Success(codeNavigation.RemoveProcedure(request));
                case "insert_code_file":
                    return Response.Success(codeNavigation.InsertCodeFile(request));
                case "inspect_code_file":
                    return Response.Success(codeNavigation.InspectCodeFile(request.Path));
                case "project_properties":
                    return Response.Success(components.ProjectProperties(request.Project));
                case "project_persistence_status":
                    return Response.Success(components.PersistenceStatus(request.Project));
                case "save_host_document":
                    return Response.Success(components.SaveHostDocument(request));
                case "save_host_document_as":
                    return Response.Success(components.SaveHostDocumentAs(request));
                case "project_signature_status":
                    return Response.Success(components.SignatureStatus(request.Project));
                case "verify_vba_signature_file":
                    return Response.Success(new VbeSignatureVerifier().Verify(request.Path));
                case "certificate_trust": return Response.Success(CertificateTrust(request));
                case "list_signing_certificates":
                    return Response.Success(ListSigningCertificates());
                case "read_project_signature_dialog":
                    return Response.Success(debugger.QueueSignatureDialog(request));
                case "sign_project":
                    return Response.Success(BeginSignProject(request));
                case "component_properties":
                    return Response.Success(components.ComponentProperties(request.Project, request.Module));
                case "component_property_value":
                    return Response.Success(components.ComponentPropertyValue(request.Project, request.Module, request.Property));
                case "component_probe":
                    return Response.Success(components.ComponentProbe(request.Project, request.Module, request.Action, request.Query));
                case "set_project_property":
                    return Response.Success(components.SetProjectProperty(request));
                case "set_component_property":
                    return Response.Success(components.SetComponentProperty(request));
                case "set_class_instancing":
                    return Response.Success(components.SetClassInstancing(request));
                case "rename_project":
                    return Response.Failure("Project rename is disabled: it correlated with an Excel process crash during validation.");
                case "rename_component":
                    return Response.Success(components.RenameComponent(request));
                case "remove_component":
                    return Response.Success(components.RemoveComponent(request));
                case "import_component":
                    return Response.Success(components.ImportComponent(request));
                case "export_component":
                    return Response.Success(components.ExportComponent(request));
                case "list_references":
                    return Response.Success(ListReferences(request.Project));
                case "list_reference_types":
                    return Response.Success(referenceTypes.ListTypes(request));
                case "list_type_members":
                    return Response.Success(referenceTypes.ListMembers(request));
                case "add_reference_guid":
                    return Response.Success(AddReferenceGuid(request));
                case "add_reference_file":
                    return Response.Success(AddReferenceFile(request));
                case "remove_reference":
                    return Response.Success(RemoveReference(request));
                case "read_module":
                    return Response.Success(ReadModule(request.Project, request.Module));
                case "create_module":
                    return Response.Success(CreateComponent(request, 1));
                case "create_class":
                    return Response.Success(CreateComponent(request, 2));
                case "preview_local_rename": return Response.Success(codeEdits.RenameLocal(request, true));
                case "apply_local_rename": return Response.Success(codeEdits.RenameLocal(request, false));
                case "preview_parameter_rename": return Response.Success(codeEdits.RenameParameter(request, true));
                case "apply_parameter_rename": return Response.Success(codeEdits.RenameParameter(request, false));
                case "preview_code_edit": return Response.Success(codeEdits.Edit(request, true));
                case "apply_code_edit": return Response.Success(codeEdits.Edit(request, false));
                case "toolbar_controls": return Response.Success(editorWindows.ToolbarControls(request));
                case "create_toolbar": return Response.Success(editorWindows.CreateToolbar(request));
                case "remove_toolbar": return Response.Success(editorWindows.RemoveToolbar(request));
                case "add_toolbar_command": return Response.Success(editorWindows.AddToolbarCommand(request));
                case "remove_toolbar_command": return Response.Success(editorWindows.RemoveToolbarCommand(request));
                case "list_toolbars": return Response.Success(editorWindows.Toolbars());
                case "set_toolbar_placement": return Response.Success(editorWindows.SetToolbarPlacement(request));
                case "set_toolbar_position": return Response.Success(editorWindows.SetToolbarPosition(request));
                case "set_toolbar_visibility": return Response.Success(editorWindows.SetToolbarVisibility(request));
                case "read_code_clipboard": return Response.Success(codeClipboard.Read());
                case "copy_code": return Response.Success(codeClipboard.Edit(request, "copy"));
                case "cut_code": return Response.Success(codeClipboard.Edit(request, "cut"));
                case "paste_code": return Response.Success(codeClipboard.Edit(request, "paste"));
                case "form_clipboard_state": return Response.Success(forms.ClipboardState(request.Project, request.Form, request.ParentPath));
                case "recover_form_cut": return Response.Success(forms.RecoverDesignerCut(request));
                case "restore_form_clipboard": return Response.Success(forms.RestoreDesignerClipboard(request));
                case "select_form_controls": return Response.Success(forms.SelectDesignerControls(request));
                case "native_form_clipboard": return Response.Success(forms.NativeClipboard(request));
                case "native_form_history": return Response.Success(forms.NativeHistory(request));
                case "native_code_history_state": return Response.Success(debugger.NativeCodeHistoryState(request));
                case "native_code_history": return Response.Success(debugger.NativeCodeHistory(request));
                case "undo_code_edit": return Response.Success(codeEdits.Replay(request, false));
                case "redo_code_edit": return Response.Success(codeEdits.Replay(request, true));
                case "replace_lines":
                    return ReplaceLines(request);
                case "debug_state":
                    return Response.Success(debugger.State(request.Project));
                case "run_form":
                    return Response.Success(forms.RunForm(request));
                case "form_run_status":
                    return Response.Success(forms.FormRunStatus(request));
                case "run_procedure":
                    return Response.Success(debugger.RunProcedure(request));
                case "run_procedure_values":
                    return Response.Success(debugger.RunProcedureValues(request));
                case "procedure_values_status":
                    return Response.Success(debugger.ProcedureValuesStatus(request));
                case "procedure_run_status":
                    return Response.Success(debugger.ProcedureRunStatus(request));
                case "run_sub":
                    return Response.Success(debugger.RunSub(request));
                case "compile_project":
                    return Response.Success(debugger.CompileProject(request));
                case "open_debug_pane":
                    return Response.Success(debugger.OpenDebugPane(request.Action, editorWindows));
                case "read_immediate":
                    return Response.Failure("Native Immediate capture requires the asynchronous bridge or tool route.");
                case "inspect_local_scalars":
                    return Response.Failure("Native local inspection requires the asynchronous bridge or tool route.");
                case "add_watch":
                    return Response.Success(debugger.QueueAddWatchDialog(request));
                case "edit_watch":
                    return Response.Success(debugger.QueueEditWatchDialog(request));
                case "quick_watch":
                    return Response.Success(debugger.QueueQuickWatchDialog(request));
                case "read_debug_options":
                    return Response.Success(debugger.QueueDebugOptionsDialog());
                case "set_vbe_option":
                case "read_vbe_options":
                    return Response.Success(debugger.QueueDebugOptionsDialog());
                case "remove_watch":
                    return Response.Success(debugger.RemoveSelectedWatch(request));
                case "debug_global":
                    return Response.Success(debugger.ExecuteGlobalDebugCommand(request));
                case "list_commands":
                    return Response.Success(debugger.ListCommands(request.Query, request.Offset, request.Limit));
                case "select_code":
                    return Response.Success(debugger.SelectCode(request));
                case "select_code_range":
                    return Response.Success(debugger.SelectCodeRange(request));
                case "invoke_debug":
                    return Response.Success(debugger.InvokeCommand(request));
                case "list_forms":
                    return Response.Success(forms.List(request.Project));
                case "list_form_control_types":
                    return Response.Success(forms.ControlTypes());
                case "form_state":
                    return Response.Success(forms.State(request.Project, request.Form));
                case "preview_fit_form_content": return Response.Success(forms.PreviewFitFormContent(request));
                case "apply_fit_form_content": return Response.Success(forms.ApplyFitFormContent(request));
                case "preview_form_layout": return Response.Success(forms.LayoutControls(request, true));
                case "apply_form_layout": return Response.Success(forms.LayoutControls(request, false));
                case "set_form_tab_order": return Response.Success(forms.SetTabOrder(request));
                case "form_tree":
                    return Response.Success(forms.Tree(request.Project, request.Form));
                case "form_list_items":
                    return Response.Success(forms.ListItems(request));
                case "set_form_list_binding": return Response.Success(forms.SetListBinding(request));
                case "set_form_list_initializer":
                    return Response.Success(forms.SetListInitializer(request));
                case "probe_append_form_list_item":
                    return Response.Success(forms.AppendListItem(request));
                case "add_form_list_item":
                    return Response.Success(forms.AddListItem(request));
                case "remove_form_list_item":
                    return Response.Success(forms.RemoveListItem(request));
                case "form_event_catalog":
                    return Response.Success(forms.EventCatalog(request.Project, request.Form, request.ControlPath));
                case "form_parent_probe":
                    return Response.Success(forms.ParentProbe(request.Project, request.Form));
                case "form_properties":
                    return Response.Success(forms.Properties(request.Project, request.Form));
                case "set_form_property":
                    return Response.Success(forms.SetProperty(request));
                case "set_form_picture":
                    return Response.Success(forms.SetPicture(request));
                case "form_control_properties":
                    return Response.Success(forms.ControlProperties(request.Project, request.Form, request.Control));
                case "create_form":
                    return Response.Success(forms.Create(request));
                case "open_form":
                    return Response.Success(forms.Open(request.Project, request.Form));
                case "add_form_control":
                    return Response.Success(forms.AddControl(request));
                case "add_nested_form_control":
                    return Response.Success(forms.AddNestedControl(request));
                case "set_form_node_property":
                    return Response.Success(forms.SetNodeProperty(request));
                case "set_form_node_picture":
                    return Response.Success(forms.SetNodePicture(request));
                case "z_order_form_control":
                    return Response.Success(forms.ZOrderControl(request));
                case "form_property_accessors":
                    return Response.Success(forms.PropertyAccessors(request));
                case "duplicate_form_label":
                    return Response.Success(forms.DuplicateLabel(request));
                case "duplicate_form_textbox":
                    return Response.Success(forms.DuplicateTextBox(request));
                case "duplicate_form_checkbox":
                    return Response.Success(forms.DuplicateCheckBox(request));
                case "duplicate_form_togglebutton":
                    return Response.Success(forms.DuplicateToggleButton(request));
                case "duplicate_form_commandbutton":
                    return Response.Success(forms.DuplicateCommandButton(request));
                case "duplicate_form_combobox":
                    return Response.Success(forms.DuplicateComboBox(request));
                case "duplicate_empty_form_frame":
                    return Response.Success(forms.DuplicateEmptyFrame(request));
                case "frame_copy_plan":
                    return Response.Success(forms.FrameCopyPlan(request));
                case "duplicate_form_frame_labels":
                    return Response.Success(forms.DuplicateFrameWithLabels(request));
                case "frame_simple_copy_plan":
                    return Response.Success(forms.FrameSimpleCopyPlan(request));
                case "duplicate_form_frame_simple_children":
                    return Response.Success(forms.DuplicateFrameWithSimpleChildren(request));
                case "frame_profile_copy_plan":
                    return Response.Success(forms.FrameProfileCopyPlan(request));
                case "duplicate_form_frame_profiled":
                    return Response.Success(forms.DuplicateFrameProfiled(request));
                case "duplicate_form_optionbutton":
                    return Response.Success(forms.DuplicateOptionButton(request));
                case "remove_form_control":
                    return Response.Success(forms.RemoveControl(request));
                case "add_form_page":
                    return Response.Success(forms.AddPageOrTab(request, "Pages"));
                case "add_form_tab":
                    return Response.Success(forms.AddPageOrTab(request, "Tabs"));
                case "remove_form_page_tab":
                    return Response.Success(forms.RemovePageOrTab(request));
                case "set_form_control_geometry":
                    return Response.Success(forms.SetControlGeometry(request));
                case "rename_form_control":
                    return Response.Success(forms.RenameControl(request));
                case "set_form_control_caption":
                    return Response.Success(forms.SetControlCaption(request));
                case "set_form_control_font":
                    return Response.Success(forms.SetControlFont(request));
                default:
                    return Response.Failure("Unknown command: " + request.Command);
            }
        }

        /// <summary>Valide le projet et le certificat, puis planifie la première signature VBA.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>Informations sur la signature planifiée et le certificat retenu.</returns>
        private object BeginSignProject(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.ExpectedProjectVersion) ||
                string.IsNullOrWhiteSpace(request.CertificateThumbprint))
                throw new ArgumentException("Project, ExpectedProjectVersion and CertificateThumbprint are required.");
            dynamic state = components.ProjectProperties(request.Project);
            if ((int)state.Mode != 2 || request.ExpectedMode != 2)
                throw new InvalidOperationException("The project must be in design mode (ExpectedMode=2).");
            if (!string.Equals((string)state.Version, request.ExpectedProjectVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since it was read.");
            dynamic liveProject = GetProject(request.Project);
            if (!(bool)liveProject.Saved)
                throw new InvalidOperationException("Save the VBA project before signing it.");
            dynamic signatureStatus = components.SignatureStatus(request.Project);
            if ((bool)signatureStatus.Available && (bool)signatureStatus.Signed)
                throw new InvalidOperationException("The host reports an existing VBA signature; this command only adds the first signature.");
            bool unsignedVerified = (bool)signatureStatus.Available && !(bool)signatureStatus.Signed;
            if (string.Equals(SigningProcessName(), "EXCEL",
                StringComparison.OrdinalIgnoreCase))
            {
                string projectPath = null;
                try { projectPath = (string)liveProject.FileName; }
                catch { }
                if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath) ||
                    !File.Exists(projectPath) || !unsignedVerified)
                    throw new InvalidOperationException("Save the macro-enabled Excel workbook before adding its first VBA signature.");
                string extension = Path.GetExtension(projectPath);
                if (!new[] { ".xlsm", ".xlam", ".xlsb", ".xltm", ".xls", ".xla", ".xlt" }
                    .Contains(extension, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The saved Excel format does not support a VBA project signature.");
                bool hasVbaContent = false;
                foreach (dynamic component in liveProject.VBComponents)
                    if ((int)component.Type != 100 || (int)component.CodeModule.CountOfLines > 0)
                    { hasVbaContent = true; break; }
                if (!hasVbaContent)
                    throw new InvalidOperationException("The Excel workbook has no VBA content to sign; add code and save it first.");
            }
            string thumbprint = request.CertificateThumbprint.Replace(" ", "").ToUpperInvariant();
            if (!Regex.IsMatch(thumbprint, "^[0-9A-F]{40}$"))
                throw new ArgumentException("CertificateThumbprint must be a SHA-1 certificate thumbprint.");
            using (var store = SigningStore(StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var matches = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (matches.Count != 1) throw new InvalidOperationException("The selected certificate is absent or ambiguous.");
                var certificate = matches[0];
                if (!certificate.HasPrivateKey || SigningClock() < certificate.NotBefore || SigningClock() > certificate.NotAfter)
                    throw new InvalidOperationException("The certificate needs a usable private key and current validity.");
                bool codeSigning = certificate.Extensions.OfType<X509EnhancedKeyUsageExtension>()
                    .SelectMany(extension => extension.EnhancedKeyUsages.Cast<System.Security.Cryptography.Oid>())
                    .Any(oid => oid.Value == "1.3.6.1.5.5.7.3.3");
                if (!codeSigning) throw new InvalidOperationException("The certificate is not intended for code signing.");
                string displayName = certificate.GetNameInfo(X509NameType.SimpleName, false);
                if (string.IsNullOrWhiteSpace(displayName))
                    throw new InvalidOperationException("The certificate display name is empty.");
                var matchingThumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in store.Certificates.Cast<X509Certificate2>())
                    if (string.Equals(item.GetNameInfo(X509NameType.SimpleName, false), displayName,
                        StringComparison.OrdinalIgnoreCase)) matchingThumbprints.Add(item.Thumbprint);
                using (var machineStore = SigningStore(StoreLocation.LocalMachine))
                {
                    machineStore.Open(OpenFlags.ReadOnly);
                    foreach (var item in machineStore.Certificates.Cast<X509Certificate2>())
                        if (string.Equals(item.GetNameInfo(X509NameType.SimpleName, false), displayName,
                            StringComparison.OrdinalIgnoreCase)) matchingThumbprints.Add(item.Thumbprint);
                }
                if (matchingThumbprints.Count != 1 || !matchingThumbprints.Contains(thumbprint))
                    throw new InvalidOperationException("The certificate display name is ambiguous across personal certificate stores.");
                object scheduled = SignatureScheduler(request);
                return new { Scheduled = true, Project = request.Project,
                    CertificateThumbprint = thumbprint, CertificateName = displayName,
                    UnsignedVerified = unsignedVerified,
                    NativeCommand = scheduled };
            }
        }

        /// <summary>Évalue hors ligne le certificat exact du magasin personnel sans utiliser la clé privée.</summary>
        /// <param name="request">Empreinte publique SHA-1 du certificat Windows.</param>
        /// <returns>Confiance de chaîne, distincte de la validité d’une signature de macro.</returns>
        private object CertificateTrust(Request request)
        {
            string thumbprint = (request.CertificateThumbprint ?? "").Replace(" ", "").ToUpperInvariant();
            if (thumbprint.Length != 40 || thumbprint.Any(c => !Uri.IsHexDigit(c)))
                throw new ArgumentException("CertificateThumbprint must be an exact SHA-1 certificate thumbprint.");
            using (var store = SigningStore(StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var matches = store.Certificates.Cast<X509Certificate2>().Where(c => c.Thumbprint == thumbprint).ToArray();
                if (matches.Length != 1) throw new InvalidOperationException("The exact certificate is absent or ambiguous in CurrentUser/My.");
                return VbeCertificateTrust.Evaluate(matches[0]);
            }
        }

        /// <summary>Retourne les certificats personnels admissibles à la signature de code.</summary>
        /// <returns>Certificats admissibles à la signature de code dans le magasin personnel.</returns>
        private object ListSigningCertificates()
        {
            using (var store = SigningStore(StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                return store.Certificates.Cast<X509Certificate2>()
                    .Where(certificate => certificate.HasPrivateKey &&
                        certificate.Extensions.OfType<X509EnhancedKeyUsageExtension>()
                            .SelectMany(extension => extension.EnhancedKeyUsages.Cast<System.Security.Cryptography.Oid>())
                            .Any(oid => oid.Value == "1.3.6.1.5.5.7.3.3"))
                    .Select(certificate => new {
                        certificate.Thumbprint,
                        Name = certificate.GetNameInfo(X509NameType.SimpleName, false),
                        certificate.Subject, certificate.Issuer,
                        NotBefore = certificate.NotBefore.ToString("o"),
                        NotAfter = certificate.NotAfter.ToString("o"),
                        EligibleNow = SigningClock() >= certificate.NotBefore && SigningClock() <= certificate.NotAfter
                    }).ToArray();
            }
        }

        /// <summary>Crée l’accès Git au projet VBA résolu.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <param name="hostPath">Chemin du document hôte associé au projet Git.</param>
        /// <returns>Adaptateur Git associé au projet VBE.</returns>
        internal VbaGitProject GitProject(string projectName, string hostPath)
        {
            return new VbaGitProject(() => (object)GetProject(projectName), hostPath);
        }

        /// <summary>Retourne le chemin absolu du document hôte enregistré utilisé comme périmètre Git.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <returns>Chemin absolu du document hôte enregistré.</returns>
        internal string GitScope(string projectName)
        {
            dynamic project = GetProject(projectName);
            string path = VbeProjectHostPath.Read((object)project);
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path)) throw new InvalidOperationException("Enregistrez le document avant d’utiliser Git.");
            return Path.GetFullPath(path);
        }

        /// <summary>Demande la persistance de la signature Excel du projet.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <returns>Résultat de la persistance de la signature Excel.</returns>
        internal object PersistProjectSignature(string projectName)
        {
            return components.PersistExcelSignature(projectName);
        }

        /// <summary>Énumère les projets VBE et leurs noms de fichier accessibles.</summary>
        /// <returns>Projets visibles avec nom, chemin accessible et mode.</returns>
        private object ListProjects()
        {
            var result = new List<object>();
            foreach (dynamic project in vbe.VBProjects)
            {
                string fileName = null;
                int? fileNameErrorHResult = null;
                string fileNameErrorType = null;
                try { fileName = (string)project.FileName; }
                catch (Exception error)
                {
                    // Preserve the original getter outcome without classifying it as unsaved.
                    fileNameErrorHResult = error.HResult;
                    fileNameErrorType = error.GetType().FullName;
                }
                string hostPath = null, hostPathError = null;
                try { hostPath = VbeProjectHostPath.Read((object)project); }
                catch (Exception error) { hostPathError = error.Message; }
                result.Add(new { Name = (string)project.Name, FileName = fileName, HostPath = hostPath,
                    HostPathError = hostPathError, Mode = (int)project.Mode,
                    FileNameErrorHResult = fileNameErrorHResult, FileNameErrorType = fileNameErrorType });
            }
            return result;
        }

        /// <summary>Énumère les composants du projet avec leur type et leur nombre de lignes.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <returns>Composants du projet avec type et nombre de lignes.</returns>
        private object ListModules(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                result.Add(new { Name = (string)component.Name, Type = (int)component.Type,
                    Lines = (int)component.CodeModule.CountOfLines });
            return result;
        }

        /// <summary>Retourne le code du module et son empreinte SHA-256.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <param name="moduleName">Nom du module VBA ciblé.</param>
        /// <returns>Code du module et empreinte SHA-256.</returns>
        private object ReadModule(string projectName, string moduleName)
        {
            dynamic module = GetModule(projectName, moduleName);
            string code = GetCode(module);
            return new { Project = projectName, Module = moduleName, Code = code, Sha256 = Hash(code) };
        }

        /// <summary>Données sérialisables d’une référence de projet VBA.</summary>
        private sealed class ReferenceInfo
        {

            /// <summary>Nom de la référence lorsqu’il est accessible.</summary>
            /// <value>Nom lu depuis la référence VBE.</value>
            public string Name { get; set; }

            /// <summary>Identifiant GUID de la bibliothèque référencée.</summary>
            /// <value>GUID de la bibliothèque référencée.</value>
            public string Guid { get; set; }

            /// <summary>Version majeure de la référence.</summary>
            /// <value>Numéro de version majeure déclaré par le VBE.</value>
            public int Major { get; set; }

            /// <summary>Version mineure de la référence.</summary>
            /// <value>Numéro de version mineure déclaré par le VBE.</value>
            public int Minor { get; set; }

            /// <summary>Indique si le VBE signale une référence manquante.</summary>
            /// <value>État de résolution de la référence indiqué par le VBE.</value>
            public bool IsBroken { get; set; }

            /// <summary>Indique si la référence est intégrée au projet hôte.</summary>
            /// <value>Indique si le VBE classe la référence comme intégrée.</value>
            public bool BuiltIn { get; set; }

            /// <summary>Chemin du fichier de bibliothèque lorsqu’il est disponible.</summary>
            /// <value>Chemin de la bibliothèque lorsqu’il est résolu.</value>
            public string FullPath { get; set; }
        }

        /// <summary>Retourne les références du projet et leur empreinte de version.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <returns>Références du projet et empreinte de leur état.</returns>
        private object ListReferences(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = ReadReferences(project);
            return new { Project = projectName, Version = ReferencesVersion(result), References = result };
        }

        /// <summary>Lit les identités et états disponibles des références d’un projet.</summary>
        /// <param name="project">Projet VBE dont les références sont lues.</param>
        /// <returns>Références accessibles sous forme sérialisable.</returns>
        private static List<ReferenceInfo> ReadReferences(dynamic project)
        {
            var result = new List<ReferenceInfo>();
            foreach (dynamic reference in project.References)
            {
                bool broken = (bool)reference.IsBroken;
                string name = null;
                string fullPath = null;
                if (!broken)
                {
                    try { name = (string)reference.Name; } catch { }
                    try { fullPath = (string)reference.FullPath; } catch { }
                }
                result.Add(new ReferenceInfo { Name = name, Guid = (string)reference.GUID,
                    Major = (int)reference.Major, Minor = (int)reference.Minor,
                    IsBroken = broken, BuiltIn = (bool)reference.BuiltIn, FullPath = fullPath });
            }
            return result;
        }

        /// <summary>Calcule une empreinte stable des informations de références lues.</summary>
        /// <param name="references">Références à inclure dans l’empreinte.</param>
        /// <returns>Empreinte SHA-256 des informations de référence.</returns>
        private static string ReferencesVersion(List<ReferenceInfo> references)
        {
            var text = new StringBuilder();
            foreach (var reference in references)
            {
                text.Append(reference.Guid).Append('|').Append(reference.Major).Append('|')
                    .Append(reference.Minor).Append('|').Append(reference.IsBroken).Append('|')
                    .Append(reference.BuiltIn).Append('|')
                    .Append(reference.Name).Append('|').Append(reference.FullPath).Append('\n');
            }
            return Hash(text.ToString());
        }

        /// <summary>Vérifie la version des références et le mode conception avant modification.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>Projet en mode conception dont la version des références correspond.</returns>
        private dynamic CheckedReferenceProject(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedReferencesVersion))
                throw new ArgumentException("ExpectedReferencesVersion is required from list_references.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The project must be in design mode.");
            if (!string.Equals(ReferencesVersion(ReadReferences(project)), request.ExpectedReferencesVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Project references changed since they were read.");
            return project;
        }

        /// <summary>Ajoute une référence par GUID et versions demandées après validation anti-concurrence.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>État des références après l’ajout par GUID.</returns>
        private object AddReferenceGuid(Request request)
        {
            System.Guid parsed;
            if (!System.Guid.TryParse(request.Guid, out parsed) || request.Major < 0 || request.Minor < 0)
                throw new ArgumentException("Guid, nonnegative Major and Minor are required.");
            dynamic project = CheckedReferenceProject(request);
            if (((List<ReferenceInfo>)ReadReferences(project)).Any(item => string.Equals(item.Guid, parsed.ToString("B"),
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("A reference with this GUID is already selected.");
            project.References.AddFromGuid(parsed.ToString("B"), request.Major, request.Minor);
            return ListReferences(request.Project);
        }

        /// <summary>Ajoute au projet une bibliothèque depuis un chemin absolu existant.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>État des références après l’ajout du fichier.</returns>
        private object AddReferenceFile(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Path) ||
                !Regex.IsMatch(request.Path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("A fully qualified reference file path is required.");
            string path = Path.GetFullPath(request.Path);
            if (!File.Exists(path)) throw new FileNotFoundException("Reference file not found.", path);
            dynamic project = CheckedReferenceProject(request);
            project.References.AddFromFile(path);
            return ListReferences(request.Project);
        }

        /// <summary>Retire la référence correspondant exactement au GUID et aux versions spécifiés.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>État des références après le retrait exact.</returns>
        private object RemoveReference(Request request)
        {
            System.Guid parsed;
            if (!System.Guid.TryParse(request.Guid, out parsed) || request.Major < 0 || request.Minor < 0)
                throw new ArgumentException("Guid, nonnegative Major and Minor are required.");
            dynamic project = CheckedReferenceProject(request);
            dynamic target = null;
            foreach (dynamic reference in project.References)
                if (string.Equals((string)reference.GUID, parsed.ToString("B"), StringComparison.OrdinalIgnoreCase) &&
                    (int)reference.Major == request.Major && (int)reference.Minor == request.Minor)
                { target = reference; break; }
            if (target == null) throw new InvalidOperationException("The exact reference was not found.");
            if ((bool)target.BuiltIn)
                throw new InvalidOperationException("The VBE marks this reference as built in and non-removable.");
            project.References.Remove(target);
            return ListReferences(request.Project);
        }

        /// <summary>Crée un composant VBA du type demandé et vérifie son identité effective.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <param name="componentType">Type VBE du composant à créer.</param>
        /// <returns>Identité, code et empreinte du composant créé.</returns>
        private object CreateComponent(Request request, int componentType)
        {
            if (string.IsNullOrWhiteSpace(request.Module) ||
                !Regex.IsMatch(request.Module, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Module must start with a letter and contain at most 40 letters, digits or underscores.");
            if (request.ExpectedMode != 2)
                throw new ArgumentException("ExpectedMode must be 2 (design mode), obtained from list_projects.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != request.ExpectedMode)
                throw new InvalidOperationException("The project is no longer in design mode.");
            foreach (dynamic existing in project.VBComponents)
                if (string.Equals((string)existing.Name, request.Module, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");

            dynamic component = project.VBComponents.Add(componentType);
            try { component.Name = request.Module; }
            catch
            {
                // Only the newly created component is rolled back when its requested name is rejected.
                try { project.VBComponents.Remove(component); } catch { }
                throw;
            }
            string actualName = (string)component.Name;
            int actualType = (int)component.Type;
            if (!string.Equals(actualName, request.Module, StringComparison.Ordinal) || actualType != componentType)
                throw new InvalidOperationException("The VBE did not create the requested component identity.");
            dynamic module = component.CodeModule;
            string code = GetCode(module);
            return new { Project = request.Project, Module = actualName, Type = actualType,
                Lines = (int)module.CountOfLines, Code = code, Sha256 = Hash(code) };
        }

        /// <summary>Remplace une plage de lignes après vérification du mode, de l’empreinte et des bornes.</summary>
        /// <param name="request">Paramètres de la commande à exécuter.</param>
        /// <returns>Réponse décrivant l’empreinte et le nombre de lignes après remplacement.</returns>
        private Response ReplaceLines(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256))
                return Response.Failure("ExpectedSha256 is required for edits.");
            if (request.StartLine < 1 || request.Count < 0 || request.Text == null)
                return Response.Failure("Invalid line range or replacement text.");

            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2) // vbext_vm_Design
                return Response.Failure("The project must be in design mode before editing.");
            dynamic module = GetModule(request.Project, request.Module);
            string before = GetCode(module);
            if (!string.Equals(Hash(before), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                return Response.Failure("The module changed since it was read.");
            int lineCount = (int)module.CountOfLines;
            if (request.StartLine > lineCount + 1 || request.Count > lineCount - request.StartLine + 1)
                return Response.Failure("The requested line range is outside the module.");

            string after;
            try
            {
                if (request.Count > 0) module.DeleteLines(request.StartLine, request.Count);
                if (request.Text.Length > 0) module.InsertLines(request.StartLine, request.Text);
                after = GetCode(module);
            }
            catch (Exception error)
            {
                try
                {
                    if (!string.Equals(GetCode(module), before, StringComparison.Ordinal))
                    {
                        int remaining = (int)module.CountOfLines;
                        if (remaining > 0) module.DeleteLines(1, remaining);
                        if (before.Length > 0) module.InsertLines(1, before);
                    }
                    if (!string.Equals(GetCode(module), before, StringComparison.Ordinal))
                        throw new InvalidOperationException("Restored source does not match the original revision.");
                }
                catch (Exception rollback)
                {
                    return Response.Failure("Code edit failed: " + error.Message + ". Rollback failed: " + rollback.Message + ". Read the module before continuing.");
                }
                return Response.Failure("Code edit failed; original source restored: " + error.Message);
            }
            codeEdits.Record(request.Project, request.Module, before, after);
            return Response.Success(new { Sha256 = Hash(after), Lines = (int)module.CountOfLines });
        }

        /// <summary>Résout un projet VBE à partir de son nom.</summary>
        /// <param name="name">Nom de projet à résoudre.</param>
        /// <returns>Projet résolu par le résolveur VBE.</returns>
        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        /// <summary>Résout le module nommé sans tenir compte de la casse.</summary>
        /// <param name="projectName">Nom du projet VBE ciblé.</param>
        /// <param name="moduleName">Nom du module VBA ciblé.</param>
        /// <returns>Module de code correspondant au nom fourni.</returns>
        private dynamic GetModule(string projectName, string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module is required.");
            dynamic project = GetProject(projectName);
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                    return component.CodeModule;
            throw new InvalidOperationException("Module not found: " + moduleName);
        }

        /// <summary>Lit toutes les lignes du module, ou retourne une chaîne vide si celui-ci est vide.</summary>
        /// <param name="module">Module VBE à lire.</param>
        /// <returns>Texte complet du module, ou chaîne vide.</returns>
        private static string GetCode(dynamic module)
        {
            int count = (int)module.CountOfLines;
            return count == 0 ? string.Empty : (string)module.Lines[1, count];
        }

        /// <summary>Calcule l’empreinte SHA-256 UTF-8 du code fourni.</summary>
        /// <param name="code">Code source à hacher.</param>
        /// <returns>Empreinte SHA-256 hexadécimale en minuscules.</returns>
        private static string Hash(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code))).Replace("-", "").ToLowerInvariant();
        }
    }
}
