using System;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Frontières natives du débogueur, conservant leurs implémentations VBE par défaut.</summary>
    internal sealed class VbeToolNativeBoundary
    {

        /// <summary>Lit la liste native de l’explorateur d’objets pour les critères demandés.</summary>
        internal Func<Request, object> ListObjectBrowser = VbeDebugWindows.ListObjectBrowser;

        /// <summary>Lit les nœuds de navigation natifs sur le worker d'accessibilité.</summary>
        internal Func<Request, object> ReadNavigationSurface = VbeDebugWindows.ReadNavigationSurface;

        /// <summary>Refuse de réutiliser un dialogue de propriétés déjà ouvert.</summary>
        internal Action EnsureNoProjectPropertiesDialog = VbeDebugWindows.EnsureNoProjectPropertiesDialog;

        /// <summary>Lit la protection native sans restituer de secret.</summary>
        internal Func<Request, object> ReadProjectProtection = VbeDebugWindows.ReadProjectProtection;

        /// <summary>Configure la protection via le dialogue natif exact.</summary>
        internal Func<Request, object> SetProjectProtection = VbeDebugWindows.SetProjectProtection;

        /// <summary>Sélectionne ou développe un nœud de navigation identifié.</summary>
        internal Func<Request, object> ChangeNavigationSurface = VbeDebugWindows.ChangeNavigationSurface;

        /// <summary>Sélectionne une entrée native de l’explorateur d’objets.</summary>
        internal Func<Request, object> SelectObjectBrowser = VbeDebugWindows.SelectObjectBrowser;

        /// <summary>Inspecte l’explorateur d’objets par son fournisseur d’automatisation Windows.</summary>
        internal Func<object> ReadObjectBrowser = VbeDebugWindows.ReadObjectBrowser;

        /// <summary>Inspecte les formulaires VBA actifs par l’automatisation Windows.</summary>
        internal Func<object> ReadRuntimeForms = VbeDebugWindows.ReadRuntimeForms;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<bool, object> Capture = VbeDebugWindows.Capture;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<object> ReadDebugDialog = VbeDebugWindows.ReadDebugDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> ChangeDebugItem = VbeDebugWindows.ChangeDebugItem;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> RespondDebugDialog = VbeDebugWindows.RespondDebugDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<string, Action<Action>, object> ExecuteImmediate = VbeDebugWindows.ExecuteImmediateGuarded;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Action EnsureNoCompileDialog = VbeDebugWindows.EnsureNoCompileDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<ManualResetEventSlim, string> AwaitCompileDialog = VbeDebugWindows.AwaitCompileDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> CompleteAddWatch = VbeDebugWindows.CompleteAddWatch;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> SelectWatch = VbeDebugWindows.SelectWatch;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> CompleteEditWatch = VbeDebugWindows.CompleteEditWatch;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> CompleteQuickWatch = VbeDebugWindows.CompleteQuickWatch;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Action EnsureNoDebugOptionsDialog = VbeDebugWindows.EnsureNoDebugOptionsDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<object> ReadVbeOptions = VbeDebugWindows.ReadVbeOptions;

        /// <summary>Écrit seulement les préférences natives d’édition/débogage reconnues.</summary>
        internal Func<Request, object> SetVbeOption = VbeDebugWindows.SetVbeOption;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<object> ReadDebugOptions = VbeDebugWindows.ReadDebugOptions;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Action EnsureNoSignatureDialog = VbeDebugWindows.EnsureNoSignatureDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<string, object> ReadSignatureDialog = VbeDebugWindows.ReadSignatureDialog;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<string, string, string, bool, object> CompleteProjectSignature = VbeDebugWindows.CompleteProjectSignature;

        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<Request, object> VerifyWatchRemoved = VbeDebugWindows.VerifyWatchRemoved;
    }

    /// <summary>Expose les opérations de session VBE par un canal nommé propre au processus.</summary>
    internal sealed class BridgeServer : IDisposable
    {

        /// <summary>Contrôle WinForms utilisé pour exécuter les appels COM sur le thread UI.</summary>
        private readonly Control dispatcher;

        /// <summary>Session qui traite les requêtes destinées au VBE.</summary>
        private readonly VbeSession session;

        /// <summary>Host-only qualification opt-in, captured at connection and absent from the LLM catalogue.</summary>
        private readonly PathVisibilityDiagnostic pathVisibility;

        /// <summary>Explicit disposable owner-Git diagnostic, never exposed to the LLM catalogue.</summary>
        private readonly OwnerGitQualification ownerGitQualification;

        /// <summary>Adaptateurs natifs du débogueur, remplaçables par instance à la frontière UI.</summary>
        internal readonly VbeToolNativeBoundary Native = new VbeToolNativeBoundary();

        /// <summary>Exécute une commande sur la session hôte, sans remplacer l’orchestration de l’outil.</summary>
        internal Func<Request, Response> Execute;

        /// <summary>Captures Immediate output while yielding to the owning VBE message loop.</summary>
        internal Func<Request, Task<object>> ReadImmediateNative;

        /// <summary>Inspects declared scalar locals on the owning VBE UI thread.</summary>
        internal Func<Request, Task<object>> InspectLocalScalarsNative;

        /// <summary>Saves and observes completion while yielding to the owning VBE STA.</summary>
        internal Func<Request, Task<object>> SaveHostDocumentNative;

        /// <summary>Optional asynchronous adapter for creating the approved SOLIDWORKS macro from the VBE UI thread.</summary>
        internal Func<Request, Task<object>> SolidWorksMacroNative;

        /// <summary>Runs the guarded native General operation on the VBE UI thread.</summary>
        internal Func<Request, bool, Task<object>> ProjectGeneralNative;

        /// <summary>Demande la sauvegarde de signature au document hôte.</summary>
        internal Func<string, Action, object> PersistSignature;
        /// <summary>Captures owner-thread project identity/content before deferred certificate selection.</summary>
        internal Func<string, Action> CaptureSignaturePersistence;

        /// <summary>Crée le canal local avec la sécurité de l’utilisateur courant.</summary>
        internal Func<PipeSecurity, NamedPipeServerStream> OpenPipe;

        /// <summary>Nom du canal nommé associé au processus hôte.</summary>
        private readonly string pipeName;

        /// <summary>Thread d’arrière-plan qui accepte les connexions du client.</summary>
        private readonly Thread worker;

        /// <summary>Indique que l’arrêt du serveur a été demandé.</summary>
        private volatile bool stopping;

        /// <summary>Connexion actuellement acceptée, fermée lors de l’arrêt.</summary>
        private NamedPipeServerStream listener;

        /// <summary>Budget de réception d'une requête complète, indépendant de la durée d'exécution VBE.</summary>
        internal TimeSpan RequestReadTimeout = TimeSpan.FromSeconds(10);

        /// <summary>Limits response delivery independently of native command execution.</summary>
        internal TimeSpan ResponseWriteTimeout = TimeSpan.FromSeconds(10);

        /// <summary>Limite UTF-8 appliquée avant l'allocation de la ligne JSON complète.</summary>
        internal int MaxRequestBytes = 10 * 1024 * 1024;

        /// <summary>Crée le serveur IPC et prépare son thread d’écoute.</summary>
        /// <param name="dispatcher">Contrôle WinForms propriétaire du thread VBE.</param>
        /// <param name="session">Session utilisée pour exécuter les commandes.</param>
        /// <param name="processId">Identifiant du processus qui distingue le canal.</param>
        public BridgeServer(Control dispatcher, VbeSession session, int processId)
        {
            this.dispatcher = dispatcher;
            this.session = session;
            pathVisibility = new PathVisibilityDiagnostic(processId);
            ownerGitQualification = new OwnerGitQualification(session, processId);
            Execute = request => session.Execute(request);
            ReadImmediateNative = request => session.ReadImmediateAsync(request);
            InspectLocalScalarsNative = request => session.InspectLocalScalarsAsync(request);
            SaveHostDocumentNative = request => session.SaveHostDocumentAsync(request);
            SolidWorksMacroNative = request => session.SolidWorksMacroAsync(request);
            ProjectGeneralNative = (request, write) => session.ProjectGeneralAsync(request, write);
            PersistSignature = (project, authorize) => session.PersistProjectSignature(project, authorize);
            CaptureSignaturePersistence = project => session.CaptureSignaturePersistence(project);
            pipeName = "VBAi." + processId;
            OpenPipe = security => new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous, 4096, 4096, security);
            worker = new Thread(Run) { IsBackground = true, Name = "VBAi pipe" };
        }

        /// <summary>Démarre le thread qui accepte les requêtes du client local.</summary>
        public void Start() { worker.Start(); }

        /// <summary>Accepte les requêtes, les distribue et renvoie une réponse JSON par connexion.</summary>
        private void Run()
        {
            while (!stopping)
            {
                try
                {
                    var security = new PipeSecurity();
                    security.AddAccessRule(new PipeAccessRule(new SecurityIdentifier(WellKnownSidType.NetworkSid, null),
                        PipeAccessRights.FullControl, AccessControlType.Deny));
                    security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,
                        PipeAccessRights.FullControl, AccessControlType.Allow));
                    using (var pipe = OpenPipe(security))
                    {
                        listener = pipe;
                        pipe.WaitForConnection();
                        if (stopping) break;
                        {
                            string line = BridgeRequestReader.ReadAsync(pipe, MaxRequestBytes, RequestReadTimeout).GetAwaiter().GetResult();
                            var json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
                            Response response;
                            Action releaseAdmission = null;
                            try
                            {
                                var request = json.Deserialize<Request>(line);
                                releaseAdmission = AdmitSessionRequest(request);
                                if (request != null && request.Command == PathVisibilityDiagnostic.CommandName)
                                {
                                    PathVisibilityDiagnostic.RequireParameterFree(line);
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Response.Success(pathVisibility.Read())));
                                }
                                else if (request != null && request.Command == OwnerGitQualificationManifest.CommandName)
                                {
                                    OwnerGitQualificationManifest.RequireExactRequest(line);
                                    var trace = VbeInspectionTrace.Begin();
                                    trace?.Record(VbeInspectionTrace.Phase.Enqueue);
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    dispatcher.BeginInvoke(new Action(async () =>
                                    {
                                        try
                                        {
                                            var result = await OwnerGitQualification.RunOnOwnerAsync(
                                                () => ownerGitQualification.ExecuteAsync(request, line), trace);
                                            completion.TrySetResult(Response.Success(result));
                                        }
                                        catch (Exception ex) { completion.TrySetResult(Response.Failure(ex.ToString())); }
                                    }));
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && request.Command == "debug_windows")
                                    response = Response.Success(Native.Capture(request.IncludeCallStack));
                                else if (request != null && request.Command == "read_navigation_surface")
                                    response = Response.Success(Native.ReadNavigationSurface(request));
                                else if (request != null && request.Command == "change_navigation_surface")
                                    response = Response.Success(Native.ChangeNavigationSurface(request));
                                else if (request != null && request.Command == "list_object_browser")
                                    response = Response.Success(Native.ListObjectBrowser(request));
                                else if (request != null && request.Command == "select_object_browser")
                                    response = Response.Success(Native.SelectObjectBrowser(request));
                                else if (request != null && request.Command == "read_runtime_forms")
                                    response = Response.Success(Native.ReadRuntimeForms());
                                else if (request != null && request.Command == "read_object_browser")
                                    response = Response.Success(Native.ReadObjectBrowser());
                                else if (request != null && request.Command == "debug_dialog")
                                    response = Response.Success(Native.ReadDebugDialog());
                                else if (request != null && request.Command == "debug_item")
                                    response = Response.Success(Native.ChangeDebugItem(request));
                                else if (request != null && request.Command == "respond_debug_dialog")
                                    response = Response.Success(Native.RespondDebugDialog(request));
                                else if (request != null && request.Command == "immediate_execute")
                                {
                                    if (string.IsNullOrWhiteSpace(request.Project) ||
                                        (request.ExpectedMode != 1 && request.ExpectedMode != 2))
                                        throw new ArgumentException("Project and ExpectedMode (1 or 2) are required.");
                                    var state = (Response)dispatcher.Invoke(new Func<Response>(() =>
                                        Execute(new Request { Command = "debug_state", Project = request.Project })));
                                    if (!state.Ok) response = state;
                                    else if ((int)((dynamic)state.Data).Mode != request.ExpectedMode)
                                        response = Response.Failure("Project mode changed before Immediate execution.");
                                    else
                                    {
                                        VbeImmediateContext.RequireProject(request.Project, state.Data);
                                        response = Response.Success(Native.ExecuteImmediate(request.Text, enter => dispatcher.Invoke(new Action(() =>
                                        {
                                            VbeImmediateContext.RequireCurrent(request.Project, request.ExpectedMode, Execute);
                                            enter();
                                        }))));
                                    }
                                }
                                else if (request != null && (request.Command == "read_project_general" || request.Command == "set_project_general"))
                                {
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    dispatcher.BeginInvoke(new Action(async () =>
                                    {
                                        int ownerThread = Thread.CurrentThread.ManagedThreadId;
                                        string command = request.Command, project = request.Project;
                                        string version = request.ExpectedProjectVersion, options = request.ExpectedOptionsVersion;
                                        string property = request.Property, caption = request.ControlCaption;
                                        object value = request.Value;
                                        request.RevalidateProjectPropertyAuthorization = live =>
                                        {
                                            if (dispatcher.IsDisposed || !dispatcher.IsHandleCreated || dispatcher.InvokeRequired ||
                                                Thread.CurrentThread.ManagedThreadId != ownerThread ||
                                                Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                                                request.Command != command || request.Project != project ||
                                                request.ExpectedProjectVersion != version || request.ExpectedOptionsVersion != options || request.ExpectedMode != 2 ||
                                                request.Property != property || request.ControlCaption != caption || !object.Equals(request.Value, value))
                                                throw new InvalidOperationException("Original bridge General request/UI context changed.");
                                        };
                                        try { completion.TrySetResult(Response.Success(await ProjectGeneralNative(request, command == "set_project_general"))); }
                                        catch (Exception ex) { completion.TrySetResult(Response.Failure(ex.Message)); }
                                        finally { request.RevalidateProjectPropertyAuthorization = null; }
                                    }));
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && (request.Command == "create_solidworks_macro" || request.Command == "publish_solidworks_macro"))
                                {
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    dispatcher.BeginInvoke(new Action(async () =>
                                    {
                                        int ownerThread = Thread.CurrentThread.ManagedThreadId;
                                        string command = request.Command, project = request.Project, path = request.Path;
                                        string version = request.ExpectedProjectVersion;
                                        int mode = request.ExpectedMode;
                                        request.RevalidateMacroAuthorization = live =>
                                        {
                                            if (dispatcher.IsDisposed || !dispatcher.IsHandleCreated || dispatcher.InvokeRequired ||
                                                Thread.CurrentThread.ManagedThreadId != ownerThread ||
                                                Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                                                request.Command != command || request.Project != project || request.Path != path ||
                                                request.ExpectedProjectVersion != version || request.ExpectedMode != mode || mode != 2)
                                                throw new InvalidOperationException("Original bridge macro request/UI context changed.");
                                        };
                                        try { completion.TrySetResult(Response.Success(await SolidWorksMacroNative(request))); }
                                        catch (Exception ex) { completion.TrySetResult(Response.Failure(ex.Message)); }
                                        finally { request.RevalidateMacroAuthorization = null; }
                                    }));
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && request.Command == "save_host_document")
                                {
                                    // Defer on the owning STA outside the current native/UI callback.
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    dispatcher.BeginInvoke(new Action(async () =>
                                    {
                                        try { completion.TrySetResult(Response.Success(await SaveHostDocumentNative(request))); }
                                        catch (Exception ex) { completion.TrySetResult(Response.Failure(ex.Message)); }
                                    }));
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && request.Command == "read_immediate")
                                {
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    dispatcher.BeginInvoke(new Action(async () =>
                                    {
                                        try { completion.TrySetResult(Response.Success(await ReadImmediateNative(request))); }
                                        catch (Exception ex) { completion.TrySetResult(Response.Failure(ex.Message)); }
                                    }));
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && request.Command == "inspect_local_scalars")
                                {
                                    var trace = VbeInspectionTrace.Begin();
                                    trace?.Record(VbeInspectionTrace.Phase.Enqueue);
                                    var completion = new TaskCompletionSource<Response>(TaskCreationOptions.RunContinuationsAsynchronously);
                                    try
                                    {
                                        dispatcher.BeginInvoke(new Action(async () =>
                                        {
                                            using (trace?.Enter())
                                            {
                                                trace?.Record(VbeInspectionTrace.Phase.CallbackEntered);
                                                Exception failure = null;
                                                try { completion.TrySetResult(Response.Success(await InspectLocalScalarsNative(request))); }
                                                catch (Exception ex) { failure = ex; completion.TrySetResult(Response.Failure(ex.Message)); }
                                                finally { trace?.Record(VbeInspectionTrace.Phase.Terminal, failure); }
                                            }
                                        }));
                                    }
                                    catch (Exception ex) { trace?.Record(VbeInspectionTrace.Phase.Terminal, ex); throw; }
                                    response = completion.Task.GetAwaiter().GetResult();
                                }
                                else if (request != null && request.Command == "compile_project")
                                {
                                    Native.EnsureNoCompileDialog();
                                    Response compileResponse = null;
                                    var completed = new ManualResetEventSlim(false);
                                    var admission = new VbeCompilationAdmission();
                                    // Retain the event until its callback returns, including cancellation.
                                    dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        try
                                        {
                                            if (admission.TryBegin()) compileResponse = Execute(request);
                                        }
                                        catch (Exception ex) { compileResponse = Response.Failure(ex.Message); }
                                        finally { completed.Set(); }
                                    }));
                                    string diagnostic;
                                    try { diagnostic = Native.AwaitCompileDialog(completed); }
                                    finally { admission.CancelPending(); }
                                    if (compileResponse == null)
                                        response = Response.Failure("The native Compile command did not return a result.");
                                    else if (!compileResponse.Ok)
                                        response = compileResponse;
                                    else response = Response.Success(VbeCompilationResult.Create(request.Project, compileResponse.Data, diagnostic));
                                }
                                else if (request != null && request.Command == "add_watch")
                                {
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(Native.CompleteAddWatch(request));
                                }
                                else if (request != null && request.Command == "edit_watch")
                                {
                                    Native.SelectWatch(request);
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(Native.CompleteEditWatch(request));
                                }
                                else if (request != null && request.Command == "quick_watch")
                                {
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(Native.CompleteQuickWatch(request));
                                }
                                else if (request != null &&
                                    (request.Command == "read_debug_options" || request.Command == "read_vbe_options" || request.Command == "set_vbe_option"))
                                {
                                    Native.EnsureNoDebugOptionsDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(request.Command == "set_vbe_option" ? Native.SetVbeOption(request) : request.Command == "read_vbe_options"
                                        ? Native.ReadVbeOptions() : Native.ReadDebugOptions());
                                }
                                else if (request != null && (request.Command == "read_project_protection" || request.Command == "set_project_protection"))
                                {
                                    Native.EnsureNoProjectPropertiesDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok)
                                    {
                                        var captured = json.Deserialize<Request>(json.Serialize(request));
                                        captured.Caption = (string)((dynamic)response.Data).ProjectName;
                                        response = Response.Success(request.Command == "set_project_protection" ?
                                            Native.SetProjectProtection(captured) : Native.ReadProjectProtection(captured));
                                    }
                                }
                                else if (request != null && request.Command == "read_project_signature_dialog")
                                {
                                    Native.EnsureNoSignatureDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(
                                        Native.ReadSignatureDialog(request.Project));
                                }
                                else if (request != null && request.Command == "sign_project")
                                {
                                    Native.EnsureNoSignatureDialog();
                                    Action revalidatePersistence = null;
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() =>
                                    {
                                        revalidatePersistence = CaptureSignaturePersistence(request.Project);
                                        return Execute(request);
                                    }));
                                    if (response.Ok)
                                    {
                                        string certificateName = (string)((dynamic)response.Data).CertificateName;
                                        bool unsignedVerified = (bool)((dynamic)response.Data).UnsignedVerified;
                                        object signed = Native.CompleteProjectSignature(request.Project,
                                            request.CertificateThumbprint, certificateName, unsignedVerified);
                                        object persistence = null;
                                        string persistenceError = null;
                                        bool persistenceAuthorized = false;
                                        try
                                        {
                                            persistence = dispatcher.Invoke(new Func<object>(() =>
                                            {
                                                // A failed observation may follow a completed Save. Never replay it.
                                                return PersistSignature(request.Project, () =>
                                                {
                                                    revalidatePersistence();
                                                    persistenceAuthorized = true;
                                                });
                                            }));
                                        }
                                        catch (Exception ex) { persistenceError = ex.Message; }
                                        Response status;
                                        try
                                        {
                                            status = !persistenceAuthorized ? Response.Failure(persistenceError) : (Response)dispatcher.Invoke(new Func<Response>(() =>
                                            {
                                                revalidatePersistence();
                                                return Execute(new Request { Command = "project_signature_status", Project = request.Project });
                                            }));
                                        }
                                        catch (Exception ex) { status = Response.Failure(ex.Message); }
                                        response = Response.Success(new
                                        {
                                            Signature = signed,
                                            Persistence = persistence,
                                            PersistenceError = persistenceError,
                                            SaveRequired = persistence == null || !((bool)((dynamic)persistence).Saved),
                                            HostStatus = status.Ok ? status.Data : null,
                                            HostStatusError = status.Ok ? null : status.Error
                                        });
                                    }
                                }
                                else if (request != null && request.Command == "remove_watch")
                                {
                                    Native.SelectWatch(request);
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                                    if (response.Ok) response = Response.Success(Native.VerifyWatchRemoved(request));
                                }
                                else response = (Response)dispatcher.Invoke(new Func<Response>(() => Execute(request)));
                            }
                            catch (Exception ex)
                            {
                                response = Response.Failure(VbeScalarProperty.FormatFailure(ex));
                            }
                            finally
                            {
                                if (releaseAdmission != null) dispatcher.Invoke(releaseAdmission);
                            }
                            string payload;
                            try { payload = json.Serialize(response); }
                            catch (Exception error)
                            {
                                LoadLog.Write("Bridge response serialization failed: " + error.GetType().Name);
                                payload = json.Serialize(Response.Failure("The command response could not be serialized. The command may already have completed; inspect its state before retrying."));
                            }
                            BridgeResponseWriter.WriteAsync(pipe, payload, ResponseWriteTimeout).GetAwaiter().GetResult();
                        }
                    }
                }
                catch (IOException) { if (stopping) break; }
                catch (ObjectDisposedException) { break; }
                finally { listener = null; }
            }
        }

        // Native worker routes bypass Execute. Hold their session admission until
        // dispatch settles, so General cannot enter between the STA check and a
        // worker native call. General claims its own in-flight state on the STA.
        /// <summary>Reserves session admission for worker native routes after checking the VBE STA and that General-page work has settled.</summary>
        /// <param name="request">Request whose command determines whether it needs a session admission reservation.</param>
        /// <returns>Release action for a reserved bridge operation, or null for status and routes that claim their own state.</returns>
        private Action AdmitSessionRequest(Request request)
        {
            if (request?.Command == "status" || session == null) return null;
            return (Action)dispatcher.Invoke(new Func<Action>(() =>
            {
                if (dispatcher.IsDisposed || !dispatcher.IsHandleCreated || dispatcher.InvokeRequired ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Bridge admission requires its owning VBE UI STA.");
                session.RequireGeneralSettled();
                if (request?.Command == "read_project_general" || request?.Command == "set_project_general" ||
                    request?.Command == "create_solidworks_macro" || request?.Command == "publish_solidworks_macro") return null;
                return session.AdmitBridgeOperation();
            }));
        }

        /// <summary>Demande l’arrêt de l’écoute et ferme la connexion en cours pour débloquer l’attente.</summary>
        public void Dispose()
        {
            stopping = true;
            listener?.Dispose();
        }
    }
}
