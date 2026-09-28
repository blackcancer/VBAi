using System;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Expose les opérations de session VBE par un canal nommé propre au processus.</summary>
    internal sealed class BridgeServer : IDisposable
    {
        /// <summary>Contrôle WinForms utilisé pour exécuter les appels COM sur le thread UI.</summary>
        private readonly Control dispatcher;
        /// <summary>Session qui traite les requêtes destinées au VBE.</summary>
        private readonly VbeSession session;
        /// <summary>Nom du canal nommé associé au processus hôte.</summary>
        private readonly string pipeName;
        /// <summary>Thread d’arrière-plan qui accepte les connexions du client.</summary>
        private readonly Thread worker;
        /// <summary>Indique que l’arrêt du serveur a été demandé.</summary>
        private volatile bool stopping;
        /// <summary>Connexion actuellement acceptée, fermée lors de l’arrêt.</summary>
        private NamedPipeServerStream listener;

        /// <summary>Crée le serveur IPC et prépare son thread d’écoute.</summary>
        /// <param name="dispatcher">Contrôle WinForms propriétaire du thread VBE.</param>
        /// <param name="session">Session utilisée pour exécuter les commandes.</param>
        /// <param name="processId">Identifiant du processus qui distingue le canal.</param>
        public BridgeServer(Control dispatcher, VbeSession session, int processId)
        {
            this.dispatcher = dispatcher;
            this.session = session;
            pipeName = "CodexVBE." + processId;
            worker = new Thread(Run) { IsBackground = true, Name = "CodexVBE pipe" };
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
                    security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,
                        PipeAccessRights.FullControl, AccessControlType.Allow));
                    using (var pipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1,
                        PipeTransmissionMode.Byte, PipeOptions.None, 4096, 4096, security))
                    {
                        listener = pipe;
                        pipe.WaitForConnection();
                        if (stopping) break;
                        using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                        using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        {
                            string line = reader.ReadLine();
                            var json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
                            Response response;
                            try
                            {
                                var request = json.Deserialize<Request>(line);
                                if (request != null && request.Command == "debug_windows")
                                    response = Response.Success(VbeDebugWindows.Capture(request.IncludeCallStack));
                                else if (request != null && request.Command == "debug_dialog")
                                    response = Response.Success(VbeDebugWindows.ReadDebugDialog());
                                else if (request != null && request.Command == "debug_item")
                                    response = Response.Success(VbeDebugWindows.ChangeDebugItem(request));
                                else if (request != null && request.Command == "respond_debug_dialog")
                                    response = Response.Success(VbeDebugWindows.RespondDebugDialog(request));
                                else if (request != null && request.Command == "immediate_execute")
                                {
                                    if (string.IsNullOrWhiteSpace(request.Project) ||
                                        (request.ExpectedMode != 1 && request.ExpectedMode != 2))
                                        throw new ArgumentException("Project and ExpectedMode (1 or 2) are required.");
                                    var state = (Response)dispatcher.Invoke(new Func<Response>(() =>
                                        session.Execute(new Request { Command = "debug_state", Project = request.Project })));
                                    if (!state.Ok) response = state;
                                    else if ((int)((dynamic)state.Data).Mode != request.ExpectedMode)
                                        response = Response.Failure("Project mode changed before Immediate execution.");
                                    else response = Response.Success(VbeDebugWindows.ExecuteImmediate(request.Text));
                                }
                                else if (request != null && request.Command == "compile_project")
                                {
                                    VbeDebugWindows.EnsureNoCompileDialog();
                                    Response compileResponse = null;
                                    var completed = new ManualResetEventSlim(false);
                                    // Keep the event alive if a timeout occurs while the UI
                                    // callback is still pending; its finally block will signal it.
                                    dispatcher.BeginInvoke(new Action(() => {
                                        try { compileResponse = session.Execute(request); }
                                        catch (Exception ex) { compileResponse = Response.Failure(ex.Message); }
                                        finally { completed.Set(); }
                                    }));
                                    string diagnostic = VbeDebugWindows.AwaitCompileDialog(completed);
                                    if (compileResponse == null)
                                        response = Response.Failure("The native Compile command did not return a result.");
                                    else if (!compileResponse.Ok)
                                        response = compileResponse;
                                    else response = Response.Success(new {
                                        Project = request.Project,
                                        Compiled = diagnostic == null,
                                        Diagnostic = diagnostic,
                                        Verification = diagnostic == null ? "NoNativeDiagnosticObserved" : "NativeDiagnosticCaptured",
                                        Command = compileResponse.Data,
                                        NextRead = diagnostic == null ? null :
                                            "Read debug_state and the active code pane in a separate request to locate the failed statement."
                                    });
                                }
                                else if (request != null && request.Command == "add_watch")
                                {
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(VbeDebugWindows.CompleteAddWatch(request));
                                }
                                else if (request != null && request.Command == "edit_watch")
                                {
                                    VbeDebugWindows.SelectWatch(request);
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(VbeDebugWindows.CompleteEditWatch(request));
                                }
                                else if (request != null && request.Command == "quick_watch")
                                {
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(VbeDebugWindows.CompleteQuickWatch(request));
                                }
                                else if (request != null &&
                                    (request.Command == "read_debug_options" || request.Command == "read_vbe_options"))
                                {
                                    VbeDebugWindows.EnsureNoDebugOptionsDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(request.Command == "read_vbe_options"
                                        ? VbeDebugWindows.ReadVbeOptions() : VbeDebugWindows.ReadDebugOptions());
                                }
                                else if (request != null && request.Command == "read_project_signature_dialog")
                                {
                                    VbeDebugWindows.EnsureNoSignatureDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(
                                        VbeDebugWindows.ReadSignatureDialog(request.Project));
                                }
                                else if (request != null && request.Command == "sign_project")
                                {
                                    VbeDebugWindows.EnsureNoSignatureDialog();
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok)
                                    {
                                        string certificateName = (string)((dynamic)response.Data).CertificateName;
                                        bool unsignedVerified = (bool)((dynamic)response.Data).UnsignedVerified;
                                        object signed = VbeDebugWindows.CompleteProjectSignature(request.Project,
                                            request.CertificateThumbprint, certificateName, unsignedVerified);
                                        object persistence = null;
                                        string persistenceError = null;
                                        for (int attempt = 0; attempt < 12; attempt++)
                                        {
                                            try
                                            {
                                                persistence = dispatcher.Invoke(new Func<object>(() =>
                                                    session.PersistProjectSignature(request.Project)));
                                                persistenceError = null;
                                                break;
                                            }
                                            catch (Exception ex)
                                            {
                                                persistenceError = ex.Message;
                                                if (attempt == 11 || ex.ToString().IndexOf("0x800AC472",
                                                    StringComparison.OrdinalIgnoreCase) < 0) break;
                                                System.Threading.Thread.Sleep(250);
                                            }
                                        }
                                        Response status = (Response)dispatcher.Invoke(new Func<Response>(() =>
                                            session.Execute(new Request { Command = "project_signature_status", Project = request.Project })));
                                        response = Response.Success(new { Signature = signed,
                                            Persistence = persistence, PersistenceError = persistenceError,
                                            SaveRequired = persistence == null || !((bool)((dynamic)persistence).Saved),
                                            HostStatus = status.Ok ? status.Data : null,
                                            HostStatusError = status.Ok ? null : status.Error });
                                    }
                                }
                                else if (request != null && request.Command == "remove_watch")
                                {
                                    VbeDebugWindows.SelectWatch(request);
                                    response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                                    if (response.Ok) response = Response.Success(VbeDebugWindows.VerifyWatchRemoved(request));
                                }
                                else response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
                            }
                            catch (Exception ex)
                            {
                                response = Response.Failure(ex.Message);
                            }
                            writer.WriteLine(json.Serialize(response));
                        }
                    }
                }
                catch (IOException) { if (stopping) break; }
                catch (ObjectDisposedException) { break; }
                finally { listener = null; }
            }
        }

        /// <summary>Demande l’arrêt de l’écoute et ferme la connexion en cours pour débloquer l’attente.</summary>
        public void Dispose()
        {
            stopping = true;
            listener?.Dispose();
        }
    }
}
