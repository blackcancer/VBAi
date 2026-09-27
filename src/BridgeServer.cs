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
    internal sealed class BridgeServer : IDisposable
    {
        private readonly Control dispatcher;
        private readonly VbeSession session;
        private readonly string pipeName;
        private readonly Thread worker;
        private volatile bool stopping;
        private NamedPipeServerStream listener;

        public BridgeServer(Control dispatcher, VbeSession session, int processId)
        {
            this.dispatcher = dispatcher;
            this.session = session;
            pipeName = "CodexVBE." + processId;
            worker = new Thread(Run) { IsBackground = true, Name = "CodexVBE pipe" };
        }

        public void Start() { worker.Start(); }

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
                                else if (request != null && request.Command == "respond_debug_dialog")
                                    response = Response.Success(VbeDebugWindows.RespondDebugDialog(request));
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

        public void Dispose()
        {
            stopping = true;
            listener?.Dispose();
        }
    }
}
