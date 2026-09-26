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
                                response = (Response)dispatcher.Invoke(new Func<Response>(() => session.Execute(request)));
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
