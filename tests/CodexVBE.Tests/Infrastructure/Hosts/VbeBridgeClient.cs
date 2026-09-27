using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace CodexVBE.Tests.Integration
{
    internal static class VbeBridgeClient
    {
        internal static IDictionary<string, object> Object(object value)
        {
            return (IDictionary<string, object>)value;
        }

        internal static IDictionary<string, object> Read(int processId, string command)
        {
            return Read(processId, new { Command = command });
        }

        internal static IDictionary<string, object> Read(int processId, object request)
        {
            var json = new JavaScriptSerializer();
            for (int attempt = 0; attempt < 20; attempt++)
            {
                try
                {
                    using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                    {
                        pipe.Connect(250);
                        using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                        {
                            writer.WriteLine(json.Serialize(request));
                            return Object(json.DeserializeObject(reader.ReadLine()));
                        }
                    }
                }
                catch (TimeoutException) { Thread.Sleep(250); }
                catch (IOException) { Thread.Sleep(250); }
            }
            return null;
        }
    }
}
