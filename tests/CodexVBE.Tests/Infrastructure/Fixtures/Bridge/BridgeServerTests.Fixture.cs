namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Pipes;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class BridgeServerTests
    {
        private static IDictionary<string, object> SendWithoutMessagePump(int processId, string request)
        {
            using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
            {
                pipe.Connect(5000);
                using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                {
                    writer.WriteLine(request);
                    return (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(reader.ReadLine());
                }
            }
        }

        private static IDictionary<string, object> SendWithMessagePump(int processId, string request)
        {
            var pending = Task.Run(() =>
            {
                using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                {
                    pipe.Connect(5000);
                    using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true)
                    {
                        AutoFlush = true
                    }

                    )
                    using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                    {
                        writer.WriteLine(request);
                        return (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(reader.ReadLine());
                    }
                }
            });
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (!pending.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(5);
            }

            Assert.IsTrue(pending.IsCompleted, "The VBE bridge did not complete a pipe request.");
            return pending.GetAwaiter().GetResult();
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; } = 2;
        }
    }
}
