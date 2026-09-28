using System;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class BridgeServerTests
    {
        [STATestMethod]
        public void BrowserAndRuntimeFormDispatchPreserveResultsAndNativeFailuresOverTheRealPipe()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    bool fail = false; string observed = null;
                    Func<string, object> native = name => { if (fail) throw new InvalidOperationException("native unavailable"); observed = name; return new { Native = name }; };
                    server.Native.ListObjectBrowser = r => native("list:" + r.Pane);
                    server.Native.SelectObjectBrowser = r => native("select:" + r.ObjectName);
                    server.Native.ReadObjectBrowser = () => native("read");
                    server.Native.ReadRuntimeForms = () => native("forms");
                    server.Start(); var json = new JavaScriptSerializer();
                    foreach (string command in new[] { "list_object_browser", "select_object_browser", "read_object_browser", "read_runtime_forms" })
                    {
                        var response = SendWithMessagePump(id, json.Serialize(new { Command = command, Pane = "classes", ObjectName = "Alpha" }));
                        Assert.AreEqual(true, response["Ok"]); Assert.IsNotNull(observed);
                        fail = true;
                        response = SendWithMessagePump(id, json.Serialize(new { Command = command }));
                        Assert.AreEqual(false, response["Ok"]); StringAssert.Contains((string)response["Error"], "native unavailable");
                        fail = false;
                    }
                }
            }
        }
    }
}
