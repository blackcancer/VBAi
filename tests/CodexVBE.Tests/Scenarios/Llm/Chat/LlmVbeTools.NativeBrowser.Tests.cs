using System;
using System.Threading.Tasks;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [TestMethod]
        public async Task BrowserToolsValidateCompleteArgumentsAndDispatchToNativeBoundaries()
        {
            var tools = new ToolFixture().Tools;
            bool fail = false; int calls = 0;
            Func<object> native = () => { calls++; if (fail) throw new InvalidOperationException("native unavailable"); return new object(); };
            tools.Native.ListObjectBrowser = r => { Assert.AreEqual("classes", r.Pane); return native(); };
            tools.Native.SelectObjectBrowser = r => native();
            tools.Native.ReadObjectBrowser = native; tools.Native.ReadRuntimeForms = native;
            foreach (var command in new[] { "list_object_browser", "select_object_browser", "read_object_browser", "read_runtime_forms" })
            {
                string args = command == "list_object_browser" ? "{\"Pane\":\"classes\",\"Query\":\"a\",\"Offset\":0,\"Limit\":2}" : command == "select_object_browser" ? "{\"ObjectName\":\"Alpha\",\"Procedure\":\"Go\",\"Context\":\"VBA\"}" : "{}";
                Success(await tools.InvokeAsync(command, args), command);
                fail = true; Failed(await tools.InvokeAsync(command, args), command); fail = false;
            }
            Success(await tools.InvokeAsync("select_object_browser", "{\"Context\":\"VBA\"}"), "library");
            int before = calls;
            foreach (string args in new[] { "null", "{}", "{\"Pane\":\"classes\",\"Unknown\":0}", "{\"Pane\":2}", "{\"Pane\":\"classes\",\"Query\":0}", "{\"Pane\":\"classes\",\"Offset\":\"0\"}", "{\"Pane\":\"classes\",\"Limit\":\"2\"}" })
                Failed(await tools.InvokeAsync("list_object_browser", args), "invalid browser list");
            foreach (string args in new[] { "null", "{}", "{\"ObjectName\":\"Alpha\",\"Unknown\":0}", "{\"ObjectName\":2}" })
                Failed(await tools.InvokeAsync("select_object_browser", args), "invalid browser selection");
            foreach (string command in new[] { "read_object_browser", "read_runtime_forms" })
            foreach (string args in new[] { "null", "{\"Unknown\":0}" })
                Failed(await tools.InvokeAsync(command, args), "invalid browser read");
            Assert.AreEqual(before, calls);
        }
    }
}
