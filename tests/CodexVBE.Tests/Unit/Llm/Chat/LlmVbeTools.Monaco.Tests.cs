using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Threading.Tasks;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmMonacoValidationTests
    {
        [STATestMethod]
        public void MonacoDiffFailuresAreLoggedAfterSynchronizationWithoutDroppingAppliedSource()
        {
            foreach (string fault in new[] { "no subscriber", "subscriber throws", "readback throws" })
            using (var fixture = new CodexVBE.Tests.Unit.Editor.ModernEditorToolFixture())
            {
                int reads = 0, logs = 0;
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" }) { MonacoWindow = create => fixture.Window, MonacoModule = (p, m) => fixture.Module };
                tools.Execute = request => {
                    Assert.AreEqual("read_module", request.Command); reads++;
                    if (fault == "readback throws" && reads > 1) throw new InvalidOperationException("Diff unavailable");
                    return Response.Success(new CodexVBE.Tests.Infrastructure.VbeToolCodeResult { Code = fixture.Module.Code, Sha256 = EditorDocument.Hash(fixture.Module.Code) });
                };
                tools.WriteLog = message => { logs++; StringAssert.Contains(message, "Monaco code diff readback failed"); };
                if (fault == "subscriber throws") tools.CodeEdited += change => { throw new InvalidOperationException("Diff consumer unavailable"); };
                string updated = fixture.Document.Text + "\n' changed by owned Monaco contract";
                string arguments = new JavaScriptSerializer().Serialize(new { Project = "P", Module = "M", ExpectedVersion = 1, Text = updated });
                var response = new JavaScriptSerializer().Deserialize<Response>(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_edit", arguments)));
                Assert.IsTrue(response.Ok, response.Error); Assert.AreEqual(updated, fixture.Module.Code); Assert.AreEqual(1, fixture.Module.Writes);
                Assert.AreEqual(fault == "no subscriber" ? 0 : 1, logs, fault);
            }
        }

        [STATestMethod]
        public void MonacoRejectsForeignWindowThreadAndRestoreHandlesEmptyAndDisposedEditors()
        {
            using (var fixture = new CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture())
            using (var ready = new System.Threading.ManualResetEventSlim())
            {
                ModernEditorWindow foreign = null;
                var thread = new System.Threading.Thread(() => { foreign = new ModernEditorWindow(); var handle = foreign.Handle; ready.Set(); System.Windows.Forms.Application.Run(new System.Windows.Forms.ApplicationContext()); }) { IsBackground = true };
                thread.SetApartmentState(System.Threading.ApartmentState.STA); thread.Start(); Assert.IsTrue(ready.Wait(5000));
                try
                {
                    var tools = new LlmVbeTools(null, null, new LlmSettings()) { MonacoModule = (p, m) => fixture.Native.Adapter, MonacoWindow = create => foreign };
                    string response = CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}")); StringAssert.Contains(response, "owning VBE UI thread");
                }
                finally { foreign.BeginInvoke(new Action(() => { foreign.Dispose(); System.Windows.Forms.Application.ExitThread(); })); Assert.IsTrue(thread.Join(5000)); }
                using (var empty = new ModernEditorWindow())
                {
                    var tools = new LlmVbeTools(null, null, new LlmSettings()) { MonacoWindow = create => empty };
                    Assert.IsFalse(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.RestoreChangesAsync(null, null)).Ok);
                    empty.Dispose(); Assert.IsFalse(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.RestoreChangesAsync(null, null)).Ok);
                }
            }
        }

        [STATestMethod]
        public void MonacoPolicyAndDefaultSessionFactoriesRejectEveryMissingOrUnapprovedContext()
        {
            const string edit = "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":1,\"Text\":\"\"}";
            var unknown = Tools("unknown"); StringAssert.Contains(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(unknown.InvokeAsync("monaco_edit", edit)), "Unknown VBE edit policy");
            foreach (bool accepted in new[] { false, true })
            {
                var tools = Tools("AskEachTime"); tools.ShowApproval = (dialog, owner) => accepted ? System.Windows.Forms.DialogResult.Yes : System.Windows.Forms.DialogResult.No;
                StringAssert.Contains(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_edit", edit)), accepted ? "HOST WAS ACCESSED" : "User rejected");
            }
            var empty = new LlmVbeTools(null, null, new LlmSettings());
            var readWindow = typeof(LlmVbeTools).GetMethod("EditorWindow", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            Assert.IsNull(readWindow.Invoke(empty, new object[] { false }));
            foreach (string argument in new[] { "null", "[]", "{\"Project\":null,\"Module\":\"M\"}", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":0,\"Text\":\"\"}" })
                Assert.IsFalse(new JavaScriptSerializer().Deserialize<Response>(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(Tools().InvokeAsync(argument.Contains("ExpectedVersion") ? "monaco_edit" : "monaco_read", argument))).Ok);
            using (var fixture = new CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture())
            {
                var session = new VbeSession(fixture.Native.Vbe);
                var tools = new LlmVbeTools(session, null, new LlmSettings());
                string read = new JavaScriptSerializer().Serialize(new { Project = fixture.Native.Project.Name, Module = fixture.Native.Original.Name });
                StringAssert.Contains(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_read", read)), "monaco_open first");
                session.ModernEditor = create => fixture.Window;
                var response = new JavaScriptSerializer().Deserialize<Response>(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_open", read)));
                Assert.IsTrue(response.Ok, response.Error);
                Assert.AreEqual(true, ((Dictionary<string, object>)response.Data)["Loading"]);
                fixture.Ready(true); fixture.Rendering = (method, values) => method == "read" ? new JavaScriptSerializer().Serialize(new { text = fixture.Document.Text, version = fixture.Get<Dictionary<string, int>>("versions")[fixture.Document.Id], selection = new { startLineNumber = 1 } }) : "null";
                Assert.IsTrue(new JavaScriptSerializer().Deserialize<Response>(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_open", read))).Ok);
                fixture.Window.Dispose(); StringAssert.Contains(CodexVBE.Tests.Infrastructure.ModernEditorDebugFixture.Wait(tools.InvokeAsync("monaco_read", read)), "monaco_open first");
            }
        }

        private static LlmVbeTools Tools(string policy = "Automatic") => new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = policy })
        { MonacoModule = (p, m) => throw new InvalidOperationException("HOST WAS ACCESSED") };

        [TestMethod]
        public async Task EveryMonacoContractValidatesRequiredTypesWhitespaceAndDispatchesAsynchronously()
        {
            var json = new JavaScriptSerializer();
            int count = 0;
            foreach (var definition in LlmVbeTools.Definitions)
            {
                var root = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(definition));
                var function = (Dictionary<string, object>)root["function"];
                string name = (string)function["name"];
                if (!name.StartsWith("monaco_", StringComparison.Ordinal)) continue;
                count++;
                var parameters = (Dictionary<string, object>)function["parameters"];
                var properties = (Dictionary<string, object>)parameters["properties"];
                var valid = new Dictionary<string, object>();
                foreach (var field in properties)
                {
                    var schema = (Dictionary<string, object>)field.Value;
                    valid[field.Key] = (string)schema["type"] == "integer" ? (object)1 : field.Key == "Text" ? "" : "P";
                }
                var tools = Tools();
                StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(valid)), "HOST WAS ACCESSED", name + ":valid dispatch");
                StringAssert.Contains(tools.Invoke(name, json.Serialize(valid)), "InvokeAsync", name + ":sync refused");
                foreach (string required in (object[])parameters["required"])
                {
                    var missing = new Dictionary<string, object>(valid); missing.Remove(required);
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(missing)), required + " is required", name);
                }
                foreach (var field in properties)
                {
                    var invalid = new Dictionary<string, object>(valid);
                    bool integer = (string)((Dictionary<string, object>)field.Value)["type"] == "integer";
                    invalid[field.Key] = integer ? (object)"1" : 1;
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(invalid)), "Invalid Monaco argument", name + ":type " + field.Key);
                    if (integer || field.Key == "Text") continue;
                    invalid[field.Key] = "  ";
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(invalid)), "Invalid Monaco argument", name + ":whitespace " + field.Key);
                }
                valid["Extra"] = true;
                StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(valid)), "Unexpected argument", name);
            }
            Assert.AreEqual(5, count);
        }

        [TestMethod]
        public async Task RefusesMalformedArgumentsAndCrossProjectReadsBeforeHostAccess()
        {
            var tools = Tools();
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":\"1\",\"Text\":\"\"}"), "Invalid Monaco argument");
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\",\"Extra\":true}"), "Unexpected argument");
            tools.BoundProject = "Other";
            StringAssert.Contains(new JavaScriptSerializer().Deserialize<Response>(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}")).Error, UiText.Get("Read access to another project is not authorized for this conversation."));
            tools.SetReadAccess(new[] { "P" }, false);
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}"), "HOST WAS ACCESSED");
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":1,\"Text\":\"\"}"), "autre projet");
            StringAssert.Contains(tools.Invoke("monaco_open", "{}"), "InvokeAsync");
        }

        [TestMethod]
        public async Task EditingHonorsModeScopeAndReadOnlyPolicy()
        {
            const string args = "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":1,\"Text\":\"\"}";
            var tools = Tools("ReadOnly");
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", args), "Read-only mode");
            tools = Tools(); tools.Mode = ChatMode.Discussion;
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", args), "monaco_edit");
            tools = Tools(); tools.ValidateScope = () => throw new InvalidOperationException("expired scope");
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}"), "expired scope");
        }
    }
}
