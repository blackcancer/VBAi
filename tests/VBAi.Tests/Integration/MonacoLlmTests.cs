using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("MonacoRuntime")]
    public sealed class MonacoLlmTests : EditorUiTestFixture
    {
        /// <summary>Finishes startup and pauses automatic work for manual synchronization scenarios.</summary>
        private static void PauseAutomaticWork(ModernEditorWindow window)
        {
            MonacoRuntimeTests.Wait(() => window.Ready && UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Enabled);
            foreach (string name in new[] { "timer", "streamTimer", "debugTimer" })
                UiInvoke.Field<System.Windows.Forms.Timer>(window, name).Stop();
            MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
        }

        [STATestMethod]
        public void CompilationCapturesUnreportedDraftBeforeNativeDispatch()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(host.Root) })
            {
                window.Show(); PauseAutomaticWork(window);
                var doc = MonacoRuntimeTests.Wait(window.OpenModule(host));
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                { MonacoWindow = create => window, Execute = request => throw new InvalidOperationException("Native dispatch was reached") };
                tools.Native.EnsureNoCompileDialog = () => throw new InvalidOperationException("Native compilation was reached");
                // apply suppresses the renderer change message: only a fresh capture can see it.
                Assert.AreNotEqual("0", MonacoRuntimeTests.Wait(window.Script("apply", doc.Id, 1, host.Code + "\n' unreported draft")));
                Assert.IsFalse(doc.Dirty);
                string result = MonacoRuntimeTests.Wait(tools.InvokeAsync("compile_project", "{\"Project\":\"P\",\"ExpectedMode\":2}"));
                StringAssert.Contains(result, "unsynchronized");
                Assert.IsTrue(doc.Dirty); Assert.AreEqual(0, host.Writes);
                StringAssert.Contains(tools.Invoke("compile_project", "{\"Project\":\"P\",\"ExpectedMode\":2}"), "unsynchronized");
                window.Close(); MonacoRuntimeTests.Wait(() => window.IsDisposed);
            }
        }

        [STATestMethod]
        public void ReconciliationPreservesTypingPumpedAfterNativeWrite()
        {
            var json = new JavaScriptSerializer();
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(host.Root) })
            {
                window.Show(); PauseAutomaticWork(window);
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                {
                    MonacoWindow = create => window,
                    MonacoModule = (p, m) => host,
                    Execute = request =>
                    {
                        dynamic payload = new System.Dynamic.ExpandoObject();
                        payload.Code = host.Code; payload.Sha256 = EditorDocument.Hash(host.Code);
                        return Response.Success((object)payload);
                    }
                };
                Func<string, object, Dictionary<string, object>> call = (name, args) =>
                    json.Deserialize<Dictionary<string, object>>(MonacoRuntimeTests.Wait(tools.InvokeAsync(name, json.Serialize(args))));
                var opened = call("monaco_open", new { Project = "P", Module = "M" });
                Assert.AreEqual(true, opened["Ok"], json.Serialize(opened));
                var state = (Dictionary<string, object>)opened["Data"];
                var doc = window.Documents.Single(); CodeChange observed = null;
                tools.CodeEdited += change =>
                {
                    observed = change;
                    // Native COM and UI subscribers may pump messages before reconciliation.
                    MonacoRuntimeTests.Wait(window.Script("insert", "' newer user typing\n"));
                    MonacoRuntimeTests.Wait(() => doc.Text.Contains("newer user typing"));
                };
                string updated = host.Code.Replace("Debug.Print 1", "Debug.Print 2");
                var result = call("monaco_edit", new { Project = "P", Module = "M", ExpectedVersion = (int)state["Version"], Text = updated });
                Assert.AreEqual(true, result["Ok"], json.Serialize(result));
                state = (Dictionary<string, object>)result["Data"];
                Assert.AreEqual(true, state["Synchronized"]); Assert.AreEqual(true, state["Dirty"]);
                Assert.AreEqual(updated, host.Code);
                StringAssert.Contains(doc.Text, "newer user typing");
                Assert.IsNotNull(observed); Assert.AreEqual(updated, observed.After);
                window.Close(); MonacoRuntimeTests.Wait(() => window.IsDisposed);
            }
        }

        [STATestMethod]
        public void RealMonacoToolsSynchronizeEmitDiffRejectStaleAndPreserveConflicts()
        {
            var json = new JavaScriptSerializer();
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(host.Root) })
            {
                window.Show(); PauseAutomaticWork(window);
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                {
                    BoundProject = "P",
                    MonacoWindow = create => window,
                    MonacoModule = (p, m) => host,
                    Execute = request =>
                    {
                        dynamic payload = new System.Dynamic.ExpandoObject();
                        payload.Code = host.Code;
                        payload.Sha256 = EditorDocument.Hash(host.Code);
                        return Response.Success((object)payload);
                    }
                };
                int diffs = 0; CodeChange observed = null; tools.CodeEdited += change => { diffs++; observed = change; };
                Func<string, object, Dictionary<string, object>> call = (name, args) =>
                    json.Deserialize<Dictionary<string, object>>(MonacoRuntimeTests.Wait(tools.InvokeAsync(name, json.Serialize(args))));
                var response = call("monaco_open", new { Project = "P", Module = "M" }); Assert.AreEqual(true, response["Ok"]);
                var state = (Dictionary<string, object>)response["Data"];
                int originalVersion = (int)state["Version"];
                string updated = host.Code.Replace("Debug.Print 1", "Debug.Print 2");
                response = call("monaco_edit", new { Project = "P", Module = "M", ExpectedVersion = originalVersion, Text = updated });
                Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                state = (Dictionary<string, object>)response["Data"]; Assert.AreEqual(true, state["Synchronized"]);
                Assert.AreEqual(updated, host.Code); Assert.AreEqual(1, diffs);
                response = call("monaco_edit", new { Project = "P", Module = "M", ExpectedVersion = originalVersion, Text = "stale" });
                Assert.AreEqual(false, response["Ok"]); Assert.AreEqual(updated, host.Code);
                response = call("monaco_read", new { Project = "P", Module = "M" }); state = (Dictionary<string, object>)response["Data"];
                response = call("monaco_navigate", new { Project = "P", Module = "M", ExpectedVersion = (int)state["Version"], StartLine = 3, StartColumn = 5, EndLine = 3, EndColumn = 16 });
                Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                state = (Dictionary<string, object>)response["Data"];
                StringAssert.Contains(json.Serialize(state["Selection"]), "Debug.Print");
                var doc = window.Documents.Single();
                // Inject renderer input immediately after its CAS but before the host snapshot.
                MonacoRuntimeTests.Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.vbai.originalApply = window.vbai.apply; window.vbai.apply = function(id,v,t) { const applied = window.vbai.originalApply(id,v,t); window.vbai.apply = window.vbai.originalApply; if(applied) window.vbai.insert('user draft'); return applied; };"));
                response = call("monaco_edit", new { Project = "P", Module = "M", ExpectedVersion = (int)state["Version"], Text = updated + "\n' model draft" });
                Assert.AreEqual(true, response["Ok"], json.Serialize(response));
                var raced = (Dictionary<string, object>)response["Data"];
                Assert.AreEqual(true, raced["AppliedToDraft"]); Assert.AreEqual(false, raced["Synchronized"]);
                Assert.AreEqual(updated, host.Code); Assert.AreEqual(1, diffs, "Concurrent user typing must not be attributed to the model diff.");
                MonacoRuntimeTests.Wait(() => doc.Dirty);
                host.Code = host.Code.Replace("Debug.Print 2", "Debug.Print 3");
                response = call("monaco_read", new { Project = "P", Module = "M" });
                state = (Dictionary<string, object>)response["Data"]; Assert.AreEqual(true, state["Conflict"]);
                response = call("monaco_sync", new { Project = "P", Module = "M", ExpectedVersion = (int)state["Version"], ExpectedSha256 = (string)state["NativeSha256"] });
                Assert.AreEqual(false, response["Ok"]); StringAssert.Contains(doc.Text, "user draft"); StringAssert.Contains(host.Code, "Debug.Print 3");
                StringAssert.Contains(tools.Invoke("replace_lines", "{}"), "unsynchronized");
                StringAssert.Contains(MonacoRuntimeTests.Wait(tools.InvokeAsync("git_pull", "{}")), "unsynchronized");
                var rollback = MonacoRuntimeTests.Wait(tools.RestoreChangesAsync(new[] { observed }, null));
                Assert.IsFalse(rollback.Ok); StringAssert.Contains(rollback.Error, "unsynchronized");
                StringAssert.Contains(doc.Text, "user draft"); StringAssert.Contains(host.Code, "Debug.Print 3");
                window.Close(); MonacoRuntimeTests.Wait(() => window.IsDisposed);
            }
        }
    }
}
