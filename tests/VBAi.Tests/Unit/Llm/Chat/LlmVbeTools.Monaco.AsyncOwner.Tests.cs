using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks native lookup admission and rollback across renderer awaits on a verified owner.</summary>
    [ToolBoundaryTestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class LlmMonacoAsyncOwnerTests
    {
        [STATestMethod]
        [DataRow("monaco_open"), DataRow("monaco_read"), DataRow("monaco_navigate"), DataRow("monaco_edit"), DataRow("monaco_sync")]
        public void EveryMonacoToolReachesTheModuleFactoryOnlyAfterOwnerWindowAdmission(string tool)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                int owner = Thread.CurrentThread.ManagedThreadId, windows = 0, modules = 0;
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                {
                    MonacoWindow = create =>
                    {
                        Assert.AreEqual(tool == "monaco_open", create);
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        windows++; return fixture.Window;
                    },
                    MonacoModule = (project, module) =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual("P", project); Assert.AreEqual("M", module);
                        modules++; throw new InvalidOperationException("owned module lookup reached");
                    }
                };
                var values = new System.Collections.Generic.Dictionary<string, object> { ["Project"] = "P", ["Module"] = "M" };
                if (tool == "monaco_edit") { values["ExpectedVersion"] = 1; values["Text"] = ""; }
                if (tool == "monaco_sync") { values["ExpectedVersion"] = 1; values["ExpectedSha256"] = EditorDocument.Hash(""); }
                if (tool == "monaco_navigate")
                    foreach (string field in new[] { "ExpectedVersion", "StartLine", "StartColumn", "EndLine", "EndColumn" }) values[field] = 1;
                var original = SynchronizationContext.Current;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(null);
                    var json = new JavaScriptSerializer();
                    var pending = tools.InvokeAsync(tool, json.Serialize(values));
                    Assert.IsNull(SynchronizationContext.Current, "Check caller restoration before the fixture pumps Windows messages.");
                    var response = json.Deserialize<Response>(ModernEditorDebugFixture.Wait(pending));
                    Assert.IsFalse(response.Ok);
                    Assert.AreEqual("owned module lookup reached", response.Error);
                    Assert.AreEqual(1, windows); Assert.AreEqual(1, modules);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }

        [STATestMethod]
        [DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void RollbackAfterCapturePreservesSourceGuardAndOwnerCleanup(bool baseContext, bool disposeDuringWrite)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                fixture.Ready(true);
                int owner = Thread.CurrentThread.ManagedThreadId, writes = 0;
                string code = "after";
                fixture.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    return "null";
                };
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                {
                    MonacoWindow = create => fixture.Window,
                    Execute = request =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        if (request.Command == "read_module") return Response.Success(new VbeToolCodeResult { Code = code, Sha256 = EditorDocument.Hash(code) });
                        Assert.AreEqual("replace_lines", request.Command);
                        Assert.AreEqual(EditorDocument.Hash(code), request.ExpectedSha256);
                        code = request.Text; writes++;
                        if (disposeDuringWrite) fixture.Window.Dispose();
                        return Response.Success(new { Replaced = true });
                    }
                };
                var change = new CodeChange("P", "M", "before", EditorDocument.Hash("before"), "after", EditorDocument.Hash("after"), 1);
                var original = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = tools.RestoreChangesAsync(new[] { change }, null);
                    Assert.AreSame(ambient, SynchronizationContext.Current, "Check caller restoration before the fixture pumps Windows messages.");
                    var response = ModernEditorDebugFixture.Wait(pending);
                    Assert.IsTrue(response.Ok, response.Error);
                    Assert.AreEqual("before", code); Assert.AreEqual(1, writes);
                    Assert.IsTrue(change.Restored);
                    Assert.AreEqual(disposeDuringWrite, fixture.Window.IsDisposed);
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }
    }
}
