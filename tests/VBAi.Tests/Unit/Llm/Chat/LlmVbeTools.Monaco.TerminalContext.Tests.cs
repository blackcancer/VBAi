using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks reply publication independently of a caller's borrowed synchronization context.</summary>
    [ToolBoundaryTestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class LlmMonacoTerminalContextTests
    {
        /// <summary>Models a host context that retains posts but does not process them.</summary>
        private sealed class HoldingContext : SynchronizationContext
        {
            internal int Posts;
            public override void Post(SendOrPostCallback callback, object state) { Interlocked.Increment(ref Posts); }
        }

        [STATestMethod]
        [DataRow("monaco_open", "holding"), DataRow("monaco_open", "plain"), DataRow("monaco_open", "none")]
        [DataRow("monaco_read", "holding"), DataRow("monaco_read", "plain"), DataRow("monaco_read", "none")]
        [DataRow("monaco_navigate", "holding"), DataRow("monaco_navigate", "plain"), DataRow("monaco_navigate", "none")]
        public void ReadyRendererReplyCompletesWithoutPostingToTheBorrowedContext(string command, string contextKind)
            => Run(command, contextKind, false);

        [STATestMethod]
        [DataRow("monaco_open", "holding"), DataRow("monaco_open", "plain"), DataRow("monaco_open", "none")]
        [DataRow("monaco_read", "holding"), DataRow("monaco_read", "plain"), DataRow("monaco_read", "none")]
        [DataRow("monaco_navigate", "holding"), DataRow("monaco_navigate", "plain"), DataRow("monaco_navigate", "none")]
        [DataRow("monaco_edit", "holding"), DataRow("monaco_edit", "plain"), DataRow("monaco_edit", "none")]
        [DataRow("monaco_sync", "holding"), DataRow("monaco_sync", "plain"), DataRow("monaco_sync", "none")]
        public void RendererFailureReplyCompletesWithoutChangingOwnerOrRetrying(string command, string contextKind)
            => Run(command, contextKind, true);

        /// <summary>Runs the same exact document through owner-pinned renderer awaits and observes its terminal reply.</summary>
        private static void Run(string command, string contextKind, bool failRenderer)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                fixture.Ready(true);
                var json = new JavaScriptSerializer();
                int owner = Thread.CurrentThread.ManagedThreadId, renders = 0, lookups = 0, writes = 0;
                fixture.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    renders++;
                    if (failRenderer) throw new InvalidOperationException("synthetic renderer refusal");
                    if (method == "snapshots") return "[]";
                    if (method == "selectRange") return "true";
                    if (method == "read") return json.Serialize(new
                    {
                        text = fixture.Document.Text,
                        version = fixture.Get<Dictionary<string, int>>("versions")[fixture.Document.Id],
                        selection = new { startLineNumber = 1, startColumn = 1, endLineNumber = 1, endColumn = 1 }
                    });
                    return "null";
                };
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
                {
                    MonacoWindow = create => fixture.Window,
                    MonacoModule = (project, module) =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        lookups++; return fixture.Native.Adapter;
                    },
                    Execute = request =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        if (request.Command != "read_module") writes++;
                        return Response.Success(new VbeToolCodeResult { Code = fixture.Document.Native, Sha256 = EditorDocument.Hash(fixture.Document.Native) });
                    }
                };
                var arguments = new Dictionary<string, object> { ["Project"] = "P", ["Module"] = "M" };
                if (command == "monaco_edit") { arguments["ExpectedVersion"] = 1; arguments["Text"] = fixture.Document.Text; }
                if (command == "monaco_sync") { arguments["ExpectedVersion"] = 1; arguments["ExpectedSha256"] = EditorDocument.Hash(fixture.Document.Native); }
                if (command == "monaco_navigate")
                    foreach (string name in new[] { "ExpectedVersion", "StartLine", "StartColumn", "EndLine", "EndColumn" }) arguments[name] = 1;
                var original = SynchronizationContext.Current;
                var holding = new HoldingContext();
                SynchronizationContext ambient = contextKind == "holding" ? holding : contextKind == "plain" ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = tools.InvokeAsync(command, json.Serialize(arguments));
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    var reply = json.Deserialize<Response>(ModernEditorDebugFixture.Wait(pending));
                    Assert.AreEqual(!failRenderer, reply.Ok, reply.Error);
                    if (failRenderer) { Assert.AreEqual("synthetic renderer refusal", reply.Error); Assert.AreEqual(1, renders); }
                    else Assert.IsTrue(renders > 0);
                    Assert.AreEqual(1, lookups); Assert.AreEqual(0, writes);
                    Assert.AreEqual(0, holding.Posts, "Reply publication must not depend on a retired or non-pumping host context.");
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }
    }
}
