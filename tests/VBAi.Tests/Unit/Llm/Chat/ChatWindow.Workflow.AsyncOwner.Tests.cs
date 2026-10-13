using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public partial class ChatWindowStateTests
    {
        /// <summary>Exercises the complete attachment-navigation task after a real renderer wait.</summary>
        [STATestMethod]
        [DataRow(false, "current"), DataRow(true, "current")]
        [DataRow(false, "stale"), DataRow(true, "stale")]
        [DataRow(false, "capture-error"), DataRow(true, "capture-error")]
        public void AttachmentNavigationPreservesOwnerAndRevisionThroughRendererAwaits(bool baseContext, string state)
        {
            using (var runtime = new RuntimeScope())
            using (var editor = new Editor.ModernEditorToolFixture())
            using (var window = LoadedWindow(runtime.Session))
            {
                int owner = Thread.CurrentThread.ManagedThreadId, shown = 0, captures = 0;
                runtime.Session.ModernEditor = create => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); if (create) shown++; return editor.Window; };
                editor.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, method + " before wait");
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, method + " after wait");
                    if (method == "snapshots")
                    {
                        captures++;
                        if (state == "capture-error") throw new InvalidOperationException("owned attachment capture failed");
                    }
                    editor.Base.Scripts.Add(Tuple.Create(method, values));
                    return "null";
                };
                var attachment = new ChatAttachment
                {
                    Label = "Owned Monaco selection",
                    Project = @"C:\Temp\P.xlsm",
                    Module = "M",
                    EditorDocumentId = editor.Document.Id,
                    Sha256 = state == "stale" ? "stale" : EditorDocument.Hash(editor.Document.Text),
                    StartLine = 3
                };
                editor.Base.Scripts.Clear();
                var original = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = (Task)Call(window, "NavigateAttachmentAsync", attachment);
                    Assert.AreSame(ambient, SynchronizationContext.Current, "Check restoration before the fixture pumps Windows messages.");
                    ModernEditorDebugFixture.Wait(pending);
                    Assert.AreEqual(1, captures);
                    Assert.AreEqual(state == "current" ? 1 : 0, shown);
                    var reveals = editor.Base.Scripts.Where(call => call.Item1 == "reveal").ToArray();
                    Assert.AreEqual(state == "current" ? 1 : 0, reveals.Length);
                    if (state == "current") { Assert.AreEqual(3, reveals[0].Item2[0]); Assert.AreEqual(1, reveals[0].Item2[1]); }
                    else StringAssert.Contains(Get<Label>(window, "status").Text,
                        state == "stale" ? UiText.Get(" is stale. Remove it and attach the current selection.") : "owned attachment capture failed");
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }
    }
}
