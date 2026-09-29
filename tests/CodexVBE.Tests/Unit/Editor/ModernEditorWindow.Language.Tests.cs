using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorLanguageTests
    {
        private static object Message(ModernEditorToolFixture f, string id = null, int version = 1, string module = null, int line = 1, int column = 1, bool nullId = false)
        {
            var type = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic); var message = Activator.CreateInstance(type, true);
            foreach (var item in new[] { Tuple.Create("id", nullId ? null : (object)(id ?? f.Document.Id)), Tuple.Create("version", (object)version), Tuple.Create("request", (object)42), Tuple.Create("module", (object)module), Tuple.Create("line", (object)line), Tuple.Create("column", (object)column) }) type.GetProperty(item.Item1).SetValue(message, item.Item2);
            return message;
        }
        private static void Call(ModernEditorToolFixture f, string method, object message) => ModernEditorDebugFixture.Wait((Task)f.Private(method, message));
        [STATestMethod]
        public void LanguageRequestsReplyOnlyToLiveOwnedWindowsAndHandleMissingStaleOrFailedSnapshots()
        {
            foreach (string state in new[] { "null-id", "missing", "stale", "error", "not-ready", "disposed", "success", "existing-worker", "failed-worker" })
            using (var f = new ModernEditorToolFixture())
            {
                EditorSyncWorker worker = null;
                if (state == "not-ready") f.Base.Ready(false); if (state == "disposed") f.Window.Dispose();
                if (state == "error") f.Override = (method, values) => method == "snapshots" ? throw new IOException("owned snapshot unavailable") : (string)null;
                if (state == "existing-worker" || state == "failed-worker")
                {
                    worker = new EditorSyncWorker(); f.Base.Set("languageWorker", worker);
                    if (state == "failed-worker") { worker.Dispose(); Assert.IsTrue(((Thread)typeof(EditorSyncWorker).GetField("thread", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(worker)).Join(5000)); }
                }
                f.Base.Scripts.Clear();
                Call(f, "LanguageRequest", Message(f, id: state == "missing" ? "missing" : null, version: state == "stale" ? 0 : 1, nullId: state == "null-id"));
                var replies = f.Base.Scripts.Where(script => script.Item1 == "languageReply").ToArray();
                Assert.AreEqual(state == "disposed" || state == "not-ready" ? 0 : 1, replies.Length, state);
                if (replies.Length != 0)
                {
                    Assert.AreEqual(42, replies[0].Item2[0]);
                    if (state == "success" || state == "existing-worker")
                    {
                        var response = f.Result(replies[0].Item2[1]); Assert.AreEqual(f.Document.Id, response["id"]); Assert.AreEqual(f.Module.Name, response["module"]); Assert.IsNotNull(response["symbols"]);
                    }
                    else Assert.IsNull(replies[0].Item2[1], state);
                }
            }
        }
        [STATestMethod]
        public void NativeLanguageOverlaysOnlyItsOwnProjectAndReadsSafeLocalTypeMetadata()
        {
            using (var f = new ModernEditorToolFixture())
            {
                var native = f.Base.Native;
                native.Original.CodeModule.Raw += "\nPublic dictionary As Scripting.Dictionary";
                f.Base.Document.AcceptRemote(native.Original.CodeModule.Raw); f.Base.Document.Edit(f.Base.Document.Text + "\nPublic draftOnly As Long");
                native.Project.References.Add(new EditorVbeContract.Reference { IsBroken = true, FullPath = "must not read" });
                native.Project.References.Add(new EditorVbeContract.Reference { FullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll") });
                var other = new EditorVbeContract.Component { Name = "Removed", Collection = native.Project.VBComponents, Type = 1 }; other.CodeModule.Raw = "Public local As Long"; native.Project.VBComponents.Items.Add(other);
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(new EditorVbeModule(native.Vbe, native.Project, other))); native.Project.VBComponents.Items.Remove(other);
                var foreign = new EditorVbeContract(); foreign.Original.Name = "Foreign"; ModernEditorDebugFixture.Wait(f.Window.OpenModule(foreign.Adapter));
                f.Base.Scripts.Clear(); Call(f, "LanguageRequest", Message(f, id: f.Base.Document.Id));
                var reply = f.Base.Scripts.Single(script => script.Item1 == "languageReply"); Assert.IsNotNull(reply.Item2[1]);
                var response = f.Result(reply.Item2[1]); Assert.AreEqual("Module1", response["module"]);
                string payload = f.Json.Serialize(response); StringAssert.Contains(payload, "draftOnly"); StringAssert.Contains(payload, "Dictionary"); Assert.IsFalse(payload.Contains("Foreign")); Assert.IsFalse(payload.Contains("Removed"));
                Assert.IsTrue(payload.Contains("\"External\":true"));
            }
        }
        [STATestMethod]
        public void DefinitionRoutesExactOwnedModulesAndClampsSourcePositions()
        {
            foreach (string state in new[] { "null-id", "missing", "missing-target", "managed", "native" })
            using (var f = new ModernEditorToolFixture())
            {
                f.Base.Scripts.Clear();
                Call(f, "OpenDefinition", Message(f, id: state == "missing" ? "missing" : state == "native" ? f.Base.Document.Id : null, module: state == "missing-target" ? "absent" : state == "native" ? "Module1" : f.Module.Name, line: -1, column: 0, nullId: state == "null-id"));
                var reveal = f.Base.Scripts.Where(script => script.Item1 == "reveal").ToArray(); Assert.AreEqual(state == "native" || state == "managed" ? 1 : 0, reveal.Length, state);
                if (reveal.Length != 0) { Assert.AreEqual(1, reveal[0].Item2[0]); Assert.AreEqual(1, reveal[0].Item2[1]); }
            }
        }
    }
}