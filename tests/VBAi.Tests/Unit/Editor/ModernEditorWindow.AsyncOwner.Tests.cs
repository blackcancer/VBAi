using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit.Editor
{
    /// <summary>Exercises native editor entry points after real awaits without a host-provided context.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ModernEditorAsyncOwnerTests
    {
        [STATestMethod]
        [DataRow("open", false), DataRow("open", true)]
        [DataRow("background", false), DataRow("background", true)]
        [DataRow("documents", false), DataRow("documents", true)]
        [DataRow("stream", false), DataRow("stream", true)]
        public void NativeReadsWritesAndUnusedAdapterCleanupRemainOnOwnerAfterRendererAwait(string entry, bool baseContext)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                fixture.Ready(true);
                int owner = Thread.CurrentThread.ManagedThreadId, reads = 0;
                fixture.Native.Original.CodeModule.BeforeRead = () => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); reads++; };
                fixture.Native.Original.CodeModule.BeforeOperation = _ => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                fixture.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, method + " before await");
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId, method + " after await");
                    return method == "apply" ? ((int)values[1] + 1).ToString() : "null";
                };
                var previous = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    if (entry == "open")
                    {
                        var released = new System.Collections.Generic.List<object>();
                        var duplicate = EditorVbeModule.AcquireActive(fixture.Native.Vbe, false, value =>
                        { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); released.Add(value); });
                        var pending = fixture.Window.OpenModule(duplicate);
                        Assert.AreSame(ambient, SynchronizationContext.Current);
                        var document = ModernEditorDebugFixture.Wait(pending);
                        Assert.AreSame(fixture.Document, document);
                        Assert.AreEqual(5, released.Count);
                        Assert.ThrowsException<ObjectDisposedException>(() => duplicate.Read());
                        Assert.IsTrue(fixture.Window.OwnsModule(fixture.Native.Adapter));
                    }
                    else
                    {
                        if (entry != "background") fixture.Document.Edit(fixture.Document.Text + "\n' owner continuation");
                        if (entry == "documents")
                        {
                            var pending = fixture.Window.ProcessDocuments(true);
                            Assert.AreSame(ambient, SynchronizationContext.Current);
                            ModernEditorDebugFixture.Wait(pending);
                        }
                        else
                        {
                            if (entry == "stream")
                            {
                                fixture.Set("streamSequence", 1);
                                fixture.Set("lastEdit", DateTime.UtcNow.AddSeconds(-2));
                            }
                            var method = typeof(ModernEditorWindow).GetMethod(entry == "stream" ? "FlushStream" : "ProcessBackgroundDocuments", BindingFlags.Instance | BindingFlags.NonPublic);
                            var pending = (Task)method.Invoke(fixture.Window, null);
                            Assert.AreSame(ambient, SynchronizationContext.Current);
                            ModernEditorDebugFixture.Wait(pending);
                        }
                        Assert.IsTrue(reads > 0, "The native adapter must actually be exercised.");
                        Assert.IsFalse(fixture.Document.Conflict);
                        Assert.IsFalse(fixture.Document.Dirty);
                        Assert.IsFalse(fixture.Get<bool>("busy"));
                        Assert.IsFalse(fixture.Get<System.Collections.Generic.Dictionary<string, string>>("documentSynchronizationErrors").ContainsKey(fixture.Document.Id));
                    }
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }

        [STATestMethod]
        [DataRow("toggle_breakpoint", false), DataRow("toggle_breakpoint", true)]
        [DataRow("show_next_statement", false), DataRow("show_next_statement", true)]
        public void NativeDebugDispatchAndSelectionObservationKeepTheCommandOwner(string command, bool baseContext)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                int owner = Thread.CurrentThread.ManagedThreadId;
                fixture.Ready(true);
                fixture.Native.Original.CodeModule.BeforeRead = () => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                fixture.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    return "null";
                };
                if (command == "show_next_statement") fixture.Native.Project.Mode = 1;
                var control = fixture.Command(command == "toggle_breakpoint" ? 51 : 1813, command == "toggle_breakpoint" ? "Toggle Breakpoint" : "Show Next Statement", () => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId));
                var previous = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var messageType = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
                    var message = Activator.CreateInstance(messageType, true);
                    messageType.GetProperty("id").SetValue(message, fixture.Document.Id);
                    messageType.GetProperty("version").SetValue(message, 1);
                    messageType.GetProperty("name").SetValue(message, command);
                    messageType.GetProperty("line").SetValue(message, 3);
                    var method = typeof(ModernEditorWindow).GetMethod("EditorCommand", BindingFlags.Instance | BindingFlags.NonPublic);
                    var pending = (Task)method.Invoke(fixture.Window, new[] { message });
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    ModernEditorDebugFixture.Wait(pending);
                    Assert.AreEqual(1, control.ExecuteCount);
                    Assert.AreEqual(1, fixture.Get<SemaphoreSlim>("debugCommands").CurrentCount);
                    Assert.IsFalse(fixture.Get<bool>("busy"));
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }

        [STATestMethod]
        [DataRow("LanguageRequest", false), DataRow("LanguageRequest", true)]
        [DataRow("OpenDefinition", false), DataRow("OpenDefinition", true)]
        public void NativeLanguageReadsAndDefinitionAdapterCleanupRemainOnOwner(string entry, bool baseContext)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                fixture.Ready(true);
                int owner = Thread.CurrentThread.ManagedThreadId, reads = 0;
                fixture.Native.Original.CodeModule.BeforeRead = () => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); reads++; };
                fixture.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    fixture.Scripts.Add(Tuple.Create(method, values));
                    return "null";
                };
                var messageType = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
                var message = Activator.CreateInstance(messageType, true);
                messageType.GetProperty("id").SetValue(message, fixture.Document.Id);
                messageType.GetProperty("version").SetValue(message, 1);
                messageType.GetProperty("request").SetValue(message, 42);
                messageType.GetProperty("module").SetValue(message, fixture.Native.Original.Name);
                var previous = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var method = typeof(ModernEditorWindow).GetMethod(entry, BindingFlags.Instance | BindingFlags.NonPublic);
                    var pending = (Task)method.Invoke(fixture.Window, new[] { message });
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    ModernEditorDebugFixture.Wait(pending);
                    if (entry == "LanguageRequest")
                    {
                        Assert.IsTrue(reads > 0);
                        var reply = fixture.Scripts.Single(call => call.Item1 == "languageReply");
                        Assert.IsNotNull(reply.Item2[1], "An affinity failure must not be hidden as an empty language reply.");
                        Assert.AreEqual(1, fixture.Get<SemaphoreSlim>("languageGate").CurrentCount);
                    }
                    else
                    {
                        Assert.AreEqual(1, fixture.Window.Documents.Count());
                        Assert.IsTrue(fixture.Scripts.Any(call => call.Item1 == "reveal"));
                        Assert.IsTrue(fixture.Native.Adapter.Read().Contains("Hello"));
                    }
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }
    }
}
