using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Unit;

namespace VBAi.Tests.Infrastructure
{
    internal sealed class ModernEditorDebugFixture : IDisposable
    {
        internal readonly EditorFixture Storage = new EditorFixture();
        internal readonly EditorVbeContract Native = new EditorVbeContract();
        internal readonly ModernEditorWindow Window = new ModernEditorWindow();
        internal readonly List<Tuple<string, object[]>> Scripts = new List<Tuple<string, object[]>>();
        internal readonly VbeDebugTests.FakeBar Bar = new VbeDebugTests.FakeBar { Name = "Debug" };
        internal readonly EditorDocument Document;
        internal string Diagnostic;
        internal Func<string, object[], string> Rendering;
        internal int Preflight, Observations;
        internal ModernEditorDebugFixture()
        {
            Window.Drafts = new EditorDraftStore(Storage.Root);
            Native.Vbe.MainWindow.HWnd = Window.Handle.ToInt32();
            Native.Vbe.CommandBars.Add(Bar);
            Window.ScriptExecution = (method, values) => { Scripts.Add(Tuple.Create(method, values)); return Task.FromResult(Rendering?.Invoke(method, values) ?? "null"); };
            Window.EnsureCompileDialogAbsent = pid => { Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(Process.GetCurrentProcess().Id, pid); Preflight++; };
            Window.ObserveCompileDialog = (completed, pid) => { Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(Process.GetCurrentProcess().Id, pid); if (!completed.Wait(5000)) throw new TimeoutException("Fixture compilation did not return"); Interlocked.Increment(ref Observations); return Diagnostic; };
            Document = Wait(Window.OpenModule(Native.Adapter));
        }
        internal VbeDebugTests.FakeControl Command(int id, string caption, Action action = null, bool enabled = true)
        {
            var control = new VbeDebugTests.FakeControl { Id = id, Caption = caption, OnExecute = action, Enabled = enabled }; Bar.Controls.Add(control); return control;
        }
        internal void Ready(bool value) => typeof(ModernEditorWindow).GetProperty("Ready", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Window, value);
        internal void Set(string field, object value) => typeof(ModernEditorWindow).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Window, value);
        internal T Get<T>(string field) => (T)typeof(ModernEditorWindow).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Window);
        internal void Observe() => Wait((Task)typeof(ModernEditorWindow).GetMethod("ObserveDebugMode", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, null));
        internal void Send(string name, int line = 3, string id = null, int? version = null, bool nullId = false)
        {
            Type type = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
            object message = Activator.CreateInstance(type, true);
            foreach (var pair in new[] { Tuple.Create("name", (object)name), Tuple.Create("id", nullId ? null : (object)(id ?? Document.Id)), Tuple.Create("version", (object)(version ?? Get<Dictionary<string, int>>("versions")[Document.Id])), Tuple.Create("line", (object)line) }) type.GetProperty(pair.Item1).SetValue(message, pair.Item2);
            Wait((Task)typeof(ModernEditorWindow).GetMethod("EditorCommand", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, new[] { message }));
        }
        internal static void Wait(Task task)
        {
            var timeout = Stopwatch.StartNew();
            while (!task.IsCompleted) { if (timeout.ElapsedMilliseconds > 15000) throw new TimeoutException("Owned editor contract task did not complete"); Application.DoEvents(); Thread.Sleep(1); }
            task.GetAwaiter().GetResult(); Application.DoEvents();
        }
        internal static T Wait<T>(Task<T> task) { Wait((Task)task); return task.GetAwaiter().GetResult(); }
        public void Dispose() { Ready(false); Window.Dispose(); Storage.Dispose(); }
    }
}
