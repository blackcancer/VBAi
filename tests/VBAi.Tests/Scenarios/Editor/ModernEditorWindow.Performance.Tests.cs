using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using VBAi;
using VBAi.Tests.Infrastructure;
using VBAi.Tests.Unit.Editor;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Scenarios.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorPerformanceTests
    {
        [STATestMethod]
        public void BreakpointUsesDirectLookupAndOneSnapshotBarrier()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                var command = f.Command(51, "Toggle Breakpoint");
                f.Send("toggle_breakpoint");
                Assert.AreEqual(1, command.ExecuteCount);
                Assert.AreEqual(2, f.Native.Vbe.CommandBars.FindControlCalls);
                Assert.AreEqual(1, f.Scripts.Count(s => s.Item1 == "snapshots"));
                Assert.IsNull(f.Get<EditorSyncWorker>("synchronizationWorker"), "Clean documents must not queue diff/persistence work.");
            }
        }
        [STATestMethod]
        public void StepObservationIsScheduledWithoutWaitingForTheSynchronizationTimer()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                f.Ready(true); f.Native.Project.Mode = 1;
                var step = f.Command(188, "Step Into", () => f.Native.Original.CodeModule.CodePane.SetSelection(4, 1, 4, 1));
                f.Command(1813, "Show Next Statement");
                f.Send("step_into");
                Integration.MonacoRuntimeTests.Wait(() => f.Get<int>("lastExecutionLine") == 4, 5);
                Assert.AreEqual(1, step.ExecuteCount);
                Assert.IsFalse(UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "timer").Enabled);
                Assert.IsTrue(f.Scripts.Any(s => s.Item1 == "executionBatch" && (int)s.Item2[1] == 4));
            }
        }
        [STATestMethod]
        public void QueuedBreakpointCannotConsumeTheOnlyPostStepObservation()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                f.Ready(true); f.Native.Project.Mode = 1;
                var snapshots = new TaskCompletionSource<string>();
                int snapshotCount = 0;
                f.Window.ScriptExecution = (method, values) => {
                    f.Scripts.Add(Tuple.Create(method, values));
                    return method == "snapshots" && ++snapshotCount == 2 ? snapshots.Task : Task.FromResult("null");
                };
                Task queued = null;
                var toggle = f.Command(51, "Toggle Breakpoint");
                var observation = f.Command(1813, "Show Next Statement");
                var step = f.Command(188, "Step Into", () => {
                    f.Native.Original.CodeModule.CodePane.SetSelection(4, 1, 4, 1);
                    var type = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
                    var message = Activator.CreateInstance(type, true);
                    type.GetProperty("id").SetValue(message, f.Document.Id);
                    type.GetProperty("version").SetValue(message, f.Get<Dictionary<string, int>>("versions")[f.Document.Id]);
                    type.GetProperty("name").SetValue(message, "toggle_breakpoint");
                    type.GetProperty("line").SetValue(message, 3);
                    queued = (Task)typeof(ModernEditorWindow).GetMethod("EditorCommand", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(f.Window, new[] { message });
                });
                f.Send("step_into");
                // Drain the posted observation while the second user command owns the gate.
                Integration.MonacoRuntimeTests.Wait(() => snapshotCount >= 2, 5);
                System.Windows.Forms.Application.DoEvents();
                Assert.IsFalse(queued.IsCompleted);
                Assert.AreEqual(0, observation.ExecuteCount);
                snapshots.SetResult("null");
                ModernEditorDebugFixture.Wait(queued);
                var clock = System.Diagnostics.Stopwatch.StartNew();
                while (observation.ExecuteCount == 0 && clock.Elapsed.TotalSeconds < 5)
                { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(1); }
                Assert.AreEqual(1, step.ExecuteCount);
                Assert.AreEqual(1, toggle.ExecuteCount);
                Assert.AreEqual(1, observation.ExecuteCount,
                    "The post-step observation was lost when a queued breakpoint consumed the command gate.");
                Assert.IsFalse(UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "timer").Enabled);
                Assert.IsFalse(UiInvoke.Field<System.Windows.Forms.Timer>(f.Window, "debugTimer").Enabled);
            }
        }
        [STATestMethod]
        public void CompactLanguageRepliesVerifyFreshSourcesAndRequestSourceTextOnlyForDefinitions()
        {
            using (var f = new ModernEditorToolFixture())
            {
                var type = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
                Func<string, bool, Dictionary<string, object>> query = (known, sources) =>
                {
                    var message = f.Json.Deserialize(f.Json.Serialize(new { id = f.Document.Id, version = f.Versions[f.Document.Id], request = 4, compact = true, knownLanguage = known, includeSources = sources }), type);
                    f.Base.Scripts.Clear(); ModernEditorDebugFixture.Wait((Task)f.Private("LanguageRequest", message));
                    return f.Result(f.Base.Scripts.Single(s => s.Item1 == "languageReply").Item2[1]);
                };
                var first = query(null, false); string key = (string)first["key"];
                Assert.AreEqual(0, ((System.Collections.ICollection)first["sources"]).Count);
                var second = query(key, false); Assert.AreEqual(true, second["unchanged"]);
                f.Document.Edit(f.Document.Text + "\nPublic added As String");
                var third = query(key, false); Assert.AreNotEqual(key, third["key"]); StringAssert.Contains(f.Json.Serialize(third), "added");
                var definitions = query((string)third["key"], true); StringAssert.Contains(f.Json.Serialize(definitions), "Text");
                Assert.IsNull(f.Base.Get<EditorSyncWorker>("synchronizationWorker"));
                Assert.IsNotNull(f.Base.Get<EditorSyncWorker>("languageWorker"));
            }
        }
    }
}
