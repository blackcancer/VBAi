using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod]
        public void ContextScopeRefreshPreservesIdentityAndRejectsUnavailableCatalogues()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var picker = Get<ComboBox>(window, "scopePicker");
                var selected = picker.SelectedItem;
                Assert.IsNotNull(selected);
                runtime.Host = r => Response.Failure("catalogue unavailable");
                Assert.AreEqual(false, Call(window, "RefreshAvailableScopes", runtime.Session));
                Assert.AreSame(selected, picker.SelectedItem);
                runtime.Host = r => Response.Success(new { Invalid = true });
                Assert.AreEqual(false, Call(window, "RefreshAvailableScopes", runtime.Session));
                runtime.Host = r => Response.Success(new object[] { null, 12,
                    new { Name = "Renamed", FileName = @"C:\Temp\P.xlsm" },
                    new { Name = "Duplicate", FileName = @"C:\Temp\P.xlsm" },
                    new { Name = "Unsaved", FileName = "" },
                    new { Name = "Relative", FileName = "relative.xlsm" } });
                Assert.AreEqual(true, Call(window, "RefreshAvailableScopes", runtime.Session));
                Assert.AreEqual(4, picker.Items.Count);
                Assert.AreSame(selected, picker.SelectedItem);
                Assert.AreNotSame(picker.Items[0], picker.Items[1]);
                Assert.IsFalse(Get<bool>(window, "loadingSession"));
                picker.SelectedIndex = -1;
                Assert.AreEqual(true, Call(window, "RefreshAvailableScopes", runtime.Session));
                Assert.AreEqual(-1, picker.SelectedIndex);
                picker.SelectedIndex = 0;
                runtime.Host = r => Response.Success(new object[0]);
                Assert.AreEqual(true, Call(window, "RefreshAvailableScopes", runtime.Session));
                Assert.AreEqual(-1, picker.SelectedIndex);
                Assert.IsFalse(Get<Button>(window, "send").Enabled);
            }
        }

        [STATestMethod]
        public void ClosedScopeStaysDisabledAfterDraftAndBusyUpdatesButStopRemainsAvailable()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var picker = Get<ComboBox>(window, "scopePicker");
                var selected = picker.SelectedItem;
                runtime.Host = r => Response.Success(new object[0]);
                Assert.AreEqual(true, Call(window, "RefreshAvailableScopes", runtime.Session));
                var prompt = Get<System.Windows.Controls.TextBox>(window, "prompt");
                var send = Get<Button>(window, "send");
                prompt.Text = "Owned synthetic draft";
                Assert.IsFalse(send.Enabled, "Typing must not reactivate Send without its project.");
                Call(window, "SetBusy", false);
                Assert.IsFalse(send.Enabled, "Finishing a turn must retain the unavailable scope state.");
                prompt.Text = "";
                Call(window, "SetBusy", true);
                Assert.IsTrue(send.Enabled, "Stop must remain available even if the project closed.");
                Set(window, "stopRequested", true);
                Call(window, "UpdateBudgetControls");
                Assert.IsFalse(send.Enabled);
                prompt.Text = "Do not queue into a closed project";
                Assert.IsFalse(send.Enabled);
                Set(window, "stopRequested", false);
                Call(window, "SetBusy", false);
                Set(window, "loadingSession", true);
                picker.Items.Add(selected); picker.SelectedItem = selected;
                Set(window, "loadingSession", false);
                Call(window, "UpdateBudgetControls");
                Assert.IsTrue(send.Enabled, "A selected valid scope must retain normal draft sending.");
            }
        }

        [STATestMethod]
        public void DirtyContextRetriesFailureAndInvalidatesTheReferenceIndex()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                int reads = 0;
                runtime.Host = r => { reads++; return Response.Failure("retry"); };
                Set(window, "contextDirty", false);
                Call(window, "RefreshDirtyContext", runtime.Session);
                Assert.AreEqual(0, reads);
                Set(window, "contextDirty", true);
                Call(window, "RefreshDirtyContext", runtime.Session);
                Assert.IsTrue(Get<bool>(window, "contextDirty"));
                runtime.Host = r => { throw new IOException("read interrupted"); };
                Assert.ThrowsException<TargetInvocationException>(() => Call(window, "RefreshDirtyContext", runtime.Session));
                Assert.IsTrue(Get<bool>(window, "contextDirty"));
                runtime.Host = r => Response.Success(new object[0]);
                Set(window, "referenceIndexReady", true);
                Call(window, "RefreshDirtyContext", runtime.Session);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
                Assert.IsFalse(Get<bool>(window, "referenceIndexReady"));
                Get<System.Windows.Controls.Primitives.Popup>(window, "referencePopup").IsOpen = true;
                Set(window, "contextDirty", true);
                Call(window, "RefreshDirtyContext", runtime.Session);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
                Set(window, "contextDirty", true);
                Set(window, "referencePopup", null);
                Call(window, "RefreshDirtyContext", runtime.Session);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
            }
        }

        [STATestMethod]
        public void ContextTimerDefersBusyHiddenAndFailingHostsWithoutLosingDirtyState()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var timer = Get<Timer>(window, "contextMonitorTimer");
                Action tick = () => typeof(Timer).GetMethod("OnTick", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(timer, new object[] { EventArgs.Empty });
                int reads = 0;
                runtime.Host = r => { reads++; return Response.Success(new object[0]); };
                window.Hide(); tick(); Assert.AreEqual(0, reads);
                window.Show();
                Set(window, "busy", true); tick(); Assert.AreEqual(0, reads);
                Set(window, "busy", false); Set(window, "loadingSession", true);
                tick(); Assert.AreEqual(0, reads);
                Set(window, "loadingSession", false);
                foreach (string field in new[] { "projectEvents", "componentEvents" })
                {
                    var observer = Get<VbeCollectionEvents>(window, field);
                    var handlers = (System.Collections.IEnumerable)typeof(VbeCollectionEvents).GetField("handlers", Fields).GetValue(observer);
                    foreach (var item in handlers)
                    {
                        var handler = (Delegate)item.GetType().GetProperty("Item2").GetValue(item);
                        handler.DynamicInvoke(handler.Method.GetParameters().Length == 1 ? new object[] { null } : new object[] { null, "old" });
                        Assert.IsTrue(Get<bool>(window, "contextDirty"));
                        Set(window, "contextDirty", false);
                    }
                }
                ((Action<object>)typeof(VbeReferenceEvents).GetField("handler", Fields).GetValue(Get<VbeReferenceEvents>(window, "referenceEvents")))(null);
                Assert.IsTrue(Get<bool>(window, "contextDirty"));
                runtime.Host = r => { reads++; return Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = @"C:\Temp\P.xlsm" } } : new { }); };
                tick();
                Assert.IsNotNull(Get<ComboBox>(window, "scopePicker").SelectedItem);
                runtime.Host = r => { reads++; return Response.Success(new object[0]); };
                Set(window, "contextDirty", true); tick(); Assert.IsTrue(reads > 0);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
                runtime.Host = r => { throw new IOException("host disconnected"); };
                Set(window, "contextDirty", true); tick();
                Assert.IsTrue(Get<bool>(window, "contextDirty"));
                window.Dispose(); tick();
            }
        }

        [STATestMethod]
        public void CollectionObservationDetachesWhenTheSourceBecomesUnavailable()
        {
            using (var observer = new VbeCollectionEvents(() => { }, false, (s, i, m, h) => { }, (s, i, m, h) => { }))
            {
                Call(null, "ObserveCollectionEvents", observer, (Func<object>)(() => new object()));
                Call(null, "ObserveCollectionEvents", observer, (Func<object>)(() => null));
                Call(null, "ObserveCollectionEvents", observer, (Func<object>)(() => { throw new IOException("closed"); }));
            }
        }
    }
}
