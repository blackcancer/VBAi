using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

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
                CompleteScopeLoad(window);
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

        /// <summary>A COM callback cannot start another polling tick while the original tick owns the timer.</summary>
        [STATestMethod]
        public void ContextPollingRejectsReentryAndRestartsAfterCompletion()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show();
                var timer = Get<Timer>(window, "contextMonitorTimer");
                timer.Stop(); Set(window, "contextDirty", false);
                int reads = 0, nested = 0;
                runtime.Host = request =>
                {
                    reads++;
                    Assert.IsTrue(Get<bool>(window, "contextMonitorReading"));
                    Assert.IsFalse(timer.Enabled, "The timer must remain suspended while COM pumps messages.");
                    int before = reads;
                    Call(window, "PollContext", runtime.Session); nested++;
                    Assert.AreEqual(before, reads, "A nested tick must return before another host read.");
                    return Response.Success(new object[0]);
                };
                Call(window, "PollContext", runtime.Session);
                Assert.AreEqual(2, reads); Assert.AreEqual(2, nested);
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading")); Assert.IsTrue(timer.Enabled);
            }
        }

        /// <summary>A failed host read releases the guard and permits the next regular tick.</summary>
        [STATestMethod]
        public void ContextPollingRecoversFromAnExceptionWithoutLosingDirtyState()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); var timer = Get<Timer>(window, "contextMonitorTimer"); timer.Stop();
                Set(window, "contextDirty", true);
                runtime.Host = request => { throw new IOException("owned read failed"); };
                Call(window, "PollContext", runtime.Session);
                Assert.IsTrue(Get<bool>(window, "contextDirty"));
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading")); Assert.IsTrue(timer.Enabled);
                int reads = 0;
                runtime.Host = request => { reads++; return Response.Success(new object[0]); };
                Call(window, "PollContext", runtime.Session);
                Assert.IsTrue(reads > 0); Assert.IsFalse(Get<bool>(window, "contextDirty"));
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading")); Assert.IsTrue(timer.Enabled);
            }
        }

        /// <summary>Pending scope and runtime transitions defer passive reads without clearing an event.</summary>
        [STATestMethod]
        public void ContextPollingDefersScopeAndRuntimeTransitions()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); int reads = 0;
                runtime.Host = request => { reads++; return Response.Success(new object[0]); };
                foreach (string field in new[] { "busy", "loadingSession", "loadingScope", "runtimeDisposed" })
                {
                    Set(window, field, true); Set(window, "contextDirty", true);
                    try
                    {
                        Call(window, "PollContext", runtime.Session);
                        Assert.AreEqual(0, reads, field);
                        Assert.IsTrue(Get<bool>(window, "contextDirty"), field);
                        Assert.IsFalse(Get<bool>(window, "contextMonitorReading"), field);
                    }
                    finally { Set(window, field, false); }
                }
            }
        }

        /// <summary>Native inspection keeps a disposable copy out of passive context acquisition.</summary>
        [STATestMethod]
        public void ContextPollingDefersNativeInspectionAndReconcilesAfterItsOwnerReleasesIt()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); Get<Timer>(window, "contextMonitorTimer").Stop();
                Set(window, "contextDirty", true); int reads = 0;
                runtime.Host = request => { reads++; return Response.Success(new object[0]); };
                using (new VbeDebugInspection())
                {
                    Call(window, "PollContext", runtime.Session);
                    Assert.AreEqual(false, Call(window, "RefreshAvailableScopes", runtime.Session));
                    Assert.AreEqual(0, reads);
                    Assert.IsTrue(Get<bool>(window, "contextDirty"));
                    Assert.IsFalse(Get<bool>(window, "contextMonitorReading"));
                }
                Call(window, "PollContext", runtime.Session);
                Assert.IsTrue(reads > 0);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
            }
        }

        /// <summary>Disposal pumped by the first read prevents later reads and timer resurrection.</summary>
        [STATestMethod]
        public void ContextPollingDoesNotReadOrRestartAfterReentrantDisposal()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); var timer = Get<Timer>(window, "contextMonitorTimer"); timer.Stop();
                Set(window, "contextDirty", false); int reads = 0;
                runtime.Host = request => { reads++; window.Dispose(); return Response.Success(new object[0]); };
                Call(window, "PollContext", runtime.Session);
                Assert.AreEqual(1, reads); Assert.IsTrue(window.IsDisposed);
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading")); Assert.IsFalse(timer.Enabled);
            }
        }

        /// <summary>Every deferred read rechecks visibility and active turns after COM pumps a transition.</summary>
        [STATestMethod]
        public void ContextPollingStopsNativeReadsWhenATurnStartsOrTheWindowHides()
        {
            foreach (bool hide in new[] { false, true })
                using (var runtime = new RuntimeScope())
                using (var window = LoadedWindow(runtime.Session))
                {
                    window.Show(); var timer = Get<Timer>(window, "contextMonitorTimer"); timer.Stop();
                    Set(window, "contextDirty", false); int reads = 0;
                    runtime.Host = request =>
                    {
                        reads++;
                        if (hide) window.Hide(); else Set(window, "busy", true);
                        return Response.Success(new object[0]);
                    };
                    Call(window, "PollContext", runtime.Session);
                    Assert.AreEqual(1, reads);
                    Assert.IsFalse(Get<bool>(window, "contextMonitorReading")); Assert.IsTrue(timer.Enabled);
                    Set(window, "busy", false);
                }
        }

        /// <summary>Records when native collection acquisition begins, independently of the injected catalogue response.</summary>
        public sealed class ContextPollingEventHost
        {
            /// <summary>Projects visible to the event-source resolver.</summary>
            public readonly List<VbeSessionTests.FakeProject> Projects = new List<VbeSessionTests.FakeProject>();
            /// <summary>Observation executed whenever a collection getter would enter native code.</summary>
            public Action SourceObserved;
            /// <summary>Returns the current project collection after recording the acquisition boundary.</summary>
            public List<VbeSessionTests.FakeProject> VBProjects { get { SourceObserved?.Invoke(); return Projects; } }
        }

        /// <summary>A removed scope is reconciled before native event sources can resolve the stale selector.</summary>
        [STATestMethod]
        public void ContextPollingReconcilesClosedScopeBeforeNativeSourceAcquisition()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); Get<Timer>(window, "contextMonitorTimer").Stop();
                Assert.IsNotNull(Get<ComboBox>(window, "scopePicker").SelectedItem);
                var source = new ContextPollingEventHost();
                var session = new VbeSession(source);
                bool reconciled = false; int acquisitions = 0;
                runtime.Host = request =>
                {
                    if (request.Command == "list_projects") reconciled = true;
                    return Response.Success(new object[0]);
                };
                source.SourceObserved = () =>
                {
                    acquisitions++;
                    Assert.IsTrue(reconciled, "Native collection acquisition must follow the catalogue refresh.");
                    Assert.IsNull(Get<ComboBox>(window, "scopePicker").SelectedItem, "The removed scope must not reach a native resolver.");
                };
                Set(window, "contextDirty", true);
                Call(window, "PollContext", session);
                Assert.AreEqual(1, acquisitions);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading"));
            }
        }

        /// <summary>An unavailable catalogue preserves the event and does not resolve native sources from stale data.</summary>
        [STATestMethod]
        public void ContextPollingDefersNativeSourcesAfterFailedScopeRefresh()
        {
            foreach (bool invalidData in new[] { false, true })
                using (var runtime = new RuntimeScope())
                using (var window = LoadedWindow(runtime.Session))
                {
                    window.Show(); Get<Timer>(window, "contextMonitorTimer").Stop();
                    var selected = Get<ComboBox>(window, "scopePicker").SelectedItem;
                    int acquisitions = 0, reads = 0;
                    var source = new ContextPollingEventHost { SourceObserved = () => acquisitions++ };
                    runtime.Host = request => { reads++; return invalidData ? Response.Success(new { Invalid = true }) : Response.Failure("closed catalogue"); };
                    Set(window, "contextDirty", true);
                    Call(window, "PollContext", new VbeSession(source));
                    Assert.AreEqual(1, reads); Assert.AreEqual(0, acquisitions);
                    Assert.AreSame(selected, Get<ComboBox>(window, "scopePicker").SelectedItem);
                    Assert.IsTrue(Get<bool>(window, "contextDirty"));
                    Assert.IsFalse(Get<bool>(window, "contextMonitorReading"));
                    Assert.IsTrue(Get<Timer>(window, "contextMonitorTimer").Enabled);
                }
        }

        /// <summary>A project event pumped during collection acquisition forces another reconciliation before component sources.</summary>
        [STATestMethod]
        public void ContextPollingReconcilesAnEventRaisedDuringCollectionAcquisition()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); Get<Timer>(window, "contextMonitorTimer").Stop();
                Set(window, "contextDirty", false);
                int acquisitions = 0, refreshed = 0;
                var source = new ContextPollingEventHost();
                source.SourceObserved = () => { acquisitions++; Set(window, "contextDirty", true); };
                runtime.Host = request => { if (request.Command == "list_projects") refreshed++; return Response.Success(new object[0]); };
                Call(window, "PollContext", new VbeSession(source));
                Assert.AreEqual(1, acquisitions);
                Assert.AreEqual(2, refreshed, "One refresh precedes component acquisition; the later read is the normal global poll.");
                Assert.IsNull(Get<ComboBox>(window, "scopePicker").SelectedItem);
                Assert.IsFalse(Get<bool>(window, "contextDirty"));
                Assert.IsFalse(Get<bool>(window, "contextMonitorReading"));
            }
        }
    }
}
