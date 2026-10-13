using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Integration
{
    /// <summary>Observes real application open/close callbacks without cancelling events or changing macro security.</summary>
    internal sealed class OfficeCoverageEventObserver : IDisposable
    {
        private delegate void BeforeClose(object document, ref bool cancel);
        private readonly OfficeVbeFixture fixture;
        private readonly object application;
        private readonly Guid eventInterface;
        private readonly int openId, closeId;
        private readonly Action<object> opened;
        private readonly BeforeClose closing;
        private readonly Func<int, object, IDictionary<string, object>> originalDispatch, pumpingDispatch;
        private bool disposed;
        private readonly EventBoundary boundary;
        private readonly IOfficeCoverageEventDispatcher dispatcher;
        private static readonly List<OfficeCoverageEventObserver> retained = new List<OfficeCoverageEventObserver>();
        internal OfficeCoverageEventTrace Trace { get; }

        /// <summary>Native event boundaries, replaceable only by mirror tests that do not activate Office.</summary>
        internal sealed class EventBoundary
        {
            internal Func<string, object> Acquire = Marshal.GetActiveObject;
            internal Func<string, object, IntPtr> Window = (host, value) => host == "Word"
                ? VbaTestWordValuesHost.ReadApplicationWindow(value) : PowerPointWindow.Read(value);
            internal Func<IntPtr, uint> Owner = window => { uint pid; return GetWindowThreadProcessId(window, out pid) == 0 ? 0 : pid; };
            internal Action<object, Guid, int, Delegate> Add = ComEventsHelper.Combine;
            internal Action<object, Guid, int, Delegate> Remove = (value, events, id, callback) => ComEventsHelper.Remove(value, events, id, callback);
            internal Action<object> Release = value => Marshal.ReleaseComObject(value);
            internal Func<IOfficeCoverageEventDispatcher> CreateDispatcher = () => new OfficeCoverageEventDispatcher();
        }

        internal OfficeCoverageEventObserver(OfficeVbeFixture fixture, Action observeCopyWarning = null, EventBoundary boundary = null)
        {
            this.fixture = fixture ?? throw new ArgumentNullException(nameof(fixture));
            this.boundary = boundary ?? new EventBoundary();
            if (fixture.Kind != "Word" && fixture.Kind != "PowerPoint") throw new ArgumentException("Word or PowerPoint is required.");
            Assert.IsFalse(fixture.NativeExecutionUnsettled);
            // Published dispatch IDs verified against the installed Office 15 PIAs:
            // Word.ApplicationEvents4.DocumentOpen/DocumentBeforeClose and
            // PowerPoint.EApplication.PresentationOpen/PresentationBeforeClose.
            eventInterface = fixture.Kind == "Word" ? new Guid("00020a01-0000-0000-c000-000000000046")
                : new Guid("914934c2-5a91-11cf-8700-00aa0060263b");
            openId = fixture.Kind == "Word" ? 4 : 2006;
            closeId = fixture.Kind == "Word" ? 6 : 2025;
            opened = document => Observe("Open", document, false);
            closing = (object document, ref bool cancel) => Observe("BeforeClose", document, cancel);
            bool openEntered = false, closeEntered = false;
            try
            {
                dispatcher = this.boundary.CreateDispatcher();
                Assert.IsNotNull(dispatcher);
                application = this.boundary.Acquire(fixture.Kind + ".Application");
                Assert.IsNotNull(application);
                IntPtr window = this.boundary.Window(fixture.Kind, application);
                Assert.AreNotEqual(IntPtr.Zero, window);
                Assert.AreEqual((uint)fixture.ProcessId, this.boundary.Owner(window));
                Trace = new OfficeCoverageEventTrace(fixture.DocumentPath, Thread.CurrentThread.ManagedThreadId);
                openEntered = true;
                this.boundary.Add(application, eventInterface, openId, opened);
                closeEntered = true;
                this.boundary.Add(application, eventInterface, closeId, closing);
            }
            catch (Exception primary)
            {
                var errors = new List<Exception> { primary };
                if (closeEntered) try { this.boundary.Remove(application, eventInterface, closeId, closing); } catch (Exception error) { errors.Add(error); }
                if (openEntered) try { this.boundary.Remove(application, eventInterface, openId, opened); } catch (Exception error) { errors.Add(error); }
                if (errors.Count == 1 && application != null)
                    try { this.boundary.Release(application); } catch (Exception error) { errors.Add(error); }
                if (errors.Count == 1 && dispatcher != null)
                    try { dispatcher.Dispose(); } catch (Exception error) { errors.Add(error); }
                if (errors.Count > 1)
                {
                    fixture.NativeExecutionUnsettled = true;
                    lock (retained) retained.Add(this);
                    throw new AggregateException("Event attachment and its ownership recovery did not settle; the original host is retained.", errors);
                }
                throw;
            }
            originalDispatch = fixture.Dispatch;
            pumpingDispatch = (pid, request) =>
            {
                Assert.AreEqual(Trace.OwnerThread, Thread.CurrentThread.ManagedThreadId);
                // Only the TCP client runs on a worker. RPC event deliveries are synchronously
                // marshalled back before their borrowed arguments are read or callbacks return.
                return ReadPumping(() => VbeBridgeClient.Read(pid, request),
                    () => { Application.DoEvents(); observeCopyWarning?.Invoke(); }, Thread.Sleep);
            };
            fixture.Dispatch = pumpingDispatch;
        }

        private void Observe(string kind, object document, bool cancel)
        {
            // Never throw into the host's event dispatch, alter Cancel, release a borrowed event
            // argument, save a document, or call another macro from an observation callback.
            try
            {
                int deliveryThread = Thread.CurrentThread.ManagedThreadId;
                dispatcher.Invoke(() =>
                {
                    Assert.AreEqual(Trace.OwnerThread, Thread.CurrentThread.ManagedThreadId);
                    Trace.Record(kind, Convert.ToString(((dynamic)document).FullName), cancel,
                        Thread.CurrentThread.ManagedThreadId, deliveryThread);
                });
            }
            catch (Exception error) { Trace.RecordFailure(error); }
        }

        public void Dispose()
        {
            if (disposed) return;
            Assert.AreEqual(Trace.OwnerThread, Thread.CurrentThread.ManagedThreadId);
            disposed = true;
            if (fixture.NativeExecutionUnsettled)
            {
                // Keep the sink and application reference with uncertain native work; disposal
                // cannot silently detach a still active callback or mask the original failure.
                lock (retained) retained.Add(this);
                return;
            }
            try
            {
                Assert.AreSame(pumpingDispatch, fixture.Dispatch, "Another dispatch owner replaced the event qualification scope.");
                fixture.Dispatch = originalDispatch;
                boundary.Remove(application, eventInterface, openId, opened);
                boundary.Remove(application, eventInterface, closeId, closing);
                boundary.Release(application);
                dispatcher.Dispose();
            }
            catch
            {
                fixture.NativeExecutionUnsettled = true;
                lock (retained) retained.Add(this);
                throw;
            }
        }
        /// <summary>Pumps the calling STA while a single transport-only operation completes; never retries a request.</summary>
        internal static T ReadPumping<T>(Func<T> network, Action pump, Action<int> pause)
        {
            var response = Task.Run(network);
            while (!response.IsCompleted) { pump(); pause(10); }
            return response.GetAwaiter().GetResult();
        }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    }

    /// <summary>Independent event oracle for the two lifecycle callbacks of one completed copy.</summary>
    internal sealed class OfficeCoverageEventTrace
    {
        internal sealed class Entry
        {
            public string Kind, Path;
            public bool Cancel;
            public int Thread;
            public int DeliveryThread;
        }
        private readonly string original;
        internal int OwnerThread { get; }
        internal List<Entry> Events { get; } = new List<Entry>();
        internal List<string> Failures { get; } = new List<string>();
        internal OfficeCoverageEventTrace(string original, int ownerThread)
        { this.original = System.IO.Path.GetFullPath(original); OwnerThread = ownerThread; }
        internal void Record(string kind, string path, bool cancel, int thread, int? deliveryThread = null)
        {
            lock (Events)
            {
                if (Events.Count >= 32) { RecordFailure(new InvalidOperationException("Event bound exceeded.")); return; }
                Events.Add(new Entry { Kind=kind, Path=path, Cancel=cancel, Thread=thread, DeliveryThread=deliveryThread ?? thread });
            }
        }
        internal void RecordFailure(Exception error)
        {
            lock (Failures)
            {
                string text = error.ToString();
                if (!Failures.Contains(text)) Failures.Add(text);
            }
        }
        internal void RequireCompletedCopy(string copy)
        {
            string path = System.IO.Path.GetFullPath(copy);
            Assert.AreNotEqual(original, path, true, "The lifecycle oracle must refer to a distinct copy.");
            Assert.AreEqual(0, Failures.Count, string.Join("\n", Failures));
            CollectionAssert.AreEqual(new[] { "Open", "BeforeClose" }, Events.Select(item => item.Kind).ToArray(),
                "Exactly one open and one close callback must belong to the completed copy.");
            foreach (var item in Events)
            {
                Assert.AreEqual(OwnerThread, item.Thread, "Event dispatch must stay on the owned STA.");
                Assert.IsFalse(string.IsNullOrEmpty(item.Path), "An event must identify its document.");
                Assert.AreEqual(path, System.IO.Path.GetFullPath(item.Path), true, "Another document received a lifecycle event.");
                Assert.IsFalse(item.Cancel, "Close was already cancelled by another handler; no successful close claim is permitted.");
            }
        }
    }
}
