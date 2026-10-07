using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    public sealed partial class ImportedFormReadbackTests
    {
        [STATestMethod]
        public void TaskYieldCanResumeBeforeTheOriginalNativeDispatchReturns()
        {
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var spy = new DispatchSpy();
            try
            {
                var task = VbeUiTask.Run(async () =>
                {
                    Attach(spy);
                    await gate.Task;
                    int calls = spy.Calls, returned = spy.Returns;
                    Assert.AreEqual(1, spy.Depth);
                    await Task.Yield();
                    Assert.AreEqual(calls, spy.Calls);
                    Assert.AreEqual(returned, spy.Returns);
                    Assert.AreEqual(1, spy.Depth, "The yield continuation is drained inside the original native dispatch.");
                    return true;
                });
                gate.SetResult(true); Pump(task); Assert.IsTrue(task.GetAwaiter().GetResult());
            }
            finally { if (spy.Handle != IntPtr.Zero) spy.ReleaseHandle(); }
        }

        [STATestMethod]
        public void OneTimerPulseUnwindsNativeDispatchAndDisposesBeforeStrictOwnerReadback()
        {
            VerifyRealPulse("exact");
        }

        [STATestMethod]
        public void OneTimerPulseDoesNotAcceptAStillRoundedFrameFont()
        {
            VerifyRealPulse("mismatch");
        }

        [STATestMethod]
        public void GuardFailureAfterRealPulsePreservesItsErrorWithoutCapture()
        {
            VerifyRealPulse("guard");
        }

        [STATestMethod]
        public void ExportFailureAfterRealPulsePreservesItsErrorWithoutAnotherReadback()
        {
            VerifyRealPulse("export");
        }

        [STATestMethod]
        public void MissingOwnerContextRefusesThePulseBeforeAnyTimerMessage()
        {
            var original = SynchronizationContext.Current;
            var filter = new TimerMessages(); Application.AddMessageFilter(filter);
            try
            {
                SynchronizationContext.SetSynchronizationContext(null);
                var failure = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormReadback.PulseOwnerAsync().GetAwaiter().GetResult());
                StringAssert.Contains(failure.Message, "owning STA");
                Assert.AreEqual(0, filter.Count);
            }
            finally { SynchronizationContext.SetSynchronizationContext(original); Application.RemoveMessageFilter(filter); }
        }

        [TestMethod]
        public void NonStaContextRefusesThePulseWithoutStartingNativeReadback()
        {
            Exception observed = null;
            var worker = new Thread(() =>
            {
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
                try { ImportedFormReadback.PulseOwnerAsync().GetAwaiter().GetResult(); }
                catch (Exception error) { observed = error; }
            });
            worker.SetApartmentState(ApartmentState.MTA); worker.Start();
            Assert.IsTrue(worker.Join(3000)); Assert.IsInstanceOfType(observed, typeof(InvalidOperationException));
            StringAssert.Contains(observed.Message, "owning STA");
        }

        private static void VerifyRealPulse(string outcome)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            var failure = new InvalidOperationException("original " + outcome);
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var spy = new DispatchSpy(); var filter = new TimerMessages();
            int owner = Thread.CurrentThread.ManagedThreadId, guards = 0, captures = 0;
            Application.AddMessageFilter(filter);
            try
            {
                var task = VbeUiTask.Run(async () =>
                {
                    Attach(spy); var context = SynchronizationContext.Current;
                    await gate.Task;
                    int originalCalls = spy.Calls, originalReturns = spy.Returns;
                    Assert.AreEqual(1, spy.Depth);
                    return await ImportedFormReadback.VerifyAsync(target, stale, true,
                        () =>
                        {
                            Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                            Assert.AreSame(context, SynchronizationContext.Current);
                            Assert.IsTrue(spy.Returns > originalReturns, "The initial native dispatch must have returned before any readback guard.");
                            Assert.IsTrue(spy.Calls > originalCalls);
                            Assert.AreEqual(1, spy.Depth); Assert.AreEqual(1, spy.MaximumDepth);
                            Assert.AreEqual(1, filter.Count);
                            Assert.IsFalse(IsWindow(filter.Window), "The one-shot native timer window must be disposed before guard/capture.");
                            guards++; if (outcome == "guard") throw failure;
                        },
                        () =>
                        {
                            Assert.AreEqual(1, guards); captures++;
                            if (outcome == "export") throw failure;
                            return outcome == "mismatch" ? stale : target;
                        });
                });
                gate.SetResult(true); Pump(task);
                if (outcome == "guard" || outcome == "export")
                    Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult()));
                else Assert.AreEqual(outcome != "mismatch", task.GetAwaiter().GetResult());
                Assert.AreEqual(1, guards); Assert.AreEqual(outcome == "guard" ? 0 : 1, captures);
                var extra = Stopwatch.StartNew();
                while (extra.ElapsedMilliseconds < 80) { Application.DoEvents(); Thread.Sleep(1); }
                Assert.AreEqual(1, filter.Count, "The disposed timer must never deliver a second tick.");
                Assert.IsFalse(IsWindow(filter.Window));
                Assert.AreEqual(1, guards); Assert.AreEqual(outcome == "guard" ? 0 : 1, captures);
            }
            finally
            {
                Application.RemoveMessageFilter(filter);
                if (spy.Handle != IntPtr.Zero) spy.ReleaseHandle();
            }
        }

        private static void Attach(DispatchSpy spy)
        {
            var context = SynchronizationContext.Current;
            var control = (Control)context.GetType().GetField("dispatcher", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(context);
            spy.AssignHandle(control.Handle);
        }

        private static void Pump(Task pending)
        {
            var clock = Stopwatch.StartNew();
            while (!pending.IsCompleted && clock.ElapsedMilliseconds < 3000) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(pending.IsCompleted, "The isolated managed WinForms pulse did not finish within its message-pump budget.");
        }

        private sealed class DispatchSpy : NativeWindow
        {
            internal int Depth, MaximumDepth, Calls, Returns;
            protected override void WndProc(ref Message message)
            {
                if (message.Msg < 0xC000) { base.WndProc(ref message); return; }
                Depth++; MaximumDepth = Math.Max(MaximumDepth, Depth); Calls++;
                try { base.WndProc(ref message); }
                finally { Depth--; Returns++; }
            }
        }

        private sealed class TimerMessages : IMessageFilter
        {
            internal int Count; internal IntPtr Window;
            public bool PreFilterMessage(ref Message message)
            {
                if (message.Msg == 0x0113 && (Window == IntPtr.Zero || Window == message.HWnd)) { Window = message.HWnd; Count++; }
                return false;
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr window);
    }
}
