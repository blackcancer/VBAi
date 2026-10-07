using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        // Called on the same STA that created this fixture. No COM activation or keyboard input.
        internal void ExecuteOwnedOptionsMenu(Action<object> record)
        {
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited);
            object editor = null, bars = null, command = null, main = null;
            try
            {
                editor = ((dynamic)application).VBE;
                main = ((dynamic)editor).MainWindow;
                var window = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                Assert.AreEqual(ProcessId, (int)pid); Assert.AreNotEqual(0u, tid);
                string privateDesktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
                if (!string.IsNullOrWhiteSpace(privateDesktop)) IsolatedTestDesktop.RequireCurrent(privateDesktop);
                string ownerDesktop = IsolatedTestDesktop.DesktopName(IsolatedTestDesktop.GetCurrentThreadId());
                string windowDesktop = IsolatedTestDesktop.DesktopName(tid);
                RequireOwnedWindowDesktop(privateDesktop, ownerDesktop, windowDesktop);
                bars = ((dynamic)editor).CommandBars;
                command = ((dynamic)bars).FindControl(1, 522);
                Assert.IsNotNull(command, "The native Options command must exist.");
                Assert.AreEqual(522, Convert.ToInt32(((dynamic)command).Id));
                Assert.IsTrue(Convert.ToBoolean(((dynamic)command).Enabled));
                record(new
                {
                    Phase = "OptionsMenuIntent",
                    CommandId = 522,
                    Caption = Convert.ToString(((dynamic)command).Caption),
                    VbeHwnd = window.ToInt64(),
                    OwnerThreadId = tid,
                    OwnerDesktop = ownerDesktop,
                    WindowDesktop = windowDesktop
                });
                ((dynamic)command).Execute(); // Exactly once; the modal lifetime is observed by the UIA worker.
                record(new { Phase = "OptionsMenuReturned" });
            }
            finally { Release(command); Release(bars); Release(main); Release(editor); }
        }
    }
}
