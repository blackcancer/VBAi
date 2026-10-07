using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OutlookVbaTestFixture
    {
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        internal static string RequirePrivateOutlookExecutable(string executable, string desktop, string workerDesktop)
        {
            if (string.IsNullOrEmpty(desktop) || desktop != workerDesktop)
                throw new InvalidOperationException("Exact private-desktop pairing is required before Outlook launch.");
            string path = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(executable);
            if (Path.GetFileName(path) != "OUTLOOK.EXE" || !File.Exists(path))
                throw new InvalidOperationException("An existing explicit Outlook executable is required.");
            return path;
        }

        private void StartOnPrivateDesktop(string desktop)
        {
            string executable = RequirePrivateOutlookExecutable(Environment.GetEnvironmentVariable("VBAi_TEST_OUTLOOK_EXE"),
                desktop, Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"));
            IsolatedTestDesktop.RequireCurrent(desktop);
            Save("private-launch-intent.json", new { Executable = executable, Desktop = desktop, InvocationLimit = 1 });
            privateDesktopChild = IsolatedTestDesktop.Launch(executable, new string[0], Root, desktop);
            ProcessId = privateDesktopChild.ProcessId;
            process = Process.GetProcessById(ProcessId); _ = process.Handle;
            Save("private-launch-returned.json", new { ProcessId, privateDesktopChild.ThreadId, Desktop = desktop });
            var wait = Stopwatch.StartNew();
            IntPtr hwnd = IntPtr.Zero; uint thread = 0;
            while (wait.Elapsed.TotalSeconds < 45)
            {
                Assert.IsFalse(privateDesktopChild.Wait(0), "Original Outlook exited before attachment.");
                process.Refresh(); hwnd = process.MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                {
                    uint pid; thread = GetWindowThreadProcessId(hwnd, out pid);
                    if (pid != ProcessId || thread == 0)
                        throw new InvalidOperationException("Outlook native window is outside its original private desktop.");
                    IsolatedTestDesktop.RequireOfficeWindowInventory(desktop, (uint)ProcessId, true, hwnd);
                    break;
                }
                System.Windows.Forms.Application.DoEvents(); Thread.Sleep(40);
            }
            if (hwnd == IntPtr.Zero) throw new TimeoutException("Owned Outlook window was not observed; no activation fallback.");
            // Only a read of the ROT is repeated during readiness; no activation, item or native action is retried.
            while (application == null && wait.Elapsed.TotalSeconds < 60)
            {
                RequireNoForeignOutlook(ProcessId);
                try { application = Marshal.GetActiveObject("Outlook.Application"); }
                catch (COMException error) when ((uint)error.ErrorCode == 0x800401E3) { Thread.Sleep(40); }
            }
            if (application == null) throw new TimeoutException("The sole explicitly launched Outlook is not in the ROT.");
            Save("private-attachment.json", new
            {
                ProcessId,
                Window = hwnd.ToInt64(),
                NativeThread = thread,
                Desktop = desktop,
                SoleProcessVerified = true,
                ComActivationInvocations = 0
            });
        }

        private static void RequireNoForeignOutlook(int ownedPid)
        {
            var live = Process.GetProcessesByName("OUTLOOK");
            try { Assert.AreEqual(1, live.Length); Assert.AreEqual(ownedPid, live[0].Id); }
            finally { foreach (var current in live) current.Dispose(); }
        }
    }
}
