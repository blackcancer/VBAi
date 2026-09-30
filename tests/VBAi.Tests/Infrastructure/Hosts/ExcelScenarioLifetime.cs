using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Observes owned Excel shutdown after a native scenario has released its local COM references.</summary>
    internal sealed class ExcelScenarioLifetime
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        private Process process;
        private readonly int[] existingIds;
        internal bool OwnsApplication { get; private set; }

        private ExcelScenarioLifetime()
        {
            var existing = Process.GetProcessesByName("EXCEL");
            try { existingIds = existing.Select(item => item.Id).ToArray(); }
            finally { foreach (var item in existing) item.Dispose(); }
        }

        internal void Capture(object application)
        {
            uint id;
            GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out id);
            Assert.IsTrue(id != 0 && !existingIds.Contains((int)id), "Excel did not create an owned process.");
            OwnsApplication = true;
            process = Process.GetProcessById((int)id);
            _ = process.Handle; // Keep fast exits observable through the retained handle.
        }

        internal static void Run(Action<ExcelScenarioLifetime> scenario)
        {
            var lifetime = new ExcelScenarioLifetime();
            ExceptionDispatchInfo failure = Execute(scenario, lifetime);
            try { lifetime.VerifyExit(); }
            catch (Exception cleanup)
            {
                // Preserve the scenario's primary assertion and report shutdown as independent evidence.
                Console.WriteLine("EXCEL SHUTDOWN FAILURE: " + cleanup);
                if (failure == null) throw;
            }
            failure?.Throw();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static ExceptionDispatchInfo Execute(Action<ExcelScenarioLifetime> scenario, ExcelScenarioLifetime lifetime)
        {
            try { scenario(lifetime); return null; }
            catch (Exception error) { return ExceptionDispatchInfo.Capture(error); }
        }

        private void VerifyExit()
        {
            // Dynamic Office access creates temporary RCWs. Collect only after the scenario frame
            // (including adapters and renderer locals) has returned, not while it still roots them.
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
            if (process == null) return;
            using (process)
            {
                bool exited = process.WaitForExit(10000);
                Console.WriteLine("EXCEL SHUTDOWN: PID=" + process.Id + "; exited=" + exited +
                    (exited ? "; exitCode=0x" + unchecked((uint)process.ExitCode).ToString("X8") : "; left running for diagnosis"));
                Assert.IsTrue(exited, "Owned Excel did not exit within 10 seconds after Quit and COM release; no process was killed. PID: " + process.Id);
                Assert.AreEqual(0, process.ExitCode, "Owned Excel exited abnormally. PID: " + process.Id);
            }
        }
    }
}
