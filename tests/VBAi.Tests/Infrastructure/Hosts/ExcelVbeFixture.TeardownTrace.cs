using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        private NativeTeardownTraceGate teardownTrace;
        private int teardownOwnerThread;
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CheckRemoteDebuggerPresent(IntPtr process, out bool present);

        /// <summary>Opts in only the two declared value tests, using retained startup/loaded identities without another COM call.</summary>
        internal void EnableProcedureValuesTeardownTrace(string scenario)
        {
            if (Environment.GetEnvironmentVariable(NativeTeardownTraceGate.EnvironmentName) != "1") return;
            PreserveForDiagnosticRecovery = true;
            lock (retainedBootstraps) retainedBootstraps.Add(this);
            if (!owned || ownedProcess == null || !retainEvidence || !NativeTeardownTraceGate.AllowedScenario(scenario) ||
                Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                Convert.ToString(startupEvidence["Phase"]) != "Ready")
                throw new InvalidOperationException("Diagnostic cleanup requires durable, ready, exact owned procedure-value startup.");
            var loaded = VbeBridgeClient.Object(VbeBridgeClient.Object(startupEvidence["BridgeStatus"])["Data"]);
            if (Convert.ToInt32(loaded["HostProcessId"]) != ProcessId ||
                Convert.ToString(loaded["AssemblyModuleVersionId"]) != typeof(VbeSession).Module.ModuleVersionId.ToString("D"))
                throw new InvalidOperationException("Loaded host candidate mismatch; preserve host without cleanup.");
            IntPtr retained = ownedProcess.Handle;
            string image = ExcelOwnedProcessImage.Read(retained);
            DateTime start = ownedProcess.StartTime.ToUniversalTime();
            teardownTrace = new NativeTeardownTraceGate(new NativeTeardownTraceGate.Identity {
                ProcessId = ProcessId, ProcessStartedUtc = start.ToString("o"), Executable = image, Root = Root,
                AssemblyMvid = Convert.ToString(loaded["AssemblyModuleVersionId"]),
                AssemblyPath = Convert.ToString(loaded["AssemblyPath"]), Scenario = scenario
            }, () => !ownedProcess.HasExited && ownedProcess.StartTime.ToUniversalTime() == start &&
                    string.Equals(ExcelOwnedProcessImage.Read(retained), image, StringComparison.OrdinalIgnoreCase),
                () => ownedProcess.HasExited,
                () => CheckRemoteDebuggerPresent(retained, out bool present) ? (bool?)present : null);
            teardownOwnerThread = Thread.CurrentThread.ManagedThreadId;
            PreserveForDiagnosticRecovery = false;
        }

        private void BeginDiagnosticCleanup()
        {
            if (teardownTrace == null) return;
            try
            {
                if (Thread.CurrentThread.ManagedThreadId != teardownOwnerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Fixture cleanup is not on the original owner STA.");
                teardownTrace.BeforeCleanup();
            }
            catch (Exception error)
            {
                PreserveForDiagnosticRecovery = true;
                lock (retainedBootstraps) retainedBootstraps.Add(this);
                ShutdownDiagnostics["CleanupSuspended"] = true;
                ShutdownDiagnostics["DiagnosticFailure"] = error.ToString();
                WriteShutdownDiagnostics(ShutdownDiagnostics);
                throw;
            }
        }
    }
}
