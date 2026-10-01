using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in rendezvous for one owned native export; diagnostic timeouts are distinct from export failures.</summary>
    internal static class NativeExportTraceGate
    {
        internal static void WaitArmed(int ownedPid, string destination, string output, IDictionary<string, object> loaded,
            IDictionary<string, object> report)
        {
            if (Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_EXPORT_TRACE_GATE") != "1") return;
            report["DiagnosticTraceRequested"] = true;
            report["Stage"] = "WAITING_FOR_DIAGNOSTIC_TRACE_BEFORE_ANY_EXPORT";
            using (var process = Process.GetProcessById(ownedPid))
                File.WriteAllText(Path.Combine(output, "trace.pending.json"), new JavaScriptSerializer().Serialize(new {
                    ProcessId = ownedPid, ProcessStartedUtc = process.StartTime.ToUniversalTime().ToString("o"),
                    Destination = destination, AssemblyMvid = loaded["AssemblyModuleVersionId"], AssemblyPath = loaded["AssemblyPath"] }));
            Wait(Path.Combine(output, "trace.armed"), ownedPid, Convert.ToString(loaded["AssemblyModuleVersionId"]),
                "Diagnostic trace did not arm before any export; this is a diagnostic setup timeout, not native export failure.");
        }

        internal static void CompleteAndWaitDetached(int ownedPid, string output, string mvid)
        {
            if (Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_EXPORT_TRACE_GATE") != "1") return;
            File.WriteAllText(Path.Combine(output, "trace.export-complete"), "Single export returned or threw; detach before host disposal.");
            Wait(Path.Combine(output, "trace.detached"), ownedPid, mvid,
                "Diagnostic debugger detachment was not proven. Preserve native error separately; no native retry or host termination.");
        }

        private static void Wait(string path, int pid, string mvid, string error)
        {
            var watch = Stopwatch.StartNew();
            while (watch.Elapsed < TimeSpan.FromSeconds(45))
            {
                if (File.Exists(path))
                {
                    string observed = File.ReadAllText(path);
                    if (observed != pid + " " + mvid) throw new InvalidOperationException("Diagnostic trace marker identity mismatch.");
                    return;
                }
                Thread.Sleep(100);
            }
            throw new TimeoutException(error);
        }
    }
}
