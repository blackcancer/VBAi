using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Guards one owner-STA observation followed immediately by one export, without converting an uncertain delivery into a terminal outcome.</summary>
    internal sealed class PairedExportQualificationState
    {
        private bool ownerObservationReady;
        private string pendingCommand;
        private bool? exportSucceeded;
        internal bool ShutdownCompleted { get; private set; }
        internal string TerminalExportOutcome { get; private set; } = "NOT_REQUESTED";
        internal bool Pending { get; private set; }
        internal int ExportRequests { get; private set; }
        internal string Stage { get; private set; } = "Prepared";

        internal void Begin(string command)
        {
            if (ShutdownCompleted) throw new InvalidOperationException("The owned host lifecycle is closed; no native reuse is permitted.");
            if (Pending) throw new InvalidOperationException("A delivery is still uncertain; no follow-up or retry is permitted.");
            if (command == "export_component")
            {
                if (!ownerObservationReady || ExportRequests != 0)
                    throw new InvalidOperationException("One verified immediately preceding owner observation and a fresh export are required.");
                ExportRequests++;
                TerminalExportOutcome = "PENDING; export delivery not recorded";
            }
            // Any intervening request consumes the immediately preceding observation.
            ownerObservationReady = false;
            pendingCommand = command;
            Pending = true; Stage = command + ": Pending; delivery not recorded";
        }

        internal void Receive(string command, IDictionary<string, object> response)
        {
            if (!Pending || command != pendingCommand) throw new InvalidOperationException("The exact pending command is required.");
            object ok;
            if (response == null || !response.TryGetValue("Ok", out ok) || !(ok is bool))
                throw new InvalidOperationException("No complete terminal response; preserve exact host and read leases.");
            Pending = false;
            Stage = command + ((bool)ok ? ": TerminalSuccess" : ": TerminalFailure; original native error retained");
            if (command == "export_component")
            {
                exportSucceeded = (bool)ok;
                TerminalExportOutcome = (exportSucceeded.Value ? "SUCCESS" : "FAILED; original terminal native error retained") + "; awaiting normal owned exit";
            }
        }

        /// <summary>Closes the diagnostic only after independent, observed normal owned exit; never changes the native export outcome.</summary>
        internal void CompleteOwnedShutdown(IDictionary<string, object> shutdown)
        {
            if (Pending || ExportRequests != 1 || !exportSucceeded.HasValue || ShutdownCompleted || shutdown == null ||
                !shutdown.ContainsKey("Exited") || !Equals(true, shutdown["Exited"]) ||
                !shutdown.ContainsKey("ExitCodeHex") || !Equals("0x00000000", shutdown["ExitCodeHex"]) ||
                !shutdown.ContainsKey("ForcedTermination") || !Equals(false, shutdown["ForcedTermination"]) ||
                (shutdown.ContainsKey("CleanupSuspended") && Equals(true, shutdown["CleanupSuspended"])))
                throw new InvalidOperationException("A terminal export and independent normal owned exit0 are required to close the diagnostic.");
            ShutdownCompleted = true;
            TerminalExportOutcome = (exportSucceeded.Value ? "SUCCESS" : "FAILED; original terminal native error retained") + "; normal owned exit verified0";
        }

        /// <summary>Requires actual host PID/native TID/STA/MVID and the exact four synthetic targets, while preserving differing visibility results.</summary>
        internal void VerifyOwner(IDictionary<string, object> observed, int pid, uint tid, string mvid, string[] paths)
        {
            ownerObservationReady = false;
            if (Pending || Stage != PathVisibilityDiagnostic.CommandName + ": TerminalSuccess" || ExportRequests != 0)
                throw new InvalidOperationException("A successful diagnostic response must immediately precede export preparation.");
            if (pid <= 0 || tid == 0 || Convert.ToInt32(observed["ProcessId"]) != pid ||
                Convert.ToInt32(observed["FinalProcessId"]) != pid || Convert.ToUInt32(observed["NativeThreadId"]) != tid ||
                Convert.ToUInt32(observed["FinalNativeThreadId"]) != tid || !Equals("STA", observed["Apartment"]) ||
                !Equals(mvid, observed["AssemblyMvid"]))
                throw new InvalidOperationException("Diagnostic owner process/thread/apartment/candidate identity changed.");
            foreach (string phase in new[] { "EffectiveTokenBefore", "EffectiveTokenAfter" })
                if (!Equals("READ", VbeBridgeClient.Object(observed[phase])["State"]))
                    throw new InvalidOperationException("Partial effective-token evidence is not a complete observation.");
            var rows = ((object[])observed["Paths"]).Select(VbeBridgeClient.Object).ToArray();
            if (paths == null || paths.Length != 4 || !paths.SequenceEqual(rows.Select(row => (string)row["Path"])))
                throw new InvalidOperationException("The diagnostic allowlist must remain exactly two owned directories and their fixed synthetic files.");
            foreach (var row in rows)
                if (!row.ContainsKey("NativeLastError") || !row.ContainsKey("NativeAttributes") || !row.ContainsKey("NativeSucceeded") ||
                    !(row.ContainsKey("ManagedAttributes") || row.ContainsKey("ManagedAttributesError")))
                    throw new InvalidOperationException("Both native and managed attribute outcomes must be recorded, including denials.");
            ownerObservationReady = true;
        }
    }
}
