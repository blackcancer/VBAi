using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Disabled-by-default rendezvous for the single cleanup of two owned procedure-value fixtures.</summary>
    internal sealed class NativeTeardownTraceGate
    {
        internal const string EnvironmentName = "VBAi_TEST_EXCEL_TEARDOWN_TRACE_GATE";
        internal sealed class Identity
        {
            public int ProcessId;
            public string ProcessStartedUtc, Executable, Root, AssemblyMvid, AssemblyPath, Scenario;
            public string Nonce = Guid.NewGuid().ToString("N");
        }

        private readonly Identity identity;
        private readonly Func<bool> stillOwned, exited;
        private readonly Func<bool?> debuggerPresent;
        private readonly HashSet<string> pending = new HashSet<string>(StringComparer.Ordinal);
        private bool uncertain;
        private bool armingStarted;
        private int emissions, invocationRequests, cleanupAttempts;
        internal Func<DateTime> Clock = () => DateTime.UtcNow;
        internal Action WaitTick = () => Thread.Sleep(100);
        internal TimeSpan Deadline = TimeSpan.FromSeconds(45);

        internal NativeTeardownTraceGate(Identity identity, Func<bool> stillOwned, Func<bool> exited, Func<bool?> debuggerPresent)
        {
            ValidateIdentity(identity);
            this.identity = identity;
            this.stillOwned = stillOwned ?? throw new ArgumentNullException(nameof(stillOwned));
            this.exited = exited ?? throw new ArgumentNullException(nameof(exited));
            this.debuggerPresent = debuggerPresent ?? throw new ArgumentNullException(nameof(debuggerPresent));
        }

        internal static void ValidateIdentity(Identity value)
        {
            if (value == null || value.ProcessId <= 0 || !DateTime.TryParse(value.ProcessStartedUtc, out _) ||
                !Guid.TryParseExact(value.Nonce, "N", out _) || !Guid.TryParse(value.AssemblyMvid, out _) ||
                !LocalAbsolute(value.Root) || !Guid.TryParseExact(Path.GetFileName(value.Root), "N", out _) ||
                !LocalAbsolute(value.Executable) || !string.Equals(Path.GetFileName(value.Executable), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase) ||
                !LocalAbsolute(value.AssemblyPath) || !string.Equals(Path.GetFileName(value.AssemblyPath), "VBAi.dll", StringComparison.OrdinalIgnoreCase) ||
                !AllowedScenario(value.Scenario))
                throw new ArgumentException("Only an exact owned Excel procedure-value identity and GUID output are allowed.");
        }

        private static bool LocalAbsolute(string path) => path != null && path.Length > 2 && char.IsLetter(path[0]) &&
            path[1] == ':' && (path[2] == '\\' || path[2] == '/');

        internal static bool AllowedScenario(string name) => name == "NativeVariantArraysRoundTripWithBoundsAndOneInvocation" ||
            name == "NativeParamArrayCallsPreserveArityValuesAndSingleInvocation";

        /// <summary>Observes actual single emissions; a missing response never authorizes replay or cleanup.</summary>
        internal IDictionary<string, object> Command(object request, Func<IDictionary<string, object>> execute)
        {
            var serializer = new JavaScriptSerializer();
            var command = serializer.Deserialize<Dictionary<string, object>>(serializer.Serialize(request));
            string name = Convert.ToString(command["Command"]);
            emissions++;
            bool invocation = name == "run_procedure_values";
            if (invocation) invocationRequests++;
            try
            {
                Write("teardown-command-intent.json", Snapshot("BeforeSingleCommandDispatch"));
                var response = execute();
                if (response == null) { uncertain = true; return null; }
                if (Convert.ToBoolean(response["Ok"]) && (invocation || name == "procedure_values_status"))
                {
                    var data = response["Data"] as IDictionary<string, object>;
                    if (data == null || !data.ContainsKey("Query") || !data.ContainsKey("Pending"))
                        throw new InvalidOperationException("Procedure status is not terminal evidence.");
                    string query = Convert.ToString(data["Query"]);
                    if (data.TryGetValue("Uncertain", out object value) && Convert.ToBoolean(value)) uncertain = true;
                    if (string.IsNullOrWhiteSpace(query)) throw new InvalidOperationException("Procedure identity is missing.");
                    if (Convert.ToBoolean(data["Pending"])) pending.Add(query);
                    else pending.Remove(query);
                }
                Write("teardown-command-state.json", Snapshot("CommandObserved"));
                return response;
            }
            catch { uncertain = true; throw; }
        }

        internal void RequireTerminal()
        {
            if (uncertain || pending.Count != 0) throw new InvalidOperationException("Native work is pending or uncertain; retain exact host without Close/Quit.");
        }

        internal void BeforeCleanup()
        {
            RequireTerminal();
            if (armingStarted) throw new InvalidOperationException("Diagnostic arming was already attempted; cleanup must not be replayed.");
            if (!stillOwned()) throw new InvalidOperationException("Owned PID/start/image changed before cleanup.");
            armingStarted = true;
            Write("teardown.pending.json", identity);
            WaitMarker("teardown.armed.json", "Armed", false);
            if (!stillOwned()) throw new InvalidOperationException("Owned identity changed while arming; no cleanup emitted.");
            cleanupAttempts++;
            Write("teardown-cleanup-intent.json", Snapshot("SingleCleanupAuthorized"));
        }

        internal void AfterCleanup()
        {
            Write("teardown-cleanup-returned.json", Snapshot("CloseQuitReleaseReturned"));
            WaitMarker("teardown.detached.json", "Detached", true);
        }

        internal void ValidateMarker(string text, string phase)
        {
            if (text.Length > 16384) throw new InvalidOperationException("Diagnostic marker exceeds its bound.");
            var marker = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(text);
            if (Convert.ToInt32(marker["ProcessId"]) != identity.ProcessId ||
                Convert.ToString(marker["ProcessStartedUtc"]) != identity.ProcessStartedUtc ||
                Convert.ToString(marker["Nonce"]) != identity.Nonce || Convert.ToString(marker["Phase"]) != phase ||
                Convert.ToString(marker["AssemblyMvid"]) != identity.AssemblyMvid)
                throw new InvalidOperationException("Diagnostic marker PID/start/nonce/MVID/phase mismatch.");
        }

        private void WaitMarker(string file, string phase, bool permitTerminalExit)
        {
            DateTime started = Clock();
            while (Clock() - started < Deadline)
            {
                if (permitTerminalExit && exited()) { Write("teardown-host-terminal.json", Snapshot("OwnedHostExited")); return; }
                string path = Path.Combine(identity.Root, file);
                if (File.Exists(path))
                {
                    ValidateMarker(File.ReadAllText(path, Encoding.UTF8), phase);
                    if (debuggerPresent() != !permitTerminalExit) throw new InvalidOperationException("Debugger attachment/detachment was not independently verified.");
                    return;
                }
                WaitTick();
            }
            throw new TimeoutException("Diagnostic " + phase + " deadline exceeded; retain host, never replay cleanup.");
        }

        private object Snapshot(string state) => new { State = state, identity.ProcessId, identity.ProcessStartedUtc,
            identity.Nonce, identity.AssemblyMvid, CommandEmissions = emissions, InvocationRequests = invocationRequests,
            PendingQueries = pending.Count, DeliveryUncertain = uncertain, CleanupAttempts = cleanupAttempts,
            MutationReplayAllowed = false, ForceTermination = false, Utc = DateTime.UtcNow.ToString("o") };

        private void Write(string file, object value)
        { File.WriteAllText(Path.Combine(identity.Root, file), new JavaScriptSerializer().Serialize(value), new UTF8Encoding(false)); }
    }
}
