using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Preserves each synthetic scalar-inspection request before delivery and all independent failure phases.</summary>
    internal sealed class ExcelScalarQualificationEvidence
    {
        private readonly string report;
        private readonly int processId;
        private readonly string fixtureRoot;
        private readonly List<IDictionary<string, object>> commands = new List<IDictionary<string, object>>();
        private Exception primaryFailure, cleanupFailure;
        internal string Shutdown { get; set; } = "NotVerified; owned host may be retained";
        internal string PhaseTracePath { get; set; }
        internal object PhaseEvidence { get; set; }

        internal ExcelScalarQualificationEvidence(string report, int processId, string fixtureRoot)
        { this.report = report; this.processId = processId; this.fixtureRoot = fixtureRoot; }

        /// <summary>Persists intent before the only send, then persists its response, exception and elapsed time.</summary>
        internal IDictionary<string, object> Send(object request, Func<IDictionary<string, object>> execute,
            Action<IDictionary<string, object>> validate)
        {
            var record = new Dictionary<string, object> {
                ["Sequence"] = commands.Count + 1, ["Request"] = request,
                ["StartedUtc"] = DateTime.UtcNow.ToString("o"), ["State"] = "Prepared; delivery outcome has not been recorded"
            };
            commands.Add(record);
            Save(); // An evidence failure prevents emission; it never causes another send.
            var timing = Stopwatch.StartNew();
            IDictionary<string, object> response = null;
            Exception failure = null;
            try { response = execute(); validate(response); record["State"] = "ResponseVerified"; }
            catch (Exception error) { failure = error; record["State"] = "Failed; delivery may be uncertain"; }
            record["Response"] = response;
            record["Error"] = failure?.ToString();
            record["ElapsedMilliseconds"] = timing.ElapsedMilliseconds;
            record["FinishedUtc"] = DateTime.UtcNow.ToString("o");
            try { Save(); }
            catch (Exception evidence)
            {
                if (failure != null) throw new AggregateException("The scalar command and its durable evidence write both failed; no request was replayed.", failure, evidence);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            return response;
        }

        /// <summary>Runs independent cleanup and final reporting while preserving the original scenario error.</summary>
        internal void Run(Action scenario, Action cleanup, Action attach)
        {
            var failures = new List<Exception>();
            try { scenario(); }
            catch (Exception error) { primaryFailure = error; failures.Add(error); }
            try { cleanup(); }
            catch (Exception error) { cleanupFailure = error; failures.Add(error); }
            try { Save(); }
            catch (Exception error) { failures.Add(error); }
            try { if (System.IO.File.Exists(report)) attach(); }
            catch (Exception error) { failures.Add(error); }
            if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
            if (failures.Count > 1) throw new AggregateException("Excel scalar qualification failed in independent scenario, cleanup or evidence phases; all errors are retained.", failures);
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(report));
            var payload = new {
                HostProcessId = processId, FixtureRoot = fixtureRoot,
                AssemblyMvid = typeof(VbeSession).Module.ModuleVersionId,
                ObservedUtc = DateTime.UtcNow.ToString("o"), Scope = "Owned disposable scalar test; synthetic requests only",
                Evidence = commands, PhaseTracePath, PhaseEvidence, Shutdown, PrimaryError = primaryFailure?.ToString(), CleanupError = cleanupFailure?.ToString()
            };
            System.IO.File.WriteAllText(report, new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 }.Serialize(payload), new UTF8Encoding(false));
        }
    }
}
