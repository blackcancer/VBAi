using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Human and compact machine reports derived from the same verified run.</summary>
    internal static class VbaTestReports
    {

        /// <summary>Formats the complete run and coverage result as readable text.</summary>
        /// <param name="run">Canonical run result; null produces a localized no-results message.</param>
        /// <returns>Human-readable summary and all result details.</returns>
        internal static string Human(VbaTestRun run)
        { return HumanReport(run, null, 0); }

        /// <summary>Formats a bounded page of test and coverage details while retaining run-wide totals.</summary>
        /// <param name="run">Canonical run result whose collections are paged.</param>
        /// <param name="offset">Zero-based item offset shared by result and coverage collections.</param>
        /// <param name="limit">Maximum page size; zero selects the default of 100.</param>
        /// <returns>Human-readable report page with total and next-offset metadata.</returns>
        internal static string HumanPage(VbaTestRun run, int offset = 0, int limit = 0)
        { return HumanReport(run, offset, PageLimit(offset, limit)); }

        /// <summary>Builds the shared readable report, optionally selecting one bounded page.</summary>
        /// <param name="run">Canonical run result.</param>
        /// <param name="offset">Zero-based start index, or null to include every item.</param>
        /// <param name="limit">Maximum number of items per collection when paged.</param>
        /// <returns>Summary, run-wide counts, and selected result/coverage details.</returns>
        private static string HumanReport(VbaTestRun run, int? offset, int limit)
        {
            if (run == null) return UiText.Get("No test run results.");
            var text = new StringBuilder();
            text.AppendLine(UiText.Get("VBA test results")).AppendLine(UiText.Get("Run: ") + run.Id)
                .AppendLine(UiText.Get("Project") + ": " + run.Project)
                .AppendLine(UiText.Get("Revision: ") + run.Revision);
            if (offset.HasValue)
            {
                text.AppendLine("Page: offset=" + offset.Value + "; limit=" + limit + "; tests total=" + run.Results.Count
                    + "; nextOffset=" + (NextOffset(offset.Value, limit, PageTotal(run))?.ToString(CultureInfo.InvariantCulture) ?? "null"));
                if (run.Coverage != null) text.AppendLine("Coverage detail totals: probes=" + run.Coverage.Hits.Count
                    + "; exclusions=" + run.Coverage.Exclusions.Count + "; diagnostics=" + run.Coverage.Diagnostics.Count);
            }
            foreach (VbaTestOutcome outcome in Enum.GetValues(typeof(VbaTestOutcome)))
            {
                int count = run.Results.Count(result => result.Outcome == outcome);
                if (count > 0) text.AppendLine(UiText.Get(outcome.ToString()) + ": " + count.ToString(CultureInfo.InvariantCulture));
            }
            int passed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed);
            int completed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed || result.Outcome == VbaTestOutcome.Failed || result.Outcome == VbaTestOutcome.Error);
            text.AppendLine(UiText.Get("Pass rate: ") + (completed == 0 ? "N/A" : (100d * passed / completed).ToString("0.##", CultureInfo.CurrentCulture) + "%") + " (" + passed + "/" + completed + ")")
                .AppendLine(CoverageText(run.Coverage));
            if (run.Coverage != null)
            {
                foreach (var hit in Page(run.Coverage.Hits, offset, limit)) text.AppendLine((hit.Entered ? "✓ " : "○ ") + hit.Probe.Module + "." + hit.Probe.Procedure + " [" + hit.Probe.Kind + "] @" + hit.Probe.OriginalLine);
                foreach (var exclusion in Page(run.Coverage.Exclusions, offset, limit)) text.AppendLine(UiText.Get("Coverage exclusions") + ": " + exclusion.Module + (string.IsNullOrEmpty(exclusion.Procedure) ? "" : "." + exclusion.Procedure) + ": " + exclusion.Reason);
                foreach (var diagnostic in Page(run.Coverage.Diagnostics, offset, limit)) text.AppendLine(diagnostic);
            }
            if (!string.IsNullOrEmpty(run.Error)) text.AppendLine(UiText.Get("Run error: ") + run.Error);
            if (run.OutcomeUnknown) text.AppendLine(UiText.Get("Execution outcome is uncertain. Inspect the host before any further run."));
            foreach (var result in Page(run.Results, offset, limit))
            {
                text.AppendLine().Append(result.Outcome == VbaTestOutcome.Passed ? "✓ " : result.Outcome == VbaTestOutcome.Failed || result.Outcome == VbaTestOutcome.Error ? "✗ " : "").Append(UiText.Get(result.Outcome.ToString())).Append(" | ").Append(result.Test.Module).Append('.').Append(result.Test.Procedure)
                    .Append(" | ").Append(result.Duration.TotalMilliseconds.ToString("0.###", CultureInfo.CurrentCulture)).AppendLine(" ms")
                    .AppendLine(UiText.Get("Phase: ") + result.Phase);
                if (!string.IsNullOrEmpty(result.Message)) text.AppendLine(result.Message);
                if (result.ErrorNumber != 0) text.AppendLine(UiText.Get("VBA error: ") + result.ErrorNumber);
            }
            return text.ToString();
        }

        /// <summary>Serializes the full canonical run as compact versioned JSON with the standard character bound.</summary>
        /// <param name="run">Canonical run result; null is represented as a versioned null run.</param>
        /// <returns>Compact JSON report; throws when it exceeds the fixed bound rather than truncating.</returns>
        internal static string Compact(VbaTestRun run) => Compact(run, 512 * 1024 * 1024);

        // A smaller bound can validate failure reporting without constructing an oversized report.
        /// <summary>Serializes compact JSON under a caller-supplied serialized-character limit.</summary>
        /// <param name="run">Canonical run result.</param>
        /// <param name="maximum">Maximum serialized character count; no fields or messages are truncated.</param>
        /// <returns>Complete JSON text within the limit.</returns>
        /// <exception cref="InvalidOperationException">The complete report exceeds the limit or cannot be serialized.</exception>
        internal static string Compact(VbaTestRun run, int maximum)
        {
            try { return SerializeCompact(run, null, 0, maximum); }
            catch (InvalidOperationException error)
            { throw new InvalidOperationException("The complete local VBA test JSON report could not be serialized within its " + maximum.ToString(CultureInfo.InvariantCulture) + " serialized-character limit. Use paged run status; no messages were truncated and the canonical run remains retained. " + error.Message, error); }
        }

        /// <summary>Serializes one compact JSON page while preserving global run and coverage summaries.</summary>
        /// <param name="run">Canonical run result.</param>
        /// <param name="offset">Zero-based offset into result and coverage detail arrays.</param>
        /// <param name="limit">Page size from 1 through 100, or zero for 100.</param>
        /// <returns>Complete JSON page within the fixed 10 MiB character limit.</returns>
        internal static string CompactPage(VbaTestRun run, int offset = 0, int limit = 0)
        {
            limit = PageLimit(offset, limit);
            return SerializeCompact(run, offset, limit, 10 * 1024 * 1024);
        }

        /// <summary>Validates page coordinates and resolves the default size.</summary>
        /// <param name="offset">Nonnegative zero-based collection index.</param>
        /// <param name="limit">Zero for the default 100, or a value from 1 through 100.</param>
        /// <returns>Validated effective page size.</returns>
        internal static int PageLimit(int offset, int limit)
        {
            if (offset < 0) throw new ArgumentException("Offset must be a nonnegative Int32.");
            if (limit < 0 || limit > 100) throw new ArgumentException("Limit must be between 1 and 100, or 0 for the default 100.");
            return limit == 0 ? 100 : limit;
        }

        /// <summary>Computes the next page start without overflowing past the collection end.</summary>
        /// <param name="offset">Current zero-based page start.</param>
        /// <param name="limit">Current page size.</param>
        /// <param name="total">Number of items in the collection.</param>
        /// <returns>Next zero-based offset, or null when this page reaches the end.</returns>
        internal static int? NextOffset(int offset, int limit, int total) =>
            offset >= total || total - offset <= limit ? (int?)null : offset + limit;

        /// <summary>Returns the largest result or coverage collection size used for shared paging.</summary>
        /// <param name="run">Run whose results and optional coverage arrays are measured.</param>
        /// <returns>Maximum item count across those collections.</returns>
        internal static int PageTotal(VbaTestRun run) => Math.Max(run.Results.Count, CoverageTotal(run.Coverage));

        /// <summary>Applies the requested slice, or returns the complete source for an unpaged report.</summary>
        /// <typeparam name="T">Collection item type.</typeparam>
        /// <param name="source">Items to select.</param>
        /// <param name="offset">Zero-based start, or null for the complete sequence.</param>
        /// <param name="limit">Maximum selected item count.</param>
        /// <returns>Deferred sequence for the selected page.</returns>
        private static IEnumerable<T> Page<T>(IEnumerable<T> source, int? offset, int limit) =>
            offset.HasValue ? source.Skip(offset.Value).Take(limit) : source;

        /// <summary>Builds compact JSON and rejects output beyond the supplied bound without truncation.</summary>
        /// <param name="run">Canonical run result, or null for the versioned null payload.</param>
        /// <param name="offset">Zero-based item offset, or null to include every detail.</param>
        /// <param name="limit">Maximum detail count when an offset is supplied.</param>
        /// <param name="maximum">Maximum serialized characters allowed.</param>
        /// <returns>Versioned compact JSON report.</returns>
        private static string SerializeCompact(VbaTestRun run, int? offset, int limit, int maximum)
        {
            if (run == null) return "{\"v\":1,\"run\":null}";
            int passed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed);
            int completed = run.Results.Count(result => result.Outcome == VbaTestOutcome.Passed || result.Outcome == VbaTestOutcome.Failed || result.Outcome == VbaTestOutcome.Error);
            var selected = offset.HasValue ? run.Results.Skip(offset.Value).Take(limit) : run.Results;
            var data = new Dictionary<string, object> {
                ["v"] = 1, ["run"] = run.Id, ["project"] = run.Project, ["revision"] = run.Revision,
                ["uncertain"] = run.OutcomeUnknown, ["error"] = run.Error,
                ["counts"] = run.Results.GroupBy(result => result.Outcome.ToString()).ToDictionary(group => group.Key, group => group.Count()),
                ["passRate"] = completed == 0 ? (double?)null : Math.Round(100d * passed / completed, 4),
                ["coverage"] = CompactCoverage(run.Coverage, offset, limit),
                ["tests"] = selected.Select(result => new { id = result.Test.Id, name = result.Test.Module + "." + result.Test.Procedure,
                    outcome = result.Outcome.ToString(), ms = Math.Round(result.Duration.TotalMilliseconds, 3), phase = result.Phase,
                    message = result.Message, error = result.ErrorNumber }).ToArray()
            };
            if (offset.HasValue)
            {
                int maximumTotal = Math.Max(run.Results.Count, CoverageTotal(run.Coverage));
                data["total"] = run.Results.Count; data["offset"] = offset.Value; data["limit"] = limit;
                data["nextOffset"] = NextOffset(offset.Value, limit, maximumTotal);
            }
            return new JavaScriptSerializer { MaxJsonLength = maximum }.Serialize(data);
        }

        /// <summary>Formats the coverage availability, completeness, denominator, hit count, and diagnostics.</summary>
        /// <param name="report">Coverage result, or null when coverage was not requested.</param>
        /// <returns>Readable coverage summary.</returns>
        internal static string CoverageText(VbaCoverageReport report)
        {
            if (report == null || !report.Available) return UiText.Get("VBA code coverage: unavailable");
            return UiText.Get("VBA procedure coverage") + ": " + (report.Percent.HasValue ? report.Percent.Value.ToString("0.##", CultureInfo.CurrentCulture) + "%" : "—") +
                " (" + report.Hit + "/" + report.Eligible + ")" + (report.Complete ? "" : " — " + UiText.Get("Partial measurement"));
        }

        /// <summary>Returns the largest coverage hit, exclusion, or diagnostic collection size.</summary>
        /// <param name="report">Coverage report, or null when no coverage details exist.</param>
        /// <returns>Maximum collection count, or zero.</returns>
        private static int CoverageTotal(VbaCoverageReport report) => report == null ? 0 :
            Math.Max(report.Hits.Count, Math.Max(report.Exclusions.Count, report.Diagnostics.Count));

        /// <summary>Returns global coverage metadata with bounded hit, exclusion, and diagnostic arrays.</summary>
        /// <param name="report">Coverage result, or null when coverage was not requested.</param>
        /// <param name="offset">Nonnegative zero-based index applied to each detail collection.</param>
        /// <param name="limit">Maximum returned items per detail array; zero selects 100.</param>
        /// <returns>Anonymous page object retaining overall totals and next-offset metadata.</returns>
        internal static object CoveragePage(VbaCoverageReport report, int offset = 0, int limit = 0)
        {
            limit = PageLimit(offset, limit);
            if (report == null) return null;
            // Preserve every existing public report field and its global summary; only collection payloads are paged.
            return new { report.Original, report.Revision, report.Metric, report.Available, report.Complete,
                report.DenominatorKnown, report.Eligible, report.Hit, report.Percent, report.StatementCoverageAvailable,
                Hits = report.Hits.Skip(offset).Take(limit).ToArray(), Exclusions = report.Exclusions.Skip(offset).Take(limit).ToArray(),
                Diagnostics = report.Diagnostics.Skip(offset).Take(limit).ToArray(),
                Total = CoverageTotal(report), HitTotal = report.Hits.Count, ExclusionTotal = report.Exclusions.Count,
                DiagnosticTotal = report.Diagnostics.Count, Offset = offset, Limit = limit,
                NextOffset = NextOffset(offset, limit, CoverageTotal(report)) };
        }

        /// <summary>Builds the compact JSON coverage object with procedure-entry probes and optional paging metadata.</summary>
        /// <param name="report">Coverage result, or null to describe coverage as not requested.</param>
        /// <param name="offset">Zero-based detail offset, or null to include all detail items.</param>
        /// <param name="limit">Maximum items in each selected detail collection.</param>
        /// <returns>JSON-serializable coverage object; statement and branch coverage remain unavailable.</returns>
        private static object CompactCoverage(VbaCoverageReport report, int? offset = null, int limit = 0)
        {
            if (report == null) return new { available = false, reason = "No coverage run was requested." };
            var hits = offset.HasValue ? report.Hits.Skip(offset.Value).Take(limit) : report.Hits;
            var exclusions = offset.HasValue ? report.Exclusions.Skip(offset.Value).Take(limit) : report.Exclusions;
            var diagnostics = offset.HasValue ? report.Diagnostics.Skip(offset.Value).Take(limit) : report.Diagnostics;
            var data = new Dictionary<string, object> {
                ["available"] = report.Available, ["metric"] = report.Metric, ["complete"] = report.Complete,
                ["revision"] = report.Revision, ["original"] = report.Original, ["denominatorKnown"] = report.DenominatorKnown,
                ["eligible"] = report.Eligible, ["hit"] = report.Hit, ["percent"] = report.Percent,
                ["statementCoverageAvailable"] = false, ["branchCoverageAvailable"] = false,
                ["probes"] = hits.Select(hit => new { id = hit.Probe.Id, module = hit.Probe.Module, procedure = hit.Probe.Procedure,
                    kind = hit.Probe.Kind, line = hit.Probe.OriginalLine, column = hit.Probe.OriginalColumn, entered = hit.Entered }).ToArray(),
                ["exclusions"] = exclusions.Select(item => new { module = item.Module, procedure = item.Procedure, kind = item.Kind,
                    line = item.OriginalLine, reason = item.Reason, intentional = item.Intentional }).ToArray(),
                ["diagnostics"] = diagnostics.ToArray() };
            if (offset.HasValue)
            {
                data["total"] = CoverageTotal(report); data["probeTotal"] = report.Hits.Count;
                data["exclusionTotal"] = report.Exclusions.Count; data["diagnosticTotal"] = report.Diagnostics.Count;
                data["offset"] = offset.Value; data["limit"] = limit;
                data["nextOffset"] = NextOffset(offset.Value, limit, CoverageTotal(report));
            }
            return data;
        }
    }
}
