using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns the vbe test explorer service state and operations.</summary>
    internal sealed partial class VbeTestExplorerService
    {

        /// <summary>Maintains the create coverage clone state for vbe test explorer service.</summary>
        internal Func<object, string, string, VbaTestCoverageClone> CreateCoverageClone = VbaTestCoverageClone.CreateOwned;

        /// <summary>Maintains the coverage root state for vbe test explorer service.</summary>
        internal Func<string> CoverageRoot = () => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "CoverageRuns");

        /// <summary>Maintains the coverage service state for vbe test explorer service.</summary>
        private VbeTestExplorerService coverageService;

        /// <summary>Maintains the compile coverage project state for vbe test explorer service.</summary>
        internal Action<object> CompileCoverageProject;

        /// <summary>Handles coverage unavailable reason for vbe test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <returns>Text produced by the operation for coverage unavailable reason on vbe test explorer service.</returns>
        public string CoverageUnavailableReason(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (catalog == null) return "Select and refresh a VBA project.";
            if (!IsExecutionHost()) return "Measured coverage currently requires an owned Excel, Word or PowerPoint host.";
            if (string.IsNullOrEmpty(catalog.Project.HostPath) || !Path.IsPathRooted(catalog.Project.HostPath))
                return "Save the macro-enabled host document before collecting coverage.";
            var extensions = Host is VbaTestWordValuesHost ? new[] { ".docm", ".dotm", ".doc", ".dot" }
                : Host is VbaTestPowerPointValuesHost
                ? new[] { ".pptm", ".ppsm", ".potm", ".ppt", ".pps", ".pot" }
                : new[] { ".xlsm", ".xlsb", ".xls" };
            if (!extensions.Contains(Path.GetExtension(catalog.Project.HostPath), StringComparer.OrdinalIgnoreCase))
                return "This document format does not support an owned coverage copy.";
            var plan = VbaCoverageInstrumentation.Create(catalog.Project);
            if (!plan.CanInstrument) return string.Join(Environment.NewLine, plan.Diagnostics);
            try {
                Validate(catalog);
                var target = Host.ResolveTarget(ResolveLive(catalog.Project.Id), catalog.Project.HostPath);
                try {
                    if (target is VbaTestWordValuesHost.OwnedTarget word && !((bool)((dynamic)word.Document).Saved))
                        return "Save all Word document changes before collecting coverage from its saved-file copy.";
                }
                finally { ReleaseReturnedTarget(target, false); }
            }
            catch (Exception error) { return error.Message; }
            return null;
        }

        /// <summary>Handles preview coverage for vbe test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        /// <param name="limit">int that supplies the limit for this operation.</param>
        /// <returns>object produced by the operation for preview coverage on vbe test explorer service.</returns>
        private object PreviewCoverage(VbaTestCatalog catalog, int offset = 0, int limit = 0)
        {
            limit = VbaTestReports.PageLimit(offset, limit);
            var plan = VbaCoverageInstrumentation.Create(catalog.Project);
            string unavailable = CoverageUnavailableReason(catalog);
            return CoveragePreviewPage(catalog.Project.Selector, catalog.Project.Revision, plan, unavailable, offset, limit);
        }

        /// <summary>Handles coverage preview page for vbe test explorer service.</summary>
        /// <param name="project">Text that supplies the project value. Use the format required by the calling operation.</param>
        /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
        /// <param name="plan">vba coverage plan that supplies the plan for this operation.</param>
        /// <param name="unavailable">Text that supplies the unavailable value. Use the format required by the calling operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        /// <param name="limit">int that supplies the limit for this operation.</param>
        /// <returns>object produced by the operation for coverage preview page on vbe test explorer service.</returns>
        internal static object CoveragePreviewPage(string project, string revision, VbaCoveragePlan plan, string unavailable, int offset = 0, int limit = 0)
        {
            limit = VbaTestReports.PageLimit(offset, limit);
            return new { Project = project, ExpectedProjectVersion = revision,
                Available = false, Reason = unavailable ?? "No measured coverage run was requested; pass rate is a separate metric.",
                Supported = plan.CanInstrument && unavailable == null, ExecutionUnavailableReason = unavailable,
                Metric = "Procedure", plan.EligibleProcedureCount, plan.DenominatorKnown,
                Probes = plan.Probes.Skip(offset).Take(limit).ToArray(),
                Exclusions = plan.Exclusions.Skip(offset).Take(limit).ToArray(), Diagnostics = plan.Diagnostics.Skip(offset).Take(limit).ToArray(),
                Total = Math.Max(plan.Probes.Count, Math.Max(plan.Exclusions.Count, plan.Diagnostics.Count)),
                ProbeTotal = plan.Probes.Count, ExclusionTotal = plan.Exclusions.Count, DiagnosticTotal = plan.Diagnostics.Count,
                Offset = offset, Limit = limit,
                NextOffset = VbaTestReports.NextOffset(offset, limit, Math.Max(plan.Probes.Count, Math.Max(plan.Exclusions.Count, plan.Diagnostics.Count))),
                StatementCoverageAvailable = false,
                Method = "Explicit coverage run on a separate instrumented host-document copy; original source is never instrumented." };
        }

        /// <summary>Runs coverage async for vbe test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="tests">i read only list&lt;vba test descriptor&gt; that supplies the tests for this operation.</param>
        /// <param name="onResult">action&lt;vba test result&gt; that supplies the on result for this operation.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>task&lt;vba test run&gt; produced by the operation for run coverage async on vbe test explorer service.</returns>
        public Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation)
        { return BeginRun(catalog, tests, onResult, cancellation, null, true); }

        /// <summary>Executes coverage async for vbe test explorer service.</summary>
        /// <param name="original">vba test catalog that supplies the original for this operation.</param>
        /// <param name="selected">i read only list&lt;vba test descriptor&gt; that supplies the selected for this operation.</param>
        /// <param name="progress">action&lt;vba test result&gt; that supplies the progress for this operation.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <param name="executionGuard">action that supplies the execution guard for this operation.</param>
        /// <returns>task&lt;vba test run&gt; produced by the operation for execute coverage async on vbe test explorer service.</returns>
        private async Task<VbaTestRun> ExecuteCoverageAsync(VbaTestCatalog original, IReadOnlyList<VbaTestDescriptor> selected,
            Action<VbaTestResult> progress, CancellationToken cancellation, Action executionGuard)
        {
            var plan = VbaCoverageInstrumentation.Create(original.Project);
            var run = new VbaTestRun { Project = original.Project.Name, Revision = original.Project.Revision };
            VbaTestCoverageClone clone = null;
            bool ownedClone = false;
            string folder = Path.Combine(CoverageRoot(), Guid.NewGuid().ToString("N"));
            Action guard = () => { RequireOwner(); executionGuard?.Invoke(); Validate(original); };
            try
            {
                cancellation.ThrowIfCancellationRequested();
                guard();
                clone = CreateCoverageClone(ResolveLive(original.Project.Id), original.Project.HostPath, folder);
                guard();
                if (clone?.Project == null || string.IsNullOrEmpty(clone.Path)) throw new InvalidOperationException("The coverage clone is unavailable.");
                if (ReferenceEquals(clone.Project, ResolveLive(original.Project.Id)) ||
                    VbeDebug.NativeProcedureValuesHost.SameComIdentity(clone.Project, ResolveLive(original.Project.Id)))
                    throw new InvalidOperationException("Coverage must not instrument the original project.");
                if (CanonicalPath(clone.Path) == CanonicalPath(original.Project.HostPath))
                    throw new InvalidOperationException("Coverage must use a separate workbook path.");
                if (!Path.GetFullPath(clone.Path).StartsWith(Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The coverage copy is outside its owned output folder.");
                var before = Snapshot("coverage", clone.Project);
                if (CanonicalPath(before.HostPath) != CanonicalPath(clone.Path)) throw new InvalidOperationException("The coverage project path does not match the owned copy.");
                ownedClone = true;
                if (before.Modules.Length != original.Project.Modules.Length || before.ReferencesHash != original.Project.ReferencesHash)
                    throw new InvalidOperationException("The coverage clone does not match all original modules and references.");
                foreach (var module in original.Project.Modules)
                {
                    var copied = before.Modules.SingleOrDefault(item => item.Name.Equals(module.Name, StringComparison.OrdinalIgnoreCase));
                    if (copied == null || copied.Hash != module.Hash || copied.ComponentType != module.ComponentType)
                        throw new InvalidOperationException("The coverage clone does not match the original source: " + module.Name);
                }
                Directory.CreateDirectory(folder);
                File.WriteAllText(Path.Combine(folder, "coverage-plan.json"), new System.Web.Script.Serialization.JavaScriptSerializer
                    { MaxJsonLength = 32 * 1024 * 1024 }.Serialize(plan), new UTF8Encoding(false));
                foreach (var module in plan.Modules)
                {
                    guard(); cancellation.ThrowIfCancellationRequested();
                    InstrumentCloneModule(clone.Project, module, () => { guard(); cancellation.ThrowIfCancellationRequested(); });
                }
                guard(); cancellation.ThrowIfCancellationRequested();
                WriteCloneModule(clone.Project, VbaCoverageInstrumentation.ModuleName, plan.RuntimeSource, true);
                coverageService = new VbeTestExplorerService((object)vbe, dispatcher) { Host = Host, IsExecutionHost = IsExecutionHost,
                    BackupRoot = () => Path.Combine(folder, "support-backups") };
                string copyId = coverageService.FindIdentity(clone.Project);
                var copiedCatalog = coverageService.Discover(copyId);
                guard(); cancellation.ThrowIfCancellationRequested();
                coverageService.ApplySupport(copiedCatalog, VbaTestRuntimeSource.Generate(copiedCatalog));
                copiedCatalog = coverageService.Discover(copyId);
                var copyTests = selected.Select(test => copiedCatalog.Tests.Single(item =>
                    item.Module.Equals(test.Module, StringComparison.OrdinalIgnoreCase) && item.Procedure.Equals(test.Procedure, StringComparison.OrdinalIgnoreCase))).ToArray();
                guard(); cancellation.ThrowIfCancellationRequested();
                if (CompileCoverageProject == null) throw new InvalidOperationException("A verified coverage compilation boundary is required.");
                var compileProject = CompileCoverageProject;
                currentCoverageCompileGuard = () => {
                    guard(); coverageService.Validate(copiedCatalog); cancellation.ThrowIfCancellationRequested();
                };
                try
                {
                    compileProject(clone.Project);
                    if (compileProject == defaultCompileCoverageDelegate)
                        await AwaitOwner(VerifyCoverageCompilationAsync(clone.Project, currentCoverageCompileGuard));
                }
                finally { currentCoverageCompileGuard = null; }
                guard(); coverageService.Validate(copiedCatalog);
                var reset = InvokeCoverageFunction(coverageService, copiedCatalog, VbaCoverageInstrumentation.ResetProcedure, guard);
                if (!(reset is bool initialized) || !initialized) throw new InvalidOperationException("The coverage runtime did not verify a successful reset; no tests were dispatched.");
                Action<VbaTestResult> translate = result => {
                    var test = selected.Single(item => item.Module.Equals(result.Test.Module, StringComparison.OrdinalIgnoreCase) && item.Procedure.Equals(result.Test.Procedure, StringComparison.OrdinalIgnoreCase));
                    var translated = new VbaTestResult { Test = test, Outcome = result.Outcome, Message = result.Message,
                        Phase = result.Phase, ErrorNumber = result.ErrorNumber, Duration = result.Duration };
                    run.Results.Add(translated); progress?.Invoke(translated);
                };
                var measured = await AwaitOwner(coverageService.BeginRun(copiedCatalog, copyTests, translate, cancellation, guard));
                run.Error = measured.Error; run.OutcomeUnknown = measured.OutcomeUnknown;
                if (measured.OutcomeUnknown) { outcomeUnknown = true; run.Coverage = VbaCoverageInstrumentation.Unavailable(plan, "A native test outcome is uncertain; further coverage dispatch was refused."); }
                else
                {
                    var hits = InvokeCoverageFunction(coverageService, copiedCatalog, VbaCoverageInstrumentation.SnapshotProcedure, guard);
                    run.Coverage = VbaCoverageInstrumentation.Read(plan, hits, !cancellation.IsCancellationRequested && measured.Error == null);
                }
            }
            catch (Exception error)
            {
                run.Error = error.Message;
                run.OutcomeUnknown |= error is VbaTestInvocationException invocation && invocation.Uncertain;
                outcomeUnknown |= run.OutcomeUnknown;
                run.Coverage = VbaCoverageInstrumentation.Unavailable(plan, error.Message);
            }
            finally
            {
                coverageService?.Dispose(); coverageService = null;
                try
                {
                    if (ownedClone && !run.OutcomeUnknown) clone.Dispose();
                    else if (clone != null && run.Coverage != null)
                    {
                        if (Host is VbaTestWordValuesHost) VbaTestWordValuesHost.RetainAcquired(clone);
                        run.Coverage.Diagnostics.Add("The copy was retained open because ownership or native completion could not be verified: " + clone.Path);
                    }
                }
                catch (Exception error) { run.Error = (run.Error == null ? "" : run.Error + Environment.NewLine) + "The coverage copy could not be closed: " + error.Message;
                    run.OutcomeUnknown |= error is VbaTestInvocationException invocation && invocation.Uncertain;
                    outcomeUnknown |= run.OutcomeUnknown;
                    if (run.Coverage != null) { run.Coverage.Complete = false; run.Coverage.Percent = null; run.Coverage.Diagnostics.Add("The owned coverage copy remains open. Inspect " + clone.Path); } }
                foreach (var test in selected.Where(test => run.Results.All(result => result.Test.Id != test.Id)))
                {
                    var result = new VbaTestResult { Test = test, Outcome = cancellation.IsCancellationRequested ? VbaTestOutcome.Cancelled : VbaTestOutcome.Blocked,
                        Phase = "Preparation", Message = run.Error ?? "The coverage run stopped before this test was dispatched." };
                    run.Results.Add(result); progress?.Invoke(result);
                }
            }
            return run;
        }

        /// <summary>Invokes coverage function for vbe test explorer service.</summary>
        /// <param name="copy">vbe test explorer service that supplies the copy for this operation.</param>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="procedure">Text that supplies the procedure value. Use the format required by the calling operation.</param>
        /// <param name="guard">action that supplies the guard for this operation.</param>
        /// <returns>object produced by the operation for invoke coverage function on vbe test explorer service.</returns>
        private object InvokeCoverageFunction(VbeTestExplorerService copy, VbaTestCatalog catalog, string procedure, Action guard)
        {
            guard(); copy.Validate(catalog);
            object target = copy.Host.ResolveTarget(copy.ResolveLive(catalog.Project.Id), catalog.Project.HostPath);
            bool uncertain = false;
            try
            {
                guard(); copy.Validate(catalog);
                try {
                    var returned = copy.Host.Invoke(target, VbaCoverageInstrumentation.ModuleName, procedure, new object[0]);
                    RequireOwner(); guard(); copy.Validate(catalog);
                    return returned;
                }
                catch (Exception error) {
                    uncertain = true;
                    throw new VbaTestInvocationException("Coverage runtime completion is uncertain: " + error.Message, true, error);
                }
            }
            finally { ReleaseReturnedTarget(target, uncertain); }
        }

        /// <summary>Determines whether it can onical path for vbe test explorer service.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>Text produced by the operation for canonical path on vbe test explorer service.</returns>
        private static string CanonicalPath(string path) => Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar).ToUpperInvariant();

        /// <summary>Writes clone module for vbe test explorer service.</summary>
        /// <param name="project">dynamic that supplies the project for this operation.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
        /// <param name="create">Indicates whether create is enabled.</param>
        private static void WriteCloneModule(dynamic project, string name, string source, bool create)
        {
            dynamic component = null;
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, name, StringComparison.OrdinalIgnoreCase)) { component = candidate; break; }
            if (create && component != null) throw new InvalidOperationException("The coverage runtime module name is occupied.");
            if (component == null)
            {
                if (!create) throw new InvalidOperationException("A copied source module is missing: " + name);
                component = project.VBComponents.Add(1); component.Name = name;
            }
            dynamic code = component.CodeModule;
            int count = (int)code.CountOfLines;
            if (count > 0) code.DeleteLines(1, count);
            code.AddFromString(source);
            if (Canonical((string)code.Lines[1, (int)code.CountOfLines]) != Canonical(source))
                throw new InvalidOperationException("The instrumented copy source did not match its plan: " + name);
        }

        /// <summary>Handles instrument clone module for vbe test explorer service.</summary>
        /// <param name="project">dynamic that supplies the project for this operation.</param>
        /// <param name="module">vba coverage module that supplies the module for this operation.</param>
        /// <param name="guard">action that supplies the guard for this operation.</param>
        private static void InstrumentCloneModule(dynamic project, VbaCoverageModule module, Action guard)
        {
            dynamic component = null;
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, module.Name, StringComparison.OrdinalIgnoreCase)) { component = candidate; break; }
            if (component == null) throw new InvalidOperationException("An owned copied module is missing: " + module.Name);
            dynamic code = component.CodeModule;
            // Preserve member metadata by retaining every original declaration and only inserting probes.
            foreach (var edit in module.Edits.OrderByDescending(item => item.OriginalLine).ThenByDescending(item => item.OriginalColumn))
            {
                guard();
                if (edit.IsWholeLine) code.InsertLines(edit.OriginalLine, edit.Text.TrimEnd('\r', '\n'));
                else
                {
                    string line = (string)code.Lines[edit.OriginalLine, 1];
                    code.ReplaceLine(edit.OriginalLine, line.Insert(edit.OriginalColumn - 1, edit.Text));
                }
            }
            string observed = (int)code.CountOfLines == 0 ? "" : (string)code.Lines[1, (int)code.CountOfLines];
            if (Canonical(observed) != Canonical(module.InstrumentedSource))
                throw new InvalidOperationException("The coverage probes do not match their reviewed original-source mapping: " + module.Name);
        }
    }
}
