using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Reads a bounded page of explicitly declared scalar identifiers from one verified paused procedure.</summary>
    internal sealed partial class VbeDebug
    {

        /// <summary>Reads the native Locals context used to verify the paused procedure.</summary>
        internal Func<string> LocalContextReader = VbeDebugWindows.ReadLocalsContext;

        /// <summary>Rejects an existing Quick Watch dialog before opening another scalar inspection.</summary>
        internal Action EnsureScalarDialogAbsent = VbeDebugWindows.EnsureNoQuickWatchDialog;

        /// <summary>Reads the result of a native Quick Watch dialog for the requested scalar expression.</summary>
        internal Func<Request, object> ReadScalarDialog = VbeDebugWindows.ReadScalarQuickWatch;

        /// <summary>Optional asynchronous evaluator override used by focused fault tests.</summary>
        internal Func<Request, Task<object>> LocalScalarEvaluator;

        /// <summary>Explicit, bounded inspection of declared scalar identifiers in one verified paused context.</summary>
        /// <param name="request">Paused-mode project/module/procedure identity, module hash, and candidate-page offset/limit.</param>
        /// <returns>Task reporting declared scalar candidates and sequential native reads; it is not an atomic Locals inventory.</returns>
        internal Task<object> InspectLocalScalarsAsync(Request request)
        {
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.CoreEntered);
            return VbeUiTask.Run(() => InspectLocalScalarsCoreAsync(request));
        }

        /// <summary>Inspects local scalars core async for vbe debug.</summary>
        /// <param name="request">Validated pause context and bounded page request for local scalar candidates.</param>
        /// <returns>One sequential result row per candidate until completion or a context/inspection error.</returns>
        private async Task<object> InspectLocalScalarsCoreAsync(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.Module) ||
                string.IsNullOrWhiteSpace(request.Procedure) || string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                request.ExpectedMode != 1 || request.ProcKind != 0 || request.Offset < 0 || request.Limit < 0 || request.Limit > 16)
                throw new ArgumentException("Project, Module, Procedure, ExpectedSha256 and ExpectedMode=1 are required; Offset must be non-negative and Limit at most 16. Only Sub/Function is supported.");
            if (VbeDebugInspection.IsActive) throw new InvalidOperationException("A native debugger inspection is already active.");
            using (new VbeDebugInspection())
            {
                dynamic project = GetProject(request.Project);
                dynamic module = GetModule(project, request.Module);
                string projectName = (string)project.Name, moduleName = (string)module.Parent.Name;
                string expectedContext = projectName + "." + moduleName + "." + request.Procedure;
                int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                Action validate = () => {
                    VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ContextValidation);
                    if (System.Threading.Thread.CurrentThread.ManagedThreadId != ownerThread)
                        throw new InvalidOperationException("Local inspection left the owning VBE thread.");
                    object activeProject = vbe.ActiveVBProject;
                    if ((int)project.Mode != 1 || activeProject == null || !SameComObject(activeProject, (object)project))
                        throw new InvalidOperationException("The requested project must remain active and paused.");
                    int lines = (int)module.CountOfLines;
                    if (lines < 1 || !string.Equals(Hash((string)module.Lines[1, lines]), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The module changed since it was read.");
                    if (!string.Equals(LocalContextReader(), expectedContext, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The native Locals context differs from the requested procedure.");
                    VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ContextValidated);
                };
                validate();
                EnsureScalarDialogAbsent();
                int start = (int)module.ProcStartLine[request.Procedure, 0];
                int body = (int)module.ProcBodyLine[request.Procedure, 0];
                int count = (int)module.ProcCountLines[request.Procedure, 0];
                int last = start + count - 1, totalLines = (int)module.CountOfLines;
                if (count < 1 || start < 1 || body < start || body > last || last > totalLines)
                    throw new InvalidOperationException("VBIDE returned an invalid procedure range.");
                string source = (string)module.Lines[1, totalLines];
                var candidates = VbaLocalScalarCandidates.Read(source, request.Procedure, body, last);
                int limit = request.Limit == 0 ? 8 : request.Limit;
                var page = candidates.Skip(request.Offset).Take(limit).ToArray();
                dynamic pane = module.CodePane;
                object previousWindow = vbe.ActiveWindow;
                int oldLine = 0, oldColumn = 0, oldEndLine = 0, oldEndColumn = 0;
                pane.GetSelection(ref oldLine, ref oldColumn, ref oldEndLine, ref oldEndColumn);
                Request lastSelection = null;
                bool selectionRestored = true, focusRestored = true;
                string interrupted = null;
                var rows = new List<object>();
                try
                {
                    foreach (var candidate in page)
                    {
                        string value = null, error = null, status = candidate.Eligible ? "Read" : "Skipped";
                        if (candidate.Eligible)
                        {
                            try
                            {
                                validate();
                                var expression = new Request { Project = request.Project, Module = moduleName, Procedure = request.Procedure,
                                    ExpectedMode = 1, ExpectedSha256 = request.ExpectedSha256, Expression = candidate.Expression,
                                    StartLine = candidate.Line, StartColumn = candidate.Column, EndColumn = candidate.Column + candidate.Expression.Length };
                                lastSelection = expression;
                                dynamic observed = await (LocalScalarEvaluator == null ? EvaluateLocalScalarAsync(expression, validate, projectName) : LocalScalarEvaluator(expression));
                                validate();
                                if ((string)observed.Expression != expression.Expression ||
                                    !string.Equals((string)observed.Context, expectedContext, StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("Quick Watch readback differs from the requested identifier or context.");
                                value = (string)observed.Value;
                            }
                            catch (Exception ex) { status = "Error"; error = interrupted = ex.Message; }
                        }
                        rows.Add(new { candidate.Name, candidate.Expression, candidate.Kind, candidate.TypeName, candidate.Line, candidate.Column,
                            Status = status, Value = value, Error = error, SkipReason = candidate.Reason });
                        if (interrupted != null) break;
                    }
                }
                finally
                {
                    if (lastSelection != null)
                    {
                        selectionRestored = focusRestored = false;
                        try
                        {
                            // Restore only the selection still owned by this operation, never a concurrent navigation.
                            validate();
                            int a = 0, b = 0, c = 0, d = 0;
                            pane.GetSelection(ref a, ref b, ref c, ref d);
                            object activePane = vbe.ActiveCodePane;
                            object activeWindow = vbe.ActiveWindow;
                            if (activePane != null && SameComObject(activePane, (object)pane) &&
                                a == lastSelection.StartLine && c == a && b == lastSelection.StartColumn && d == lastSelection.EndColumn)
                            {
                                pane.SetSelection(oldLine, oldColumn, oldEndLine, oldEndColumn);
                                pane.GetSelection(ref a, ref b, ref c, ref d);
                                selectionRestored = a == oldLine && b == oldColumn && c == oldEndLine && d == oldEndColumn;
                                if (previousWindow != null && activeWindow != null && SameComObject(activeWindow, (object)pane.Window))
                                {
                                    ((dynamic)previousWindow).SetFocus();
                                    object restored = vbe.ActiveWindow;
                                    focusRestored = restored != null && SameComObject(restored, previousWindow);
                                }
                            }
                        }
                        catch { /* A changed context is reported without navigating into it. */ }
                    }
                }
                if (interrupted == null) validate();
                VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.CoreTerminal);
                return new { Project = projectName, Module = moduleName, request.Procedure, Context = expectedContext,
                    Sha256 = request.ExpectedSha256, Coverage = "DeclaredScalarCandidatesOnly", RuntimeInventoryComplete = false,
                    TotalCandidates = candidates.Length, EligibleCandidates = candidates.Count(c => c.Eligible), request.Offset,
                    NextOffset = interrupted == null && request.Offset + page.Length < candidates.Length ? (int?)(request.Offset + page.Length) : null,
                    Aborted = interrupted != null, Error = interrupted, Items = rows,
                    SelectionRestored = selectionRestored, FocusRestored = focusRestored,
                    Limit = "Sequential native values, not an atomic Locals snapshot. Undeclared variables, function results and module members are not enumerated. Arrays, Variant, objects, UDTs, ambiguous and conditional declarations are not evaluated. No expression is synthesized or inserted." };
            }
        }

        /// <summary>Evaluates one approved scalar using Quick Watch and verifies the dialog's expression and procedure context.</summary>
        /// <param name="request">Exact identifier expression, source range, project/module, and paused-mode revision.</param>
        /// <param name="validate">Rechecks active paused project, module hash, Locals procedure, and owner thread around native interaction.</param>
        /// <param name="projectName">Expected project name included in the native Quick Watch context check.</param>
        /// <returns>Observed scalar value and dialog evidence; mismatch or context change throws.</returns>
        private async Task<object> EvaluateLocalScalarAsync(Request request, Action validate, string projectName)
        {
            EnsureScalarDialogAbsent();
            validate();
            dynamic pane = GetModule(GetProject(request.Project), request.Module).CodePane;
            pane.Show();
            vbe.ActiveCodePane = pane;
            pane.Window.SetFocus();
            var command = FindAvailableCommand(229, entry => {
                string caption = (entry.Caption ?? "").Replace("&", "");
                return caption.IndexOf("Espion express", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    caption.IndexOf("Quick Watch", StringComparison.OrdinalIgnoreCase) >= 0;
            });
            if (command == null) throw new InvalidOperationException("Native Quick Watch is unavailable.");
            validate();
            SelectCode(request);
            object active = vbe.ActiveCodePane;
            if (active == null || !SameComObject(active, (object)pane))
                throw new InvalidOperationException("The selected native code pane is no longer active.");
            // The observer runs off the UI thread so it can read and close the modal dialog.
            // Native captions use the resolved project name even when the request identifies it by path.
            var dialogRequest = new Request { Project = projectName, Module = request.Module,
                Procedure = request.Procedure, Expression = request.Expression };
            Task<object> reading = Task.Run(() => ReadScalarDialog(dialogRequest));
            Exception commandError = null;
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.Command229Before);
            try { ((dynamic)command.Control).Execute(); }
            catch (Exception ex) { commandError = ex; }
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.Command229Returned, commandError);
            object result = await reading;
            if (commandError != null) throw new InvalidOperationException("Quick Watch failed; its command was not retried.", commandError);
            return result;
        }
    }
}
