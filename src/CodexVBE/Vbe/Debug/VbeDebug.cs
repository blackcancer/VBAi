using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        private readonly dynamic vbe;

        public VbeDebug(object vbe) { this.vbe = vbe; }

        public object State(string projectName)
        {
            dynamic project = GetProject(projectName);
            string selectedProject = null;
            string selectedProjectPath = null;
            string activeModule = null;
            object selection = null;
            try
            {
                dynamic pane = vbe.ActiveCodePane;
                if (pane != null)
                {
                    activeModule = (string)pane.CodeModule.Parent.Name;
                    selectedProject = (string)vbe.ActiveVBProject.Name;
                    try { selectedProjectPath = (string)vbe.ActiveVBProject.FileName; } catch { }
                    int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
                    pane.GetSelection(ref startLine, ref startColumn, ref endLine, ref endColumn);
                    selection = new { StartLine = startLine, StartColumn = startColumn,
                        EndLine = endLine, EndColumn = endColumn };
                }
            }
            catch (Exception ex)
            {
                selection = new { Error = ex.Message };
            }
            return new { Project = (string)project.Name, Mode = (int)project.Mode,
                SelectedProject = selectedProject, SelectedProjectPath = selectedProjectPath,
                ActiveModule = activeModule, Selection = selection };
        }

        public object ListCommands(string query, int offset, int limit)
        {
            if (offset < 0 || limit < 0)
                throw new ArgumentOutOfRangeException("Offset and Limit must be non-negative.");
            int pageSize = limit == 0 ? 200 : Math.Min(limit, 200);
            var entries = EnumerateCommands();
            if (!string.IsNullOrWhiteSpace(query))
                entries = entries.Where(e => e.Path.Replace("&", "").IndexOf(query.Replace("&", ""), StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            return entries.Skip(offset).Take(pageSize).Select(e => new { e.Path, e.Caption, e.Id, e.Enabled }).ToArray();
        }

        public object OpenObjectBrowser(VbeEditorWindows windows)
        {
            if (windows == null) throw new ArgumentNullException(nameof(windows));
            object before = windows.Windows();
            bool visibleBefore = HasVisibleObjectBrowser(before);
            // Office's native Object Browser command is 473. The caption check
            // also prevents executing an unrelated control with a reused ID.
            var candidates = EnumerateCommands().Where(e => e.Id == 473 && e.Enabled &&
                IsObjectBrowserCaption(e.Caption)).ToList();
            if (candidates.Count == 0)
                throw new InvalidOperationException("The VBE Object Browser command (Id 473) is absent or disabled.");
            CommandEntry selected = candidates[0];
            ((dynamic)selected.Control).Execute();
            object after = windows.Windows();
            bool visibleAfter = HasVisibleObjectBrowser(after);
            return new { Executed = true, Control = selected.Path, ControlId = selected.Id,
                AlreadyVisible = visibleBefore,
                Verification = visibleAfter ? "Visible" : "Pending",
                VerificationPending = !visibleAfter,
                NextRead = visibleAfter ? null : "Call vbe_windows in a separate request; the VBE may open the browser after Execute returns.",
                VerificationLimit = "Only the native Object Browser window is observed; its classes and members are not read structurally.",
                WindowsBefore = before, WindowsAfter = after };
        }

        public object OpenDebugPane(string paneName, VbeEditorWindows windows)
        {
            int id;
            string[] captions;
            switch ((paneName ?? "").Trim().ToLowerInvariant())
            {
                case "locals": id = 2555; captions = new[] { "Variables locales", "Locals Window" }; break;
                case "watches": id = 2556; captions = new[] { "Espions", "Watch Window" }; break;
                case "immediate": id = 2554; captions = new[] { "Exécution", "Immediate Window" }; break;
                default: throw new ArgumentException("Pane must be locals, watches or immediate.");
            }
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == id && entry.Enabled &&
                captions.Any(caption => (entry.Caption ?? "").Replace("&", "")
                    .IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The requested VBE debug pane command is absent or disabled.");
            ((dynamic)command.Control).Execute();
            return new { Pane = paneName, Executed = true, ControlId = id, Control = command.Path,
                VerificationPending = true, NextRead = "Call vbe_windows or debug_windows in a separate request to confirm the pane is visible." };
        }

        public object QueueAddWatchDialog(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Module) || string.IsNullOrWhiteSpace(request.Expression))
                throw new ArgumentException("Project, Module and Expression are required.");
            if (request.Expression.Length > 1024)
                throw new ArgumentException("Watch expression exceeds 1024 characters.");
            if (!string.IsNullOrWhiteSpace(request.WatchType) && request.WatchType != "expression" &&
                request.WatchType != "break_when_true" && request.WatchType != "break_when_changed")
                throw new ArgumentException("WatchType must be expression, break_when_true or break_when_changed.");
            dynamic state = State(request.Project);
            if ((int)state.Mode != 1 || request.ExpectedMode != 1 ||
                !string.Equals((string)(System.IO.Path.IsPathRooted(request.Project) ? state.SelectedProjectPath : state.SelectedProject),
                    request.Project, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals((string)state.ActiveModule, request.Module, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The requested project/module must be active in break mode before adding a watch.");
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 1820 && entry.Enabled &&
                ((entry.Caption ?? "").Replace("&", "").IndexOf("espion", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (entry.Caption ?? "").IndexOf("Add Watch", StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native Add Watch command is unavailable.");
            SynchronizationContext context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ => {
                try { ((dynamic)command.Control).Execute(); }
                catch (Exception ex) { LoadLog.Write("Add Watch dialog failed: " + ex.Message); }
            }, null);
            return new { Scheduled = true, ControlId = command.Id, request.Project, request.Module,
                NextRead = "Complete the native Add Watch dialog after this command returns." };
        }

        public object QueueEditWatchDialog(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Expression) || string.IsNullOrWhiteSpace(request.Context) ||
                string.IsNullOrWhiteSpace(request.NewExpression) || request.NewExpression.Length > 1024)
                throw new ArgumentException("Project, Expression, Context and NewExpression (at most 1024 characters) are required.");
            if (!string.IsNullOrWhiteSpace(request.WatchType) && request.WatchType != "expression" &&
                request.WatchType != "break_when_true" && request.WatchType != "break_when_changed")
                throw new ArgumentException("WatchType must be expression, break_when_true or break_when_changed.");
            dynamic state = State(request.Project);
            if ((int)state.Mode != request.ExpectedMode)
                throw new InvalidOperationException("Project mode changed before editing the watch.");
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 940 && entry.Enabled &&
                ((entry.Caption ?? "").Replace("&", "").IndexOf("Modifier un espion", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (entry.Caption ?? "").IndexOf("Edit Watch", StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native Edit Watch command is unavailable.");
            SynchronizationContext context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ => {
                try { ((dynamic)command.Control).Execute(); }
                catch (Exception ex) { LoadLog.Write("Edit Watch dialog failed: " + ex.Message); }
            }, null);
            return new { Scheduled = true, ControlId = command.Id, request.Expression, request.Context };
        }

        public object QueueQuickWatchDialog(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Expression) ||
                request.Expression.Length > 1024 || request.StartColumn < 1 ||
                request.EndColumn <= request.StartColumn)
                throw new ArgumentException("A selected single-line Expression with columns is required.");
            dynamic state = State(request.Project);
            if ((int)state.Mode != 1 || request.ExpectedMode != 1)
                throw new InvalidOperationException("Quick Watch requires the project in break mode.");
            SelectCode(request);
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 229 && entry.Enabled &&
                ((entry.Caption ?? "").Replace("&", "").IndexOf("Espion express", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (entry.Caption ?? "").IndexOf("Quick Watch", StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native Quick Watch command is unavailable.");
            SynchronizationContext context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ => {
                try { ((dynamic)command.Control).Execute(); }
                catch (Exception ex) { LoadLog.Write("Quick Watch dialog failed: " + ex.Message); }
            }, null);
            return new { Scheduled = true, ControlId = command.Id, request.Project, request.Module,
                request.Expression, request.StartLine, request.StartColumn, request.EndColumn };
        }

        public object QueueDebugOptionsDialog()
        {
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 522 && entry.Enabled &&
                string.Equals((entry.Caption ?? "").Replace("&", "").TrimEnd('.'),
                    "Options", StringComparison.OrdinalIgnoreCase));
            if (command == null) throw new InvalidOperationException("The native VBE Tools > Options command is unavailable.");
            SynchronizationContext context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ => {
                try { ((dynamic)command.Control).Execute(); }
                catch (Exception ex) { LoadLog.Write("VBE Options dialog failed: " + ex.Message); }
            }, null);
            return new { Scheduled = true, ControlId = command.Id, Control = command.Path };
        }

        public object RemoveSelectedWatch(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Expression) || string.IsNullOrWhiteSpace(request.Context))
                throw new ArgumentException("Project, Expression and Context are required.");
            dynamic state = State(request.Project);
            if ((int)state.Mode != request.ExpectedMode)
                throw new InvalidOperationException("Project mode changed before removing the watch.");
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 1083 && entry.Enabled &&
                ((entry.Caption ?? "").Replace("&", "").IndexOf("Supprimer un espion", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 (entry.Caption ?? "").IndexOf("Delete Watch", StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native Delete Watch command is unavailable.");
            ((dynamic)command.Control).Execute();
            return new { Executed = true, ControlId = command.Id, request.Expression, request.Context,
                VerificationPending = true, NextRead = "Read debug_windows in a separate request to verify the selected watch is absent." };
        }

        public object ExecuteGlobalDebugCommand(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project))
                throw new ArgumentException("Project is required.");
            dynamic project = GetProject(request.Project);
            int beforeMode = (int)project.Mode;
            if (beforeMode != request.ExpectedMode)
                throw new InvalidOperationException("Project mode changed before the native debug command.");
            int id;
            string[] captions;
            switch (request.Action)
            {
                case "break":
                    if (beforeMode != 0) throw new InvalidOperationException("Break requires run mode.");
                    id = 189; captions = new[] { "Arrêt", "Break" }; break;
                case "reset":
                    if (beforeMode != 1) throw new InvalidOperationException("Reset requires break mode.");
                    id = 228; captions = new[] { "Réinitialiser", "Reset" }; break;
                case "clear_all_breakpoints":
                    if (beforeMode != 1 && beforeMode != 2)
                        throw new InvalidOperationException("Clear All Breakpoints requires break or design mode.");
                    id = 579; captions = new[] { "Effacer tous les points d'arrêt", "Effacer tous les points d’arrêt", "Clear All Breakpoints" };
                    break;
                case "show_next_statement":
                    if (beforeMode != 1) throw new InvalidOperationException("Show Next Statement requires break mode.");
                    dynamic activeProject = vbe.ActiveVBProject;
                    if (activeProject == null ||
                        !SameComObject((object)activeProject, (object)project))
                        throw new InvalidOperationException("Show Next Statement requires the requested project to be active in the VBE.");
                    dynamic activePane = vbe.ActiveCodePane;
                    if (activePane == null)
                        throw new InvalidOperationException("Show Next Statement requires an active code pane in the requested project.");
                    dynamic activeComponent = activePane.CodeModule.Parent;
                    dynamic requestedComponent;
                    try { requestedComponent = activeProject.VBComponents.Item((string)activeComponent.Name); }
                    catch (Exception) { throw new InvalidOperationException("The active code pane belongs to another project."); }
                    if (!SameComObject(activeComponent, requestedComponent))
                        throw new InvalidOperationException("The active code pane belongs to another project.");
                    id = 1813; captions = new[] { "Afficher l'instruction suivante", "Afficher l’instruction suivante", "Show Next Statement" };
                    break;
                default: throw new ArgumentException("Action must be break, reset, clear_all_breakpoints or show_next_statement.");
            }
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == id && entry.Enabled &&
                captions.Any(caption => (entry.Caption ?? "").Replace("&", "")
                    .IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native VBE debug command is absent or disabled.");
            ((dynamic)command.Control).Execute();
            int afterMode = (int)project.Mode;
            bool verifiedBreak = request.Action == "break" && afterMode == 1;
            bool verifiedReset = request.Action == "reset" && afterMode == 2;
            return new { request.Action, request.Project, Scope = request.Action == "clear_all_breakpoints" || request.Action == "break" ? "Entire VBE" : "Active project",
                Executed = true, ControlId = id, Control = command.Path,
                ModeBefore = beforeMode, ModeAfter = afterMode,
                Verification = verifiedBreak || verifiedReset ? "Verified" : "Unverified",
                VerificationPending = (request.Action == "break" && !verifiedBreak) ||
                    (request.Action == "reset" && !verifiedReset) ||
                    request.Action == "show_next_statement",
                VerificationLimit = request.Action == "clear_all_breakpoints"
                    ? "VBIDE has no breakpoint inventory; the command invocation alone does not prove every marker was cleared."
                    : request.Action == "show_next_statement" ? "Selection after navigation is not an independent execution-pointer inventory." :
                        verifiedBreak || verifiedReset ? null : "The VBE may apply the command after Execute returns.",
                NextRead = request.Action == "break" ? "Call debug_state in a separate request to confirm break mode."
                    : request.Action == "reset" ? "Call debug_state in a separate request to confirm design mode."
                    : request.Action == "show_next_statement" ? "Call debug_state in a separate request to read the resulting code selection."
                    : "Run a disposable procedure or inspect the native editor to verify breakpoint behavior." };
        }

        public object CompileProject(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project))
                throw new ArgumentException("Project is required.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2 || request.ExpectedMode != 2)
                throw new InvalidOperationException("Compilation requires the selected project in design mode.");
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 578 && entry.Enabled &&
                (((entry.Caption ?? "").Replace("&", "").IndexOf("Compiler ", StringComparison.OrdinalIgnoreCase) >= 0) ||
                 ((entry.Caption ?? "").Replace("&", "").IndexOf("Compile ", StringComparison.OrdinalIgnoreCase) >= 0)));
            if (command == null)
                throw new InvalidOperationException("The native Compile command is absent or disabled.");
            string caption = (command.Caption ?? "").Replace("&", "");
            if (caption.IndexOf((string)project.Name, StringComparison.OrdinalIgnoreCase) < 0)
                throw new InvalidOperationException("The native Compile command targets a different project: " + caption);
            ((dynamic)command.Control).Execute();
            return new { Executed = true, Project = request.Project, ControlId = command.Id, Control = command.Path };
        }

        private static bool IsObjectBrowserCaption(string caption)
        {
            string name = (caption ?? "").Replace("&", "").Trim();
            return name.IndexOf("Explorateur d'objets", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Explorateur d’objets", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Object Browser", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasVisibleObjectBrowser(object windowState)
        {
            foreach (dynamic item in ((dynamic)windowState).Windows)
            {
                var properties = (IDictionary<string, object>)item.Properties;
                object type, visible;
                if (properties.TryGetValue("Type", out type) && Convert.ToInt32(type) == 2 &&
                    properties.TryGetValue("Visible", out visible) && Convert.ToBoolean(visible))
                    return true;
            }
            return false;
        }

        public object SelectCode(Request request)
        {
            dynamic project = GetProject(request.Project);
            dynamic module = GetModule(project, request.Module);
            string line = ValidateLocation(request, module);
            if ((request.StartColumn == 0) != (request.EndColumn == 0))
                throw new ArgumentException("StartColumn and EndColumn must be supplied together.");
            int startColumn = request.StartColumn == 0 ? 1 : request.StartColumn;
            int endColumn = request.EndColumn == 0 ? 1 : request.EndColumn;
            if (startColumn < 1 || endColumn < startColumn || endColumn > line.Length + 1)
                throw new ArgumentOutOfRangeException("The selected columns are outside the current code line.");
            string selectedText = line.Substring(startColumn - 1, endColumn - startColumn);
            if (!string.IsNullOrEmpty(request.Expression) &&
                !string.Equals(selectedText, request.Expression, StringComparison.Ordinal))
                throw new InvalidOperationException("The selected source text does not match Expression.");
            dynamic pane = module.CodePane;
            pane.Show();
            pane.SetSelection(request.StartLine, startColumn, request.StartLine, endColumn);
            int actualStartLine = 0, actualStartColumn = 0, actualEndLine = 0, actualEndColumn = 0;
            pane.GetSelection(ref actualStartLine, ref actualStartColumn, ref actualEndLine, ref actualEndColumn);
            if (actualStartLine != request.StartLine || actualEndLine != request.StartLine ||
                actualStartColumn != startColumn || actualEndColumn != endColumn)
                throw new InvalidOperationException("The native code pane did not retain the requested selection.");
            return new
            {
                request.Project,
                request.Module,
                Line = request.StartLine, Text = line, SelectedText = selectedText,
                StartColumn = startColumn, EndColumn = endColumn, Mode = (int)project.Mode,
                State = State(request.Project) };
        }

        public object InvokeCommand(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Action))
                throw new ArgumentException("Action is required.");
            if (request.ControlId <= 0 || string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("ControlId and ControlCaption are required.");
            dynamic project = GetProject(request.Project);
            dynamic module = GetModule(project, request.Module);
            string line = ValidateLocation(request, module);
            int mode = (int)project.Mode;
            if (mode != request.ExpectedMode)
                throw new InvalidOperationException("Project mode changed before the debug command.");
            if (!IsAllowed(request.Action, request.ControlCaption, mode))
                throw new InvalidOperationException("The control caption is not allowed for the requested debug action.");
            if ((request.Action == "toggle_breakpoint" || request.Action == "set_next_statement") &&
                (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("'", StringComparison.Ordinal)))
                throw new InvalidOperationException("The requested debug action requires an executable line.");
            if (request.Action == "set_next_statement")
                ValidateSetNextStatementProcedure(request, module);

            dynamic pane = module.CodePane;
            pane.Show();
            pane.SetSelection(request.StartLine, 1, request.StartLine, 1);
            dynamic activePane = vbe.ActiveCodePane;
            if (activePane == null || !SameComObject(pane, activePane))
                throw new InvalidOperationException("The requested code pane is not active in the VBE.");

            var matches = EnumerateCommands().Where(e => e.Id == request.ControlId &&
                string.Equals(e.Caption, request.ControlCaption, StringComparison.Ordinal)).ToList();
            if (matches.Count == 0)
                throw new InvalidOperationException("The requested VBE command was not found.");
            var selected = matches.FirstOrDefault(e => e.Enabled && IsAllowed(request.Action, e.Caption, mode)) ?? throw new InvalidOperationException("The requested VBE command is disabled.");
            dynamic control = selected.Control;
            object before = State(request.Project);
            control.Execute();
            object after = null;
            string afterError = null;
            try { after = State(request.Project); }
            catch (Exception ex) { afterError = ex.Message; }
            string evidence = DebugEffect(request.Action, before, after);
            bool pending = evidence == null && request.Action != "toggle_breakpoint";
            return new
            {
                request.Action,
                Control = selected.Path,
                request.Project,
                request.Module,
                Line = request.StartLine, Text = line,
                ModeBefore = mode, Executed = true,
                Verification = evidence == null ? "Unverified" : "Verified",
                VerificationPending = pending,
                Evidence = evidence,
                NextRead = pending ? "Call debug_state in a separate request after the VBE processes the command; inspect native debug windows for the visible effect." : null,
                VerificationLimit = request.Action == "toggle_breakpoint"
                    ? "VBIDE exposes no breakpoint inventory through this command; toggle effect was not verified."
                    : pending ? "The VBE may process this command asynchronously; immediate state did not yet prove an effect." : null,
                StateBefore = before, StateAfter = after, StateAfterError = afterError };
        }

        public object SelectCodeRange(Request request)
        {
            if (request == null || request.StartColumn < 1 || request.EndColumn < 1 ||
                request.EndLine < request.StartLine ||
                (request.EndLine == request.StartLine && request.EndColumn < request.StartColumn))
                throw new ArgumentException("A forward one-based code range is required.");
            dynamic project = GetProject(request.Project);
            dynamic module = GetModule(project, request.Module);
            string startText = ValidateLocation(request, module);
            int lineCount = (int)module.CountOfLines;
            if (request.EndLine > lineCount)
                throw new ArgumentOutOfRangeException("EndLine", "The range ends outside the code module.");
            string endText = (string)module.Lines[request.EndLine, 1];
            if (request.StartColumn > startText.Length + 1 || request.EndColumn > endText.Length + 1)
                throw new ArgumentOutOfRangeException("The selected columns are outside their code lines.");
            dynamic pane = module.CodePane;
            pane.Show();
            pane.SetSelection(request.StartLine, request.StartColumn, request.EndLine, request.EndColumn);
            int actualStartLine = 0, actualStartColumn = 0, actualEndLine = 0, actualEndColumn = 0;
            pane.GetSelection(ref actualStartLine, ref actualStartColumn, ref actualEndLine, ref actualEndColumn);
            if (actualStartLine != request.StartLine || actualStartColumn != request.StartColumn ||
                actualEndLine != request.EndLine || actualEndColumn != request.EndColumn)
                throw new InvalidOperationException("The native code pane did not retain the requested range.");
            return new { request.Project, request.Module, request.StartLine, request.StartColumn,
                request.EndLine, request.EndColumn, Sha256 = request.ExpectedSha256,
                Verified = true, Mode = (int)project.Mode };
        }

        public object QueueSignatureDialog(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || request.ExpectedMode != 2)
                throw new ArgumentException("Project and ExpectedMode=2 are required.");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The selected project must be in design mode.");
            dynamic active = vbe.ActiveVBProject;
            if (active == null || !SameComObject((object)active, (object)project))
                throw new InvalidOperationException("Select the exact project in VBE before opening its signature dialog.");
            // VBE reuses Id 746 for both Digital Signature and Remove Component.
            // When a removable component is selected, Execute can dispatch the
            // Remove action even on the Tools > Signature control. Show a host
            // document component and refuse the action if the collision remains.
            if (HasEnabledRemove746())
            {
                dynamic document = null;
                foreach (dynamic component in project.VBComponents)
                    if ((int)component.Type == 100) { document = component; break; }
                if (document == null)
                    throw new InvalidOperationException("VBE Id 746 also targets Remove Component; no safe document component is available.");
                document.CodeModule.CodePane.Show();
                if (HasEnabledRemove746())
                    throw new InvalidOperationException("VBE Id 746 still targets Remove Component; signature was not started.");
            }
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 746 && entry.Enabled &&
                (entry.Path.IndexOf("Outils", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 entry.Path.IndexOf("Tools", StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null || (command.Caption ?? "").Replace("&", "")
                .IndexOf("Signature", StringComparison.OrdinalIgnoreCase) < 0)
                throw new InvalidOperationException("The native VBE Digital Signature command is unavailable.");
            SynchronizationContext context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ => {
                try { ((dynamic)command.Control).Execute(); }
                catch (Exception ex) { LoadLog.Write("VBE Digital Signature dialog failed: " + ex.Message); }
            }, null);
            return new { Scheduled = true, Project = request.Project, ControlId = command.Id,
                Control = command.Path };
        }

        private bool HasEnabledRemove746()
        {
            return EnumerateCommands().Any(entry => entry.Id == 746 && entry.Enabled &&
                ((entry.Caption ?? "").Replace("&", "").TrimStart().StartsWith("Supprimer ", StringComparison.OrdinalIgnoreCase) ||
                 (entry.Caption ?? "").Replace("&", "").TrimStart().StartsWith("Remove ", StringComparison.OrdinalIgnoreCase)));
        }

        public object RunSub(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Procedure) ||
                !Regex.IsMatch(request.Procedure, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("An exact VBA Sub name is required.");
            if (request.ExpectedMode != 2)
                throw new ArgumentException("ExpectedMode must be 2 (design mode).");
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2)
                throw new InvalidOperationException("The selected project is no longer in design mode.");
            dynamic component = null;
            foreach (dynamic item in project.VBComponents)
                if (string.Equals((string)item.Name, request.Module, StringComparison.OrdinalIgnoreCase))
                { component = item; break; }
            if (component == null || (int)component.Type != 1)
                throw new InvalidOperationException("Run Sub requires a standard module in the selected project.");
            dynamic module = component.CodeModule;
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256))
                throw new ArgumentException("ExpectedSha256 from read_module is required.");
            int count = (int)module.CountOfLines;
            string code = count == 0 ? string.Empty : (string)module.Lines[1, count];
            if (!string.Equals(Hash(code), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int body = (int)module.ProcBodyLine[request.Procedure, 0];
            string declaration = (string)module.Lines[body, 1];
            string pattern = @"^\s*(?:(?:Public|Private|Friend|Static)\s+)*Sub\s+" +
                Regex.Escape(request.Procedure) + @"\s*(?:\(\s*\))?\s*(?:'.*)?$";
            if (!Regex.IsMatch(declaration, pattern, RegexOptions.IgnoreCase))
                throw new InvalidOperationException("Only a parameterless standard-module Sub can be run by name.");
            dynamic pane = module.CodePane;
            pane.Show();
            pane.SetSelection(body, 1, body, 1);
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == 186 &&
                entry.Enabled && IsAllowed("run", entry.Caption, 2));
            if (command == null)
                throw new InvalidOperationException("The native Run Sub command is absent or disabled.");
            request.StartLine = body;
            request.Action = "run";
            request.ControlId = command.Id;
            request.ControlCaption = command.Caption;
            return InvokeCommand(request);
        }

        private void ValidateSetNextStatementProcedure(Request request, dynamic module)
        {
            dynamic activePane = vbe.ActiveCodePane;
            if (activePane == null || !SameComObject(activePane, (object)module.CodePane))
                throw new InvalidOperationException("Show Next Statement in the target code pane before setting a new execution line.");
            int currentLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
            activePane.GetSelection(ref currentLine, ref startColumn, ref endLine, ref endColumn);
            int currentKind = 0, targetKind = 0;
            string currentProcedure = (string)module.ProcOfLine[currentLine, ref currentKind];
            string targetProcedure = (string)module.ProcOfLine[request.StartLine, ref targetKind];
            if (string.IsNullOrWhiteSpace(currentProcedure) ||
                string.IsNullOrWhiteSpace(targetProcedure) ||
                currentKind != targetKind ||
                !string.Equals(currentProcedure, targetProcedure, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Set Next Statement requires a target in the currently selected procedure. Use Show Next Statement first.");
        }

        private static string DebugEffect(string action, object before, object after)
        {
            if (after == null || action == "toggle_breakpoint") return null;
            dynamic initial = before;
            dynamic current = after;
            int oldMode = (int)initial.Mode;
            int newMode = (int)current.Mode;
            if (oldMode != newMode)
                return "Project mode changed from " + oldMode + " to " + newMode + ".";
            if (action == "step_into" || action == "step_over" || action == "step_out" ||
                action == "run_to_cursor" || action == "continue")
            {
                try
                {
                    int oldLine = (int)initial.Selection.StartLine;
                    int newLine = (int)current.Selection.StartLine;
                    string oldModule = (string)initial.ActiveModule;
                    string newModule = (string)current.ActiveModule;
                    if (oldLine != newLine ||
                        !string.Equals(oldModule, newModule, StringComparison.OrdinalIgnoreCase))
                        return "Active code location changed from " + oldModule + ":" + oldLine +
                            " to " + newModule + ":" + newLine + ".";
                }
                catch { /* Selection may be unavailable in a native debug window. */ }
            }
            return null;
        }

        private static bool IsAllowed(string action, string caption, int mode)
        {
            string label = caption.Replace("&", "").Trim();
            switch (action)
            {
                case "toggle_breakpoint":
                    return (mode == 1 || mode == 2) &&
                        (label.IndexOf("breakpoint", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("point d'arr", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("point d’arrêt", StringComparison.OrdinalIgnoreCase) >= 0);
                case "run":
                    return mode == 2 &&
                        (label.StartsWith("Run Sub", StringComparison.OrdinalIgnoreCase) ||
                         label.StartsWith("Exécuter Sub", StringComparison.OrdinalIgnoreCase) ||
                         label.StartsWith("Exécuter la macro", StringComparison.OrdinalIgnoreCase));
                case "continue":
                    return mode == 1 &&
                        (label.StartsWith("Continue", StringComparison.OrdinalIgnoreCase) ||
                         label.StartsWith("Continuer", StringComparison.OrdinalIgnoreCase));
                case "step_into":
                    return mode == 1 &&
                        (label.IndexOf("Step Into", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("pas à pas détaillé", StringComparison.OrdinalIgnoreCase) >= 0);
                case "step_over":
                    return mode == 1 &&
                        (label.IndexOf("Step Over", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("pas à pas principal", StringComparison.OrdinalIgnoreCase) >= 0);
                case "step_out":
                    return mode == 1 &&
                        (label.IndexOf("Step Out", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("pas à pas sortant", StringComparison.OrdinalIgnoreCase) >= 0);
                case "run_to_cursor":
                    return mode == 1 &&
                        (label.IndexOf("Run To Cursor", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("Exécuter jusqu'au curseur", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("Exécuter jusqu’au curseur", StringComparison.OrdinalIgnoreCase) >= 0);
                case "set_next_statement":
                    return mode == 1 &&
                        (label.IndexOf("Set Next Statement", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("Définir l'instruction suivante", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         label.IndexOf("Définir l’instruction suivante", StringComparison.OrdinalIgnoreCase) >= 0);
                default:
                    return false;
            }
        }

        private static string ValidateLocation(Request request, dynamic module)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256))
                throw new ArgumentException("ExpectedSha256 is required for debug actions.");
            int count = (int)module.CountOfLines;
            if (request.StartLine < 1 || request.StartLine > count)
                throw new ArgumentOutOfRangeException("StartLine");
            string code = (string)module.Lines[1, count];
            if (!string.Equals(Hash(code), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            return (string)module.Lines[request.StartLine, 1];
        }

        private static string Hash(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code)))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static bool SameComObject(object first, object second)
        {
            IntPtr firstUnknown = IntPtr.Zero, secondUnknown = IntPtr.Zero;
            try
            {
                firstUnknown = Marshal.GetIUnknownForObject(first);
                secondUnknown = Marshal.GetIUnknownForObject(second);
                return firstUnknown == secondUnknown;
            }
            finally
            {
                if (firstUnknown != IntPtr.Zero) Marshal.Release(firstUnknown);
                if (secondUnknown != IntPtr.Zero) Marshal.Release(secondUnknown);
            }
        }

        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        private static dynamic GetModule(dynamic project, string moduleName)
        {
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module is required.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                    return component.CodeModule;
            throw new InvalidOperationException("Module not found: " + moduleName);
        }

        private sealed class CommandEntry
        {
            public object Control;
            public string Path;
            public string Caption;
            public int Id;
            public bool Enabled;
        }

        private List<CommandEntry> EnumerateCommands()
        {
            var entries = new List<CommandEntry>();
            foreach (dynamic bar in vbe.CommandBars)
            {
                string name;
                try { name = (string)bar.Name; }
                catch { continue; }
                AddControls(entries, bar.Controls, name, 0);
            }
            return entries;
        }

        private static void AddControls(List<CommandEntry> entries, dynamic controls, string path, int depth)
        {
            if (depth > 4 || entries.Count >= 2000) return;
            foreach (dynamic control in controls)
            {
                if (entries.Count >= 2000) return;
                try
                {
                    string caption = (string)control.Caption;
                    string currentPath = path + " > " + caption;
                    entries.Add(new CommandEntry { Control = control, Caption = caption,
                        Id = (int)control.Id, Enabled = (bool)control.Enabled, Path = currentPath });
                    try { AddControls(entries, control.Controls, currentPath, depth + 1); }
                    catch { /* Buttons do not have child controls. */ }
                }
                catch { /* Skip unavailable controls. */ }
            }
        }
    }
}
