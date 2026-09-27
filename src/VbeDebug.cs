using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace CodexVBE
{
    internal sealed class VbeDebug
    {
        private readonly dynamic vbe;

        public VbeDebug(object vbe) { this.vbe = vbe; }

        public object State(string projectName)
        {
            dynamic project = GetProject(projectName);
            string selectedProject = null;
            string activeModule = null;
            object selection = null;
            try
            {
                dynamic pane = vbe.ActiveCodePane;
                if (pane != null)
                {
                    activeModule = (string)pane.CodeModule.Parent.Name;
                    selectedProject = (string)vbe.ActiveVBProject.Name;
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
                SelectedProject = selectedProject, ActiveModule = activeModule, Selection = selection };
        }

        public object ListCommands(string query)
        {
            var entries = EnumerateCommands();
            if (!string.IsNullOrWhiteSpace(query))
                entries = entries.Where(e => e.Path.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            return entries.Take(200).Select(e => new { e.Path, e.Caption, e.Id, e.Enabled }).ToArray();
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
                !string.Equals((string)state.SelectedProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
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
                case "reset":
                    if (beforeMode != 1) throw new InvalidOperationException("Reset requires break mode.");
                    id = 228; captions = new[] { "Réinitialiser", "Reset" }; break;
                case "clear_all_breakpoints":
                    if (beforeMode != 1 && beforeMode != 2)
                        throw new InvalidOperationException("Clear All Breakpoints requires break or design mode.");
                    id = 579; captions = new[] { "Effacer tous les points d'arrêt", "Effacer tous les points d’arrêt", "Clear All Breakpoints" };
                    break;
                default: throw new ArgumentException("Action must be reset or clear_all_breakpoints.");
            }
            var command = EnumerateCommands().FirstOrDefault(entry => entry.Id == id && entry.Enabled &&
                captions.Any(caption => (entry.Caption ?? "").Replace("&", "")
                    .IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0));
            if (command == null) throw new InvalidOperationException("The native VBE debug command is absent or disabled.");
            ((dynamic)command.Control).Execute();
            int afterMode = (int)project.Mode;
            bool verifiedReset = request.Action == "reset" && afterMode == 2;
            return new { request.Action, request.Project, Scope = request.Action == "clear_all_breakpoints" ? "Entire VBE" : "Active project",
                Executed = true, ControlId = id, Control = command.Path,
                ModeBefore = beforeMode, ModeAfter = afterMode,
                Verification = verifiedReset ? "Verified" : "Unverified",
                VerificationPending = request.Action == "reset" && !verifiedReset,
                VerificationLimit = request.Action == "clear_all_breakpoints"
                    ? "VBIDE has no breakpoint inventory; the command invocation alone does not prove every marker was cleared."
                    : verifiedReset ? null : "The VBE may apply reset after Execute returns.",
                NextRead = request.Action == "reset" ? "Call debug_state in a separate request to confirm design mode."
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
            dynamic pane = module.CodePane;
            pane.Show();
            pane.SetSelection(request.StartLine, 1, request.StartLine, 1);
            return new
            {
                request.Project,
                request.Module,
                Line = request.StartLine, Text = line, Mode = (int)project.Mode,
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
            if (request.Action == "toggle_breakpoint" &&
                (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("'", StringComparison.Ordinal)))
                throw new InvalidOperationException("A breakpoint requires an executable line.");

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
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Project is required.");
            var matches = new List<dynamic>();
            foreach (dynamic project in vbe.VBProjects)
                if (string.Equals((string)project.Name, name, StringComparison.OrdinalIgnoreCase))
                    matches.Add(project);
            if (matches.Count != 1) throw new InvalidOperationException("Project name is absent or ambiguous: " + name);
            return matches[0];
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
