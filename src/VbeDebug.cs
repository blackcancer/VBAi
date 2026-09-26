using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

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
            if (action == "step_into" || action == "step_over" || action == "continue")
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
