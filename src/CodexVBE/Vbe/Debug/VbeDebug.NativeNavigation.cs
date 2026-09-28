using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        private sealed class NavigationOperation
        {
            public string Id, State, Error, Action;
            public object Before, After;
            public bool CommandCompleted, NavigationObserved;
            public DateTime ReadbackDeadlineUtc;
        }
        private readonly List<NavigationOperation> nativeNavigations = new List<NavigationOperation>();

        public object NativeNavigation(Request request)
        {
            if (request.Action == "status")
            {
                var found = nativeNavigations.SingleOrDefault(x => x.Id == request.Query);
                if (found == null) throw new InvalidOperationException("Unknown navigation operation in this session.");
                ObserveNavigation(found);
                return NavigationResult(found);
            }
            if (request.Action != "definition" && request.Action != "last_position") throw new ArgumentException("Use definition, last_position or status.");
            foreach (var prior in nativeNavigations) ObserveNavigation(prior);
            if (nativeNavigations.Any(x => x.State == "Queued" || x.State == "Running" || x.State == "Observing"))
                throw new InvalidOperationException("A native navigation is still pending. Inspect its status and any native dialog.");
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            dynamic project = GetProject(request.Project);
            int mode = (int)project.Mode;
            if ((mode != 1 && mode != 2) || request.ExpectedMode != mode) throw new InvalidOperationException("The expected design or break mode is required.");
            dynamic module = GetModule(project, request.Module);
            ValidateLocation(request, module);
            int id = request.Action == "definition" ? 939 : 1822;
            if (string.IsNullOrWhiteSpace(request.ControlCaption) || !EnumerateCommands().Any(x => x.Id == id && x.Enabled && x.Caption == request.ControlCaption))
                throw new InvalidOperationException("The exact enabled native command is required from list_commands.");
            if (request.Action == "definition")
            {
                if (string.IsNullOrWhiteSpace(request.Expression) || request.StartColumn < 1 || request.EndColumn <= request.StartColumn)
                    throw new ArgumentException("Definition requires an exact nonempty source selection and Expression.");
                SelectCode(request);
            }
            else
            {
                dynamic active = vbe.ActiveCodePane;
                if (active == null || !SameComObject((object)active.CodeModule, (object)module)) throw new InvalidOperationException("The requested module must already be active.");
                int line = 0, column = 0, end = 0, endColumn = 0;
                active.GetSelection(ref line, ref column, ref end, ref endColumn);
                if (line != request.StartLine || end != line || column != request.StartColumn || endColumn != request.EndColumn)
                    throw new InvalidOperationException("The active selection changed. Read code_panes again.");
            }
            var operation = new NavigationOperation { Id = Guid.NewGuid().ToString("N"), Action = request.Action, State = "Queued", Before = NativeNavigationPosition() };
            nativeNavigations.Add(operation);
            if (nativeNavigations.Count > 20) nativeNavigations.RemoveAt(0);
            // Capture primitive values: the caller may reuse its Request after scheduling.
            string projectSelector = request.Project, moduleName = request.Module, sha = request.ExpectedSha256, caption = request.ControlCaption;
            int sourceLine = request.StartLine, expectedMode = request.ExpectedMode;
            try
            {
                context.Post(_ => {
                    operation.State = "Running";
                    try
                    {
                        dynamic currentProject = GetProject(projectSelector);
                        if ((int)currentProject.Mode != expectedMode) throw new InvalidOperationException("Project mode changed before navigation.");
                        ValidateLocation(new Request { StartLine = sourceLine, ExpectedSha256 = sha }, GetModule(currentProject, moduleName));
                        if (PositionKey(operation.Before) != PositionKey(NativeNavigationPosition())) throw new InvalidOperationException("Active source position changed before navigation.");
                        dynamic command = vbe.CommandBars.FindControl(1, id);
                        if (command == null || !(bool)command.Enabled || (string)command.Caption != caption)
                            throw new InvalidOperationException("Native navigation command is no longer available.");
                        // A visible/active CodePane can coexist with keyboard focus in chat.
                        // Native Definition consumes the focused editor selection.
                        ((dynamic)vbe.ActiveCodePane).Window.SetFocus();
                        if (PositionKey(operation.Before) != PositionKey(NativeNavigationPosition())) throw new InvalidOperationException("Source position changed while focusing the native editor.");
                        command.Execute();
                        operation.CommandCompleted = true;
                        operation.ReadbackDeadlineUtc = DateTime.UtcNow.AddSeconds(2);
                        operation.State = "Observing";
                        // The in-process native command may post its navigation messages.
                        // Read after returning to the host message loop, never inside Execute.
                        context.Post(__ => ObserveNavigation(operation), null);
                    }
                    catch (Exception ex) { operation.Error = ex.Message; operation.State = "Failed"; }
                }, null);
            }
            catch (Exception ex) { operation.Error = ex.Message; operation.State = "Failed"; }
            return NavigationResult(operation);
        }
        private void ObserveNavigation(NavigationOperation operation)
        {
            if (operation.State != "Observing") return;
            if (DateTime.UtcNow > operation.ReadbackDeadlineUtc) { operation.State = "Completed"; return; }
            try
            {
                operation.After = NativeNavigationPosition();
                operation.NavigationObserved = NavigationChanged(operation.Before, operation.After);
                if (operation.NavigationObserved) operation.State = "Completed";
            }
            catch (Exception ex) { operation.Error = ex.Message; operation.State = "Failed"; }
        }

        private object NativeNavigationPosition()
        {
            dynamic window = vbe.ActiveWindow;
            object activeWindow = window == null ? null : new { Type = (int)window.Type, Caption = (string)window.Caption };
            bool browserVisible = false;
            foreach (dynamic candidate in vbe.Windows)
                if ((int)candidate.Type == 2 && (bool)candidate.Visible) browserVisible = true;
            dynamic pane = vbe.ActiveCodePane;
            if (pane == null) return new { ActiveWindow = activeWindow, ObjectBrowserVisible = browserVisible, CodePaneAvailable = false };
            dynamic component = pane.CodeModule.Parent;
            string project = (string)component.Collection.Parent.Name;
            try { string path = (string)component.Collection.Parent.FileName; if (!string.IsNullOrWhiteSpace(path)) project = path; } catch { }
            int line = 0, column = 0, end = 0, endColumn = 0;
            pane.GetSelection(ref line, ref column, ref end, ref endColumn);
            return new { ActiveWindow = activeWindow, ObjectBrowserVisible = browserVisible, CodePaneAvailable = true, Project = project, Module = (string)component.Name, StartLine = line, StartColumn = column, EndLine = end, EndColumn = endColumn };
        }
        internal static bool NavigationChanged(dynamic before, dynamic after)
        {
            // Captions can change when VBE restores/maximizes its MDI windows;
            // that alone is not a navigation. Object Browser is a distinct destination.
            if ((bool)after.ObjectBrowserVisible && !(bool)before.ObjectBrowserVisible) return true;
            if (after.ActiveWindow != null && (int)after.ActiveWindow.Type == 2 &&
                (before.ActiveWindow == null || (int)before.ActiveWindow.Type != 2)) return true;
            if (!(bool)before.CodePaneAvailable || !(bool)after.CodePaneAvailable) return false;
            return before.Project != after.Project || before.Module != after.Module ||
                before.StartLine != after.StartLine || before.StartColumn != after.StartColumn ||
                before.EndLine != after.EndLine || before.EndColumn != after.EndColumn;
        }
        private static string PositionKey(object position) { return new JavaScriptSerializer().Serialize(position); }
        private static object NavigationResult(NavigationOperation operation)
        {
            return new { OperationId = operation.Id, operation.Action, operation.State, operation.CommandCompleted,
                operation.NavigationObserved, Before = operation.Before, After = operation.After, NativeError = operation.Error,
                DefinitionResolved = false, VerificationPending = true,
                NextRead = operation.State == "Queued" || operation.State == "Running" || operation.State == "Observing" ? "native_code_navigation status (Query=OperationId), debug_dialog" : "vbe_windows, code_panes, read_module or debug_dialog",
                Limit = "Native command completion and a changed selection or active window do not prove semantic resolution. Inspect the destination. A native modal dialog can outlive command completion; inspect debug_dialog when no navigation is observed." };
        }
    }
}
