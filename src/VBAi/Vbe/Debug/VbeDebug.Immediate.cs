using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    internal sealed partial class VbeDebug
    {
        internal VbeDebugClipboard ImmediateClipboard = new VbeDebugClipboard();

        /// <summary>Explicit legacy-host fallback; never used by passive debugger polling.</summary>
        public Task<object> ReadImmediateAsync(Request request)
        {
            return VbeUiTask.Run(() => ReadImmediateCoreAsync(request));
        }

        private async Task<object> ReadImmediateCoreAsync(Request request)
        {
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            string phase = "initial validation";
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                (request.ExpectedMode != 1 && request.ExpectedMode != 2))
                throw new ArgumentException("Project and ExpectedMode (1 or 2) are required.");
            object project = GetProject(request.Project);
            Action validateProject = () => {
                if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                    throw new InvalidOperationException("Immediate capture left the owning VBE thread.");
                object active = vbe.ActiveVBProject;
                if ((int)((dynamic)project).Mode != request.ExpectedMode || active == null || !SameComObject(active, project))
                    throw new InvalidOperationException("The requested project must remain active in the expected mode for Immediate capture.");
            };
            validateProject();
            var matches = new List<object>();
            foreach (dynamic window in vbe.Windows)
                if ((int)window.Type == 5 && (bool)window.Visible) matches.Add((object)window);
            if (matches.Count != 1)
                throw new InvalidOperationException("Exactly one visible Immediate window is required. Open it with open_debug_pane first.");
            object target = matches[0];
            object previous = vbe.ActiveWindow;
            bool focusRestored = false;
            bool focusTouched = false;
            string text;
            Action validateTarget = () => {
                validateProject();
                object active = vbe.ActiveWindow;
                if (active == null || !SameComObject(active, target) || !(bool)((dynamic)target).Visible)
                    throw new InvalidOperationException("The Immediate window is not the active native target at " + phase +
                        "; observed window type " + (active == null ? "none" : Convert.ToString(((dynamic)active).Type)) + ". Copy was refused.");
            };
            try
            {
                text = await ImmediateClipboard.ReadAsync(() => {
                    focusTouched = true;
                    phase = "focus";
                    ((dynamic)target).SetFocus();
                    validateTarget();
                    var selectAll = FindAvailableCommand(756, entry => IsImmediateCopyCaption(entry.Caption, true));
                    if (selectAll == null) throw new InvalidOperationException("Native Select All is unavailable in Immediate.");
                    phase = "Select All";
                    validateTarget();
                    ((dynamic)selectAll.Control).Execute();
                }, () => {
                    phase = "Copy preparation";
                    validateTarget();
                    var copy = FindAvailableCommand(19, entry => IsImmediateCopyCaption(entry.Caption, false));
                    if (copy == null)
                        throw new InvalidOperationException("Native Copy is unavailable after Select All. The pane may be empty; its content is not inferred.");
                    validateTarget();
                    ((dynamic)copy.Control).Execute();
                    phase = "Copy completion";
                }, validateTarget);
            }
            finally
            {
                // Do not activate a different window after a concurrent native navigation.
                object active = vbe.ActiveWindow;
                if (focusTouched && previous != null && active != null && SameComObject(active, target))
                {
                    try
                    {
                        validateProject();
                        ((dynamic)previous).SetFocus();
                        object restored = vbe.ActiveWindow;
                        focusRestored = restored != null && SameComObject(restored, previous);
                    }
                    catch { focusRestored = false; }
                }
            }
            return new { Text = text, Method = "NativeCopy", ClipboardRestored = true,
                FocusRestored = focusRestored, SelectionRestored = false,
                Scope = "The Immediate buffer is shared by every project in this VBE instance.",
                Limit = "The text remains selected in Immediate; its previous selection cannot be restored through VBIDE. No expression was executed." };
        }

        internal static bool IsImmediateCopyCaption(string caption, bool selectAll)
        {
            string text = (caption ?? "").Replace("&", "").Trim();
            return selectAll ? string.Equals(text, "Select All", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "Sélectionner tout", StringComparison.OrdinalIgnoreCase) :
                string.Equals(text, "Copy", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "Copier", StringComparison.OrdinalIgnoreCase);
        }
    }
}
