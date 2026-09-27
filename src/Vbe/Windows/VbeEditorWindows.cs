using System;
using System.Collections.Generic;

namespace CodexVBE
{
    // Reads the VBIDE collections directly. In particular, do not use
    // VBComponent.CodePane for inspection: its getter opens and activates a pane.
    internal sealed class VbeEditorWindows
    {
        private readonly dynamic vbe;

        public VbeEditorWindows(object vbe) { this.vbe = vbe; }

        public object Windows()
        {
            var windows = new List<object>();
            int index = 0;
            foreach (dynamic window in vbe.Windows)
                windows.Add(WindowSnapshot(window, ++index));
            object active = null;
            try
            {
                dynamic window = vbe.ActiveWindow;
                if (window != null) active = WindowSnapshot(window, null);
            }
            catch (Exception ex) { active = new { Error = ex.Message }; }
            return new { Windows = windows, ActiveWindow = active };
        }

        public object CodePanes()
        {
            var panes = new List<object>();
            int index = 0;
            foreach (dynamic pane in vbe.CodePanes)
                panes.Add(PaneSnapshot(pane, ++index));
            object active = null;
            try
            {
                dynamic pane = vbe.ActiveCodePane;
                if (pane != null) active = PaneSnapshot(pane, null);
            }
            catch (Exception ex) { active = new { Error = ex.Message }; }
            return new { CodePanes = panes, ActiveCodePane = active };
        }

        public object Environment()
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Version", () => (string)vbe.Version);
            Read(fields, errors, "ProjectCount", () => (int)vbe.VBProjects.Count);
            Read(fields, errors, "WindowCount", () => (int)vbe.Windows.Count);
            Read(fields, errors, "CodePaneCount", () => (int)vbe.CodePanes.Count);
            Read(fields, errors, "AddInCount", () => (int)vbe.AddIns.Count);
            Read(fields, errors, "ActiveProject", () => (string)vbe.ActiveVBProject.Name);
            return new { Properties = fields, Errors = errors };
        }

        public object AddIns()
        {
            var addIns = new List<object>();
            int index = 0;
            foreach (dynamic addIn in vbe.AddIns)
            {
                var fields = new Dictionary<string, object>();
                var errors = new Dictionary<string, string>();
                Read(fields, errors, "ProgId", () => (string)addIn.ProgId);
                Read(fields, errors, "Guid", () => (string)addIn.Guid);
                Read(fields, errors, "Description", () => (string)addIn.Description);
                Read(fields, errors, "Connect", () => (bool)addIn.Connect);
                addIns.Add(new { Index = ++index, Properties = fields, Errors = errors });
            }
            return new { AddIns = addIns, Count = addIns.Count,
                Scope = "VBE.AddIns contains VBE-registered add-ins, not the host application's COMAddIns." };
        }

        public object FocusWindow(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            if (!(bool)target.Visible)
                throw new InvalidOperationException("The requested window is hidden; SetFocus requires a visible window.");
            target.SetFocus();
            dynamic active = vbe.ActiveWindow;
            bool verified = active != null &&
                string.Equals((string)active.Caption, caption, StringComparison.Ordinal) &&
                (int)active.Type == type;
            return new { WindowCaption = caption, WindowType = type, SetFocusInvoked = true,
                Verification = verified ? "ActiveWindowReadback" : "Unverified",
                ActiveWindow = active == null ? null : WindowSnapshot(active, null) };
        }

        public object ShowWindow(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            bool wasVisible = (bool)target.Visible;
            if (!wasVisible) target.Visible = true;
            bool nowVisible = (bool)target.Visible;
            if (!nowVisible)
                throw new InvalidOperationException("The native VBE window did not become visible.");
            target.SetFocus();
            dynamic active = vbe.ActiveWindow;
            bool activeVerified = active != null &&
                string.Equals((string)active.Caption, caption, StringComparison.Ordinal) &&
                (int)active.Type == type;
            return new { WindowCaption = caption, WindowType = type, WasVisible = wasVisible,
                Visible = nowVisible, FocusVerified = activeVerified,
                ActiveWindow = active == null ? null : WindowSnapshot(active, null) };
        }

        public object WindowLinkage(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Visible", () => (bool)target.Visible);
            try
            {
                dynamic frame = target.LinkedWindowFrame;
                fields["IsLinked"] = frame != null;
                if (frame != null)
                {
                    Read(fields, errors, "FrameCaption", () => (string)frame.Caption);
                    var linked = new List<object>();
                    foreach (dynamic window in frame.LinkedWindows)
                        linked.Add(new { Caption = (string)window.Caption, Type = (int)window.Type });
                    fields["LinkedWindows"] = linked;
                }
            }
            catch (Exception ex) { errors["LinkedWindowFrame"] = ex.Message; }
            return new { WindowCaption = caption, WindowType = type,
                Properties = fields, Errors = errors };
        }

        public object CloseWindow(string caption, int type)
        {
            if (string.Equals(caption, "CodexVBE", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The CodexVBE tool window cannot close itself through this command.");
            dynamic target = FindExactWindow(caption, type);
            if (!(bool)target.Visible)
                throw new InvalidOperationException("The requested VBE window is already hidden.");
            target.Close();
            int matches = 0;
            bool visible = false;
            foreach (dynamic window in vbe.Windows)
                if (string.Equals((string)window.Caption, caption, StringComparison.Ordinal) &&
                    (int)window.Type == type)
                { matches++; visible |= (bool)window.Visible; }
            return new { WindowCaption = caption, WindowType = type, CloseInvoked = true,
                Verification = matches == 0 ? "RemovedFromWindows" :
                    matches == 1 && !visible ? "HiddenInWindows" : "Unverified",
                RemainingMatches = matches, RemainingVisible = visible };
        }

        private dynamic FindExactWindow(string caption, int type)
        {
            if (string.IsNullOrWhiteSpace(caption) || type < 0)
                throw new ArgumentException("WindowCaption and WindowType from vbe_windows are required.");
            dynamic target = null;
            foreach (dynamic window in vbe.Windows)
            {
                if (!string.Equals((string)window.Caption, caption, StringComparison.Ordinal) ||
                    (int)window.Type != type) continue;
                if (target != null)
                    throw new InvalidOperationException("More than one VBE window matches the requested caption and type.");
                target = window;
            }
            if (target == null) throw new InvalidOperationException("The requested VBE window is no longer present.");
            return target;
        }

        private static object WindowSnapshot(dynamic window, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Caption", () => (string)window.Caption);
            Read(fields, errors, "Type", () => (int)window.Type);
            Read(fields, errors, "Visible", () => (bool)window.Visible);
            Read(fields, errors, "WindowState", () => (int)window.WindowState);
            Read(fields, errors, "Left", () => (int)window.Left);
            Read(fields, errors, "Top", () => (int)window.Top);
            Read(fields, errors, "Width", () => (int)window.Width);
            Read(fields, errors, "Height", () => (int)window.Height);
            return new { Index = index, Properties = fields, Errors = errors };
        }

        private static object PaneSnapshot(dynamic pane, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Module", () => (string)pane.CodeModule.Parent.Name);
            Read(fields, errors, "Project", () => (string)pane.CodeModule.Parent.Collection.Parent.Name);
            Read(fields, errors, "ProjectPath", () => (string)pane.CodeModule.Parent.Collection.Parent.FileName);
            Read(fields, errors, "CodePaneView", () => (int)pane.CodePaneView);
            Read(fields, errors, "TopLine", () => (int)pane.TopLine);
            Read(fields, errors, "CountOfVisibleLines", () => (int)pane.CountOfVisibleLines);
            Read(fields, errors, "WindowCaption", () => (string)pane.Window.Caption);
            try
            {
                int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
                pane.GetSelection(ref startLine, ref startColumn, ref endLine, ref endColumn);
                fields["Selection"] = new { StartLine = startLine, StartColumn = startColumn,
                    EndLine = endLine, EndColumn = endColumn };
            }
            catch (Exception ex) { errors["Selection"] = ex.Message; }
            return new { Index = index, Properties = fields, Errors = errors };
        }

        private static void Read(IDictionary<string, object> fields,
            IDictionary<string, string> errors, string name, Func<object> getter)
        {
            try { fields[name] = getter(); }
            catch (Exception ex) { errors[name] = ex.Message; }
        }
    }
}
