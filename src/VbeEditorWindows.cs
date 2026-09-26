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
            Read(fields, errors, "Docked", () => (bool)window.Docked);
            return new { Index = index, Properties = fields, Errors = errors };
        }

        private static object PaneSnapshot(dynamic pane, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Module", () => (string)pane.CodeModule.Parent.Name);
            Read(fields, errors, "Project", () => (string)pane.CodeModule.Parent.Collection.Parent.Name);
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
