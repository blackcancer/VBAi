using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")]
        private static extern bool HistoryWindowEnabled(IntPtr window);

        public object NativeCodeHistoryState(Request request)
        {
            var snapshot = CaptureCodeHistory(request.Project);
            return new { request.Project, HistoryVersion = snapshot.Version,
                Modules = snapshot.Modules.Select(x => new { x.Module, x.Sha256 }).ToArray(),
                Commands = snapshot.Commands, Scope = "Native shared code history of the only open project; may change another module.",
                Limit = "Projects containing UserForms and multiple open projects are refused. Native stack entries are not exposed." };
        }

        public object NativeCodeHistory(Request request)
        {
            int id = request.Action == "undo" ? 128 : request.Action == "redo" ? 129 : 0;
            if (id == 0) throw new ArgumentException("Action must be undo or redo.");
            if (request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion) ||
                string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("ExpectedMode=2, ExpectedProjectVersion from native_code_history_state and exact ControlCaption are required.");
            var before = CaptureCodeHistory(request.Project);
            if (before.Version != request.ExpectedProjectVersion)
                throw new InvalidOperationException("The project or native history command changed since inspection.");
            if (!HistoryWindowEnabled(new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd))))
                throw new InvalidOperationException("The VBE is disabled by a modal window.");
            var command = EnumerateCommands().FirstOrDefault(x => x.Id == id && x.Enabled &&
                string.Equals(x.Caption, request.ControlCaption, StringComparison.Ordinal));
            if (command == null) throw new InvalidOperationException("The exact native history command is absent or disabled.");
            string error = null; HistorySnapshot after = null;
            try { ((dynamic)command.Control).Execute(); }
            catch (Exception ex) { error = ex.Message; }
            try { after = CaptureCodeHistory(request.Project); }
            catch (Exception ex) { error = error ?? ex.Message; }
            var changes = new List<CodeChange>();
            bool topologyChanged = after != null && !before.Modules.Select(x => x.Module).SequenceEqual(after.Modules.Select(x => x.Module));
            if (after != null)
                foreach (var old in before.Modules)
                {
                    var current = after.Modules.FirstOrDefault(x => x.Module == old.Module);
                    if (current != null && current.Code != old.Code)
                        changes.Add(new CodeChange(request.Project, old.Module, old.Code, old.Sha256,
                            current.Code, current.Sha256, CodeRollback.Lines(current.Code).Length));
                }
            return new { request.Project, request.Action, ControlId = id, Executed = error == null,
                Changes = changes, TopologyChanged = topologyChanged,
                Verified = error == null && changes.Count > 0 && !topologyChanged,
                VerificationPending = after == null || (changes.Count == 0 && !topologyChanged),
                NativeError = error, NextRead = "native_code_history_state, read_module",
                Limit = "One shared native history action. No stack inventory or persistence guarantee. Do not retry automatically. Changes may affect a different module than the active one." };
        }

        private sealed class HistoryModule
        {
            public string Module { get; set; }
            public string Code { get; set; }
            public string Sha256 { get; set; }
        }
        private sealed class HistorySnapshot
        {
            public string Version;
            public List<HistoryModule> Modules;
            public object[] Commands;
        }
        private HistorySnapshot CaptureCodeHistory(string selector)
        {
            dynamic project = GetProject(selector);
            if ((int)vbe.VBProjects.Count != 1)
                throw new InvalidOperationException("Native history is shared: close other projects before using this command.");
            if ((int)project.Mode != 2) throw new InvalidOperationException("Native code history requires design mode.");
            var modules = new List<HistoryModule>();
            foreach (dynamic component in project.VBComponents)
            {
                if ((int)component.Type == 3)
                    throw new InvalidOperationException("Native code history refuses projects containing UserForms: the pending action may modify a Designer.");
                dynamic module = component.CodeModule;
                int count = (int)module.CountOfLines;
                string code = count == 0 ? string.Empty : (string)module.Lines[1, count];
                modules.Add(new HistoryModule { Module = (string)component.Name, Code = code, Sha256 = Hash(code) });
            }
            modules = modules.OrderBy(x => x.Module, StringComparer.Ordinal).ToList();
            object[] commands = EnumerateCommands().Where(x => x.Id == 128 || x.Id == 129)
                .Select(x => (object)new { x.Id, x.Caption, x.Enabled }).ToArray();
            var json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
            string version = Hash(json.Serialize(new { Project = selector, Modules = modules.Select(x => new { x.Module, x.Sha256 }), Commands = commands }));
            return new HistorySnapshot { Version = version, Modules = modules, Commands = commands };
        }
    }
}
