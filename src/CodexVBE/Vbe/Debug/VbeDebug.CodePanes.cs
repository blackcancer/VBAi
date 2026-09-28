using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        private readonly Dictionary<string, object> inspectedPanes = new Dictionary<string, object>(StringComparer.Ordinal);

        private readonly Dictionary<string, string> inspectedPaneVersions = new Dictionary<string, string>(StringComparer.Ordinal);

        public object CodePaneLayout(Request request)
        {
            dynamic module = GetModule(GetProject(request.Project), request.Module);
            inspectedPanes.Clear();
            inspectedPaneVersions.Clear();
            var results = new List<object>();
            var unavailable = new List<object>();
            int nativeIndex = 0;
            foreach (dynamic pane in vbe.CodePanes)
            {
                nativeIndex++;
                if (!TryLivePaneModule((object)pane, out object paneModule, out string paneError))
                { unavailable.Add(new { Index = nativeIndex, Error = paneError }); continue; }
                if (!SameComObject(paneModule, (object)module)) continue;
                if (results.Count >= 2) throw new InvalidOperationException("Unexpected native code-pane topology.");
                string token = Guid.NewGuid().ToString("N");
                object snapshot = CodePaneState(pane);
                inspectedPanes.Add(token, (object)pane);
                inspectedPaneVersions.Add(token, PaneVersion(snapshot));
                results.Add(new { Pane = token, State = snapshot, WindowVersion = PaneVersion(snapshot) });
            }
            return new { request.Project, request.Module, Panes = results, UnavailablePanes = unavailable,
                Scope = "Open panes only. Tokens expire at the next code_pane_layout call in this session." };
        }

        public object ScrollCodePane(Request request)
        {
            object raw;
            if (string.IsNullOrWhiteSpace(request.Pane) || !inspectedPanes.TryGetValue(request.Pane, out raw))
                throw new InvalidOperationException("Read code_pane_layout and use its current Pane token.");
            dynamic pane = raw;
            dynamic project = GetProject(request.Project);
            int mode = (int)project.Mode;
            if ((mode != 1 && mode != 2) || request.ExpectedMode != mode)
                throw new InvalidOperationException("Scrolling requires the expected design or break mode.");
            dynamic module = GetModule(project, request.Module);
            if (!SameComObject((object)pane.CodeModule, (object)module))
                throw new InvalidOperationException("The pane belongs to another module.");
            bool exists = false;
            foreach (dynamic candidate in vbe.CodePanes)
                if (SameComObject((object)candidate, raw)) exists = true;
            if (!exists) throw new InvalidOperationException("The native pane was closed or replaced. Read code_pane_layout again.");
            ValidateLocation(request, module);
            object before = CodePaneState(pane);
            if (string.IsNullOrWhiteSpace(request.ExpectedWindowVersion) || !string.Equals(request.ExpectedWindowVersion, PaneVersion(before), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The pane changed. Read code_pane_layout again.");
            string error = null;
            try { pane.TopLine = request.StartLine; } catch (Exception ex) { error = ex.Message; }
            object after = null; string readError = null;
            try { after = CodePaneState(pane); } catch (Exception ex) { readError = ex.Message; }
            dynamic oldState = before, newState = after;
            bool verified = error == null && readError == null && (int)newState.TopLine == request.StartLine &&
                oldState.Selection == newState.Selection && oldState.Sha256 == newState.Sha256;
            return new { Applied = error == null ? (bool?)true : null, Verified = verified, VerificationPending = !verified,
                Before = before, After = after, NativeError = error, ReadbackError = readError,
                WindowVersion = after == null ? null : PaneVersion(after), PersistenceVerified = false };
        }

        public object SetCodePaneView(Request request)
        {
            if (request.Action != "procedure" && request.Action != "module")
                throw new ArgumentException("Action must be procedure or module.");
            if (string.IsNullOrWhiteSpace(request.Pane) || !inspectedPanes.TryGetValue(request.Pane, out object raw))
                throw new InvalidOperationException("Read code_pane_layout and use its current Pane token.");
            dynamic pane = raw, project = GetProject(request.Project);
            int mode = (int)project.Mode;
            if ((mode != 1 && mode != 2) || request.ExpectedMode != mode)
                throw new InvalidOperationException("Changing code view requires the expected design or break mode.");
            dynamic module = GetModule(project, request.Module);
            if (!SameComObject((object)pane.CodeModule, (object)module)) throw new InvalidOperationException("The pane belongs to another module.");
            var windowPanes = new List<object>();
            foreach (dynamic candidate in vbe.CodePanes)
                if (TryLivePaneModule((object)candidate, out object candidateModule, out string ignored) && SameComObject(candidateModule, (object)module))
                    windowPanes.Add((object)candidate);
            bool exists = windowPanes.Any(candidate => SameComObject(candidate, raw));
            if (windowPanes.Count < 1 || windowPanes.Count > 2) throw new InvalidOperationException("Unexpected native code-window topology.");
            if (!exists) throw new InvalidOperationException("The inspected code pane was closed or replaced.");
            ValidateLocation(request, module);
            object before = CodePaneState(pane);
            if (string.IsNullOrWhiteSpace(request.ExpectedWindowVersion) || !string.Equals(request.ExpectedWindowVersion, PaneVersion(before), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The pane changed. Read code_pane_layout again.");
            foreach (object candidate in windowPanes)
            {
                string token = inspectedPanes.Where(pair => SameComObject(pair.Value, candidate)).Select(pair => pair.Key).SingleOrDefault();
                if (token == null || inspectedPaneVersions[token] != PaneVersion(CodePaneState((dynamic)candidate)))
                    throw new InvalidOperationException("A pane in the code window changed. Read code_pane_layout again.");
            }
            object[] beforePanes = windowPanes.Select(candidate => (object)CodePaneState((dynamic)candidate)).ToArray();
            int target = request.Action == "procedure" ? 0 : 1;
            if (beforePanes.All(state => (int)((dynamic)state).View == target))
                return new { Verified = true, Changed = false, Scope = "Entire code window, including both split panes", Before = before, After = before, BeforePanes = beforePanes, AfterPanes = beforePanes };
            pane.Show(); pane.Window.SetFocus();
            if (!SameComObject((object)vbe.ActiveCodePane, raw) || PaneVersion(CodePaneState(pane)) != PaneVersion(before))
                throw new InvalidOperationException("The target pane or viewport changed while focusing it.");
            for (int index = 0; index < windowPanes.Count; index++)
                if (PaneVersion(CodePaneState((dynamic)windowPanes[index])) != PaneVersion(beforePanes[index]))
                    throw new InvalidOperationException("A pane changed while focusing the code window.");
            object native = VbeDebugWindows.ChangeCodeView((string)pane.Window.Caption, target == 0);
            object after = CodePaneState(pane);
            object[] afterPanes = windowPanes.Select(candidate => (object)CodePaneState((dynamic)candidate)).ToArray();
            bool verified = afterPanes.All(state => (int)((dynamic)state).View == target &&
                (string)((dynamic)state).Sha256 == (string)((dynamic)before).Sha256);
            return new { Verified = verified, Changed = true, Scope = "Entire code window, including both split panes", BeforePanes = beforePanes, AfterPanes = afterPanes, VerificationPending = !verified, Before = before, After = after,
                Native = native, WindowVersion = PaneVersion(after), PersistenceVerified = false,
                NextRead = "code_pane_layout; native view changes can also change the visible range or selection." };
        }

        internal static bool TryLivePaneModule(object pane, out object module, out string error)
        {
            try { module = ((dynamic)pane).CodeModule; error = null; return true; }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x80020010))
            {
                // VBE can retain a destroyed split-pane entry in CodePanes.
                // Preserve the diagnostic; never create a token for that entry.
                module = null; error = ex.Message; return false;
            }
        }

        private static object CodePaneState(dynamic pane)
        {
            int start = 0, column = 0, end = 0, endColumn = 0;
            pane.GetSelection(ref start, ref column, ref end, ref endColumn);
            dynamic module = pane.CodeModule;
            int count = (int)module.CountOfLines;
            string code = count == 0 ? "" : (string)module.Lines[1, count];
            return new { TopLine = (int)pane.TopLine, VisibleLines = (int)pane.CountOfVisibleLines,
                View = (int)pane.CodePaneView, Selection = start + ":" + column + ":" + end + ":" + endColumn,
                Sha256 = Hash(code), LineCount = count };
        }
        private static string PaneVersion(object state) { return Hash(new JavaScriptSerializer().Serialize(state)); }
    }
}
