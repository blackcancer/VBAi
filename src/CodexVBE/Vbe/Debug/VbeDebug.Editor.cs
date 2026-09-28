using System;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        public object SetCodeSplit(Request request)
        {
            if (request.Action != "split" && request.Action != "unsplit") throw new ArgumentException("Use split or unsplit.");
            if (string.IsNullOrWhiteSpace(request.ControlCaption)) throw new ArgumentException("ControlCaption from native list_commands ID 302 is required.");
            dynamic project = GetProject(request.Project);
            int mode = (int)project.Mode;
            if ((mode != 1 && mode != 2) || request.ExpectedMode != mode)
                throw new InvalidOperationException("Code splitting requires the expected break or design mode.");
            dynamic module = GetModule(project, request.Module);
            ValidateLocation(request, module);
            // SelectCode checks the revision again and activates this exact module.
            SelectCode(request);
            dynamic active = vbe.ActiveCodePane;
            if (active == null || !SameComObject((object)active.CodeModule, (object)module))
                throw new InvalidOperationException("The requested module is not the active native code pane.");
            int before = ModulePaneCount(module), desired = request.Action == "split" ? 2 : 1;
            if (before < 1 || before > 2) throw new InvalidOperationException("Unexpected native code-pane topology.");
            if (before == desired) return new { Applied = false, Verified = true, PaneCount = before, request.Project, request.Module };
            var command = EnumerateCommands().FirstOrDefault(x => x.Id == 302 && x.Enabled && string.Equals(x.Caption, request.ControlCaption, StringComparison.Ordinal));
            if (command == null) throw new InvalidOperationException("The exact native Split command is absent or disabled. Refresh list_commands.");
            string error = null;
            try { ((dynamic)command.Control).Execute(); } catch (Exception ex) { error = ex.Message; }
            int? after = null;
            try { after = ModulePaneCount(module); } catch (Exception ex) { error = error ?? ex.Message; }
            bool verified = error == null && after == desired;
            return new { Applied = error == null ? (bool?)true : null, Verified = verified, VerificationPending = !verified,
                request.Project, request.Module, BeforePaneCount = before, PaneCount = after, NativeError = error,
                PersistenceVerified = false, NextRead = verified ? null : "code_panes" };
        }

        private int ModulePaneCount(object module)
        {
            int count = 0;
            foreach (dynamic pane in vbe.CodePanes)
                if (TryLivePaneModule((object)pane, out object paneModule, out string ignored) && SameComObject(paneModule, module)) count++;
            return count;
        }
    }
}
