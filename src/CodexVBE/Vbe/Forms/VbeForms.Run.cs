using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        private sealed class FormRunOperation
        {
            public string Id, Project, Form, State, Error;
            public bool CommandCompleted;
        }
        private readonly List<FormRunOperation> formRuns = new List<FormRunOperation>();

        public object FormRunStatus(Request request)
        {
            var operation = formRuns.SingleOrDefault(item => item.Id == request.Query && item.Project == request.Project);
            if (operation == null) throw new ArgumentException("Unknown UserForm run operation for this project.");
            return FormRunSnapshot(operation);
        }

        public object RunForm(Request request)
        {
            if (request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) || string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) || string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("Project, Form, ExpectedMode=2, ExpectedSha256, ExpectedTreeVersion and exact Run ControlCaption are required.");
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("A VBE UI synchronization context is required.");
            if (formRuns.Any(item => item.State == "Queued" || item.State == "Running"))
                throw new InvalidOperationException("A UserForm run command is still pending.");
            string projectName = request.Project, formName = request.Form, sha = request.ExpectedSha256,
                tree = request.ExpectedTreeVersion, caption = request.ControlCaption;
            object project = GetDesignProject(projectName);
            object form = GetForm(project, formName);
            ValidateFormRun(form, sha, tree);
            FocusFormRun(project, form);
            RequireFormRunCommand(caption);
            var operation = new FormRunOperation { Id = Guid.NewGuid().ToString("N"), Project = projectName, Form = formName, State = "Queued" };
            formRuns.Add(operation);
            if (formRuns.Count > 20) formRuns.RemoveAt(0);
            try
            {
                context.Post(_ =>
                {
                    operation.State = "Running";
                    try
                    {
                        object currentProject = GetDesignProject(projectName);
                        object currentForm = GetForm(currentProject, formName);
                        if (!SameComIdentity(project, currentProject) || !SameComIdentity(form, currentForm))
                            throw new InvalidOperationException("The selected project or UserForm was replaced before execution.");
                        ValidateFormRun(currentForm, sha, tree);
                        FocusFormRun(currentProject, currentForm);
                        ValidateFormRun(currentForm, sha, tree);
                        dynamic command = RequireFormRunCommand(caption);
                        command.Execute();
                        operation.CommandCompleted = true;
                        operation.State = "CommandReturned";
                    }
                    catch (Exception ex) { operation.State = "Failed"; operation.Error = ex.Message; }
                }, null);
            }
            catch { formRuns.Remove(operation); throw; }
            return FormRunSnapshot(operation);
        }

        private static object FormRunSnapshot(FormRunOperation operation)
        {
            return new { OperationId = operation.Id, operation.Project, operation.Form, operation.State,
                operation.CommandCompleted, operation.Error, RuntimeVerified = false,
                NextRead = "form_run_status, read_runtime_forms, debug_state and debug_dialog; command return alone does not prove that the UserForm was displayed or initialized." };
        }

        private static void ValidateFormRun(dynamic form, string expectedSha, string expectedTree)
        {
            dynamic module = form.CodeModule;
            int count = (int)module.CountOfLines;
            string code = count == 0 ? string.Empty : (string)module.Lines[1, count];
            string hash;
            using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code))).Replace("-", "").ToLowerInvariant();
            if (!string.Equals(hash, expectedSha, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm code changed before execution.");
            if (!string.Equals((string)Version(form), expectedTree, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed before execution.");
        }

        private void FocusFormRun(object project, dynamic form)
        {
            dynamic window = form.DesignerWindow();
            window.Visible = true;
            window.SetFocus();
            if (!SameComIdentity(project, (object)vbe.ActiveVBProject) || !SameComIdentity((object)window, (object)vbe.ActiveWindow))
                throw new InvalidOperationException("The exact UserForm designer is not active.");
        }

        internal object RequireFormRunCommand(string caption)
        {
            dynamic command = vbe.CommandBars.FindControl(1, 186);
            // The stable native command identity and exact inspected caption work
            // for every VBE language; English/French substring checks do not.
            if (string.IsNullOrWhiteSpace(caption) || command == null || (int)command.Id != 186 ||
                !(bool)command.Enabled || (string)command.Caption != caption)
                throw new InvalidOperationException("The exact enabled native Run command (186) is required from list_commands.");
            return command;
        }
    }
}
