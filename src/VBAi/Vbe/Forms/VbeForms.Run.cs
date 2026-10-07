using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Exécute une commande Run native après validation de l’identité et des révisions du UserForm.</summary>
    internal sealed partial class VbeForms
    {

        /// <summary>État suivi d’une commande native de lancement de formulaire.</summary>
        private sealed class FormRunOperation
        {

            /// <summary>Session-local operation ID, target project and form, current state, and failure message.</summary>
            public string Id, Project, Form, State, Error;

            /// <summary>True only after the native Run command returned; this does not verify form display or initialization.</summary>
            public bool CommandCompleted;
        }

        /// <summary>Opérations de lancement récentes conservées pour interrogation de leur état.</summary>
        private readonly List<FormRunOperation> formRuns = new List<FormRunOperation>();

        /// <summary>Retourne l’état d’une opération Run identifiée dans le projet demandé.</summary>
        /// <param name="request">Identifiants de projet et d’opération transmis dans Query.</param>
        /// <returns>État courant de la commande et indicateur explicite que l’affichage réel n’est pas vérifié.</returns>
        /// <exception cref="ArgumentException">Aucune opération ne correspond au projet et à l’identifiant fournis.</exception>
        public object FormRunStatus(Request request)
        {
            var operation = formRuns.SingleOrDefault(item => item.Id == request.Query && item.Project == request.Project);
            return operation == null
                ? throw new ArgumentException("Unknown UserForm run operation for this project.")
                : FormRunSnapshot(operation);
        }

        /// <summary>Place l’exécution de la commande Run sur le contexte UI après deux contrôles des versions et de l’identité COM.</summary>
        /// <param name="request">Projet, formulaire, révisions source/arbre et légende exacte de la commande.</param>
        /// <returns>Instantané de l’opération mise en file; il ne confirme pas que le formulaire s’est affiché.</returns>
        /// <exception cref="ArgumentException">Une version, un projet, un formulaire ou la légende de commande manque.</exception>
        /// <exception cref="InvalidOperationException">Le contexte UI, l’identité du formulaire ou la commande Run ne peut pas être validé.</exception>
        public object RunForm(Request request)
        {
            if (request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.Form) || string.IsNullOrWhiteSpace(request.ExpectedSha256) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion) || string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("Project, Form, ExpectedMode=2, ExpectedSha256, ExpectedTreeVersion and exact Run ControlCaption are required.");
            var context = SynchronizationContext.Current ?? throw new InvalidOperationException("A VBE UI synchronization context is required.");
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

        /// <summary>Construit le résultat public de suivi sans confondre le retour de commande avec une preuve d’affichage.</summary>
        /// <param name="operation">État interne à exposer.</param>
        /// <returns>Opération, état, erreur éventuelle et lectures natives recommandées.</returns>
        private static object FormRunSnapshot(FormRunOperation operation)
        {
            return new
            {
                OperationId = operation.Id,
                operation.Project,
                operation.Form,
                operation.State,
                operation.CommandCompleted,
                operation.Error,
                RuntimeVerified = false,
                NextRead = "form_run_status, read_runtime_forms, debug_state and debug_dialog; command return alone does not prove that the UserForm was displayed or initialized."
            };
        }

        /// <summary>Vérifie l’empreinte du code et la version de l’arbre avant de lancer le formulaire.</summary>
        /// <param name="form">UserForm sélectionné.</param>
        /// <param name="expectedSha">SHA-256 du code lu précédemment.</param>
        /// <param name="expectedTree">Version de l’arbre Designer lue précédemment.</param>
        /// <exception cref="InvalidOperationException">Le code ou l’arbre a changé depuis l’inspection.</exception>
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

        /// <summary>Active la fenêtre Designer et vérifie que le projet et le formulaire exacts sont au premier plan.</summary>
        /// <param name="project">Projet VBA attendu comme actif.</param>
        /// <param name="form">UserForm dont la fenêtre doit recevoir le focus.</param>
        /// <exception cref="InvalidOperationException">La fenêtre active ou le projet ne correspond pas à la cible.</exception>
        private void FocusFormRun(object project, dynamic form)
        {
            dynamic window = form.DesignerWindow();
            window.Visible = true;
            window.SetFocus();
            if (!SameComIdentity(project, (object)vbe.ActiveVBProject) || !SameComIdentity((object)window, (object)vbe.ActiveWindow))
                throw new InvalidOperationException("The exact UserForm designer is not active.");
        }

        /// <summary>Récupère la commande VBE d’identifiant 186 et vérifie sa disponibilité et sa légende observée.</summary>
        /// <param name="caption">Légende exacte obtenue depuis l’inventaire des commandes.</param>
        /// <returns>Contrôle natif activable correspondant à Run.</returns>
        /// <exception cref="InvalidOperationException">La commande exacte est absente, désactivée ou renommée.</exception>
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
