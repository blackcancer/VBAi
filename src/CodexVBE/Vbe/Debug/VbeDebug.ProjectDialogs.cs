using System;
using System.Linq;
using System.Threading;

namespace CodexVBE
{
    /// <summary>Ouvre les propriétés natives d’un projet après vérification de son identité active.</summary>
    internal sealed partial class VbeDebug
    {
        /// <summary>Programme l'ouverture du dialogue natif des propriétés du projet actif exact.</summary>
        /// <param name="request">Projet, version attendue et légende exacte de la commande 2578.</param>
        /// <param name="versionCheck">Garde de version du projet appelée avant programmation puis avant exécution UI.</param>
        /// <returns>Nom natif du projet à transmettre au worker, sans modification de sélection.</returns>
        public object QueueProjectPropertiesDialog(Request request, Func<Request, bool> versionCheck)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) ||
                string.IsNullOrWhiteSpace(request.ExpectedProjectVersion) ||
                string.IsNullOrWhiteSpace(request.ControlCaption) || versionCheck == null)
                throw new ArgumentException("Project, ExpectedProjectVersion and exact ControlCaption are required.");
            var captured = new Request { Project = request.Project, ExpectedProjectVersion = request.ExpectedProjectVersion,
                ControlCaption = request.ControlCaption };
            object project = GetProject(captured.Project);
            RequireProjectPropertiesSelection(project, captured, versionCheck);
            string name = (string)((dynamic)project).Name;
            var commands = EnumerateCommands().Where(entry => entry.Id == 2578 && entry.Enabled &&
                string.Equals(entry.Caption, captured.ControlCaption, StringComparison.Ordinal)).ToList();
            if (commands.Count == 0)
                throw new InvalidOperationException("The exact native project properties command 2578 is unavailable.");
            var command = commands[0];
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            context.Post(_ =>
            {
                int stage = 0;
                try
                {
                    object current = GetProject(captured.Project);
                    if (!SameComObject(current, project) || (string)((dynamic)current).Name != name)
                        throw new InvalidOperationException("Project identity changed.");
                    stage = 1;
                    RequireProjectPropertiesSelection(current, captured, versionCheck);
                    stage = 2;
                    var liveCommand = ResolvePostedNativeCommand(command);
                    stage = 3;
                    ((dynamic)liveCommand.Control).Execute();
                }
                catch { LoadLog.Write("Project properties dialog was not executed; native preflight stage " + stage + " failed."); }
            }, null);
            return new { Scheduled = true, Project = captured.Project, ProjectName = name,
                ControlId = 2578, ControlCaption = captured.ControlCaption };
        }

                /// <summary>Résout une commande différée par chemin/ID/légende/type sans conserver un proxy Office instable.</summary>
                /// <param name="captured">Commande native capturée avant la programmation du travail UI.</param>
                /// <returns>Unique commande encore présente avec le même chemin, identifiant, légende et type.</returns>
        private CommandEntry ResolvePostedNativeCommand(CommandEntry captured)
        {
            var matches = EnumerateCommands().Where(entry => entry.Id == captured.Id && entry.Enabled &&
                entry.Caption == captured.Caption && entry.Path == captured.Path).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The exact native command path changed or is ambiguous.");
            if (System.Runtime.InteropServices.Marshal.IsComObject(captured.Control) &&
                (!System.Runtime.InteropServices.Marshal.IsComObject(matches[0].Control) ||
                 (int)((dynamic)captured.Control).Type != (int)((dynamic)matches[0].Control).Type))
                throw new InvalidOperationException("The native command type changed.");
            return matches[0];
        }

                /// <summary>Exige l'identité active exacte, le mode conception, l'absence de verrouillage et la version attendue.</summary>
                /// <param name="selected">Projet sélectionné au moment du contrôle.</param>
                /// <param name="request">Identité et version attendues du projet.</param>
                /// <param name="versionCheck">Contrôle la version de projet demandée.</param>
        private void RequireProjectPropertiesSelection(object selected, Request request, Func<Request, bool> versionCheck)
        {
            dynamic project = selected;
            if ((int)project.Mode != 2 || (int)project.Protection != 0 ||
                !SameComObject((object)vbe.ActiveVBProject, selected) || !versionCheck(request))
                throw new InvalidOperationException("Select the exact unprotected project in design mode and refresh its version.");
        }
    }
}
