using System;
using System.Linq;
using System.Threading;

namespace CodexVBE
{
    /// <summary>Planifie l’ouverture de surfaces IDE natives sous contrôle du contexte VBE courant.</summary>
internal sealed partial class VbeDebug
    {
                /// <summary>Source exacte utilisée pour garder la cible du dialogue Imprimer.</summary>
        /// <param name="code">Module de code à relire.</param>
        /// <returns>Texte complet du module, ou chaîne vide si le module ne contient aucune ligne.</returns>
private static string NativeDialogSource(dynamic code) => (int)code.CountOfLines == 0 ? "" : (string)code.Lines[1, (int)code.CountOfLines];

                /// <summary>Ouvre uniquement une surface IDE native reconnue, sans valider ni lancer une impression.</summary>
        /// <param name="request">Projet, action, mode et éventuellement module, SHA et légende de commande.</param>
        /// <returns>Confirmation de mise en file; l’exécution native et les effets du dialogue ne sont pas attestés.</returns>
public object QueueNativeIdeDialog(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("Project, Action and the exact native ControlCaption are required.");
            int id;
            switch (request.Action)
            {
                case "macros": id = 930; break;
                case "help": id = 984; break;
                case "print": id = 4; break;
                case "additional_controls": id = 642; break;
                default: throw new ArgumentException("Action must be macros/help/print/additional_controls.");
            }
            object project = GetProject(request.Project);
            if (!SameComObject((object)vbe.ActiveVBProject, project) || (int)((dynamic)project).Mode != request.ExpectedMode ||
                request.ExpectedMode != 1 && request.ExpectedMode != 2)
                throw new InvalidOperationException("Select the exact project in the inspected break or design mode.");
            string module = request.Module, sha = request.ExpectedSha256;
            if (request.Action == "print")
            {
                dynamic code = GetModule(project, module);
                if (string.IsNullOrWhiteSpace(sha) || !string.Equals(Hash(NativeDialogSource(code)), sha, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Inspect the exact module source before opening its Print dialog.");
                dynamic active = vbe.ActiveCodePane;
                if (active == null || !SameComObject((object)active.CodeModule, (object)code))
                    throw new InvalidOperationException("The module to print must already be the active code pane.");
            }
            var command = EnumerateCommands().FirstOrDefault(c => c.Id == id && c.Enabled && c.Caption == request.ControlCaption);
            if (command == null) throw new InvalidOperationException("The exact native IDE command is absent or disabled.");
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            string selector = request.Project, caption = request.ControlCaption, action = request.Action;
            int mode = request.ExpectedMode;
            context.Post(_ => {
                try
                {
                    object current = GetProject(selector);
                    if (!SameComObject(current, project) || !SameComObject((object)vbe.ActiveVBProject, current) || (int)((dynamic)current).Mode != mode)
                        throw new InvalidOperationException("Native IDE context changed.");
                    if (action == "print")
                    {
                        dynamic code = GetModule(current, module);
                        dynamic active = vbe.ActiveCodePane;
                        if (active == null || !SameComObject((object)active.CodeModule, (object)code) ||
                            !string.Equals(Hash(NativeDialogSource(code)), sha, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Print source context changed.");
                    }
                    var liveCommand = ResolvePostedNativeCommand(command);
                    ((dynamic)liveCommand.Control).Execute();
                }
                catch { LoadLog.Write("Native IDE surface was not opened because its context changed or the command failed."); }
            }, null);
            return new { Scheduled = true, Project = selector, Action = action, ControlId = id,
                CommandReturned = false, Verified = false, PrintSubmitted = false,
                NextRead = action == "macros" ? "read_navigation_surface Pane=macros" : "debug_dialog",
                Limit = "Scheduling only. Native Help may open in another application; no topic, print job or installed control is certified." };
        }
    }

    /// <summary>Vérifications des dialogues natifs du VBE avant les opérations de fenêtre.</summary>
internal static partial class VbeDebugWindows
    {
        /// <summary>Refuse de réutiliser ou de modifier un dialogue de propriétés déjà ouvert par l'utilisateur.</summary>
        internal static void EnsureNoProjectPropertiesDialog()
        {
            bool found = false;
            uint current = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            EnumWindows((handle, ignored) => {
                GetWindowThreadProcessId(handle, out uint pid);
                if (pid == current && ClassName(handle) == "#32770" && IsWindowVisible(handle))
                {
                    string title = WindowText(handle);
                    if (title.EndsWith(" - Project Properties", StringComparison.Ordinal) || title.EndsWith(" - Propriétés du projet", StringComparison.Ordinal))
                    { found = true; return false; }
                }
                return true;
            }, IntPtr.Zero);
            if (found) throw new InvalidOperationException("A native project properties dialog is already open; finish or cancel it first.");
        }
    }
}
