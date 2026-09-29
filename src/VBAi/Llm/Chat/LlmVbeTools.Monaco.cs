using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Expose les outils du chat qui inspectent ou modifient le code VBA.</summary>
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Fabrique la fenêtre d’éditeur moderne, avec possibilité de la créer.</summary>
        internal Func<bool, ModernEditorWindow> MonacoWindow;
        /// <summary>Résout un module dans l’éditeur moderne pour un projet et un module précis.</summary>
        internal Func<string, string, IEditorModule> MonacoModule;
        /// <summary>Retourne la fenêtre moderne injectée ou celle portée par la session.</summary>
        /// <param name="create">Autorise la création de la fenêtre si elle n’existe pas.</param>
        /// <returns>Fenêtre associée à cette session, ou <see langword="null"/> si aucune n’est disponible.</returns>
        private ModernEditorWindow EditorWindow(bool create) => MonacoWindow != null ? MonacoWindow(create) : session?.ModernEditor?.Invoke(create);
        /// <summary>Résout le module dans le résolveur injecté ou dans la session VBE.</summary>
        /// <param name="project">Clé du projet à résoudre.</param><param name="module">Nom du module exact.</param>
        /// <returns>Adaptateur du module dans l’éditeur moderne.</returns>
        private IEditorModule EditorModule(string project, string module) => MonacoModule != null ? MonacoModule(project, module) : session.ResolveEditorModule(project, module);

        /// <summary>Construit la définition JSON d’un outil Monaco avec les champs Project et Module.</summary>
        /// <param name="name">Nom de l’outil exposé au modèle.</param><param name="description">Comportement et limites de l’opération.</param>
        /// <param name="extra">Champs supplémentaires requis.</param>
        /// <returns>Définition JSON avec champs requis et propriétés de paramètres.</returns>
        private static object MonacoDefinition(string name, string description, params string[] extra)
        {
            var fields = new[] { "Project", "Module" }.Concat(extra).ToArray();
            return Definition(name, description, fields, fields);
        }
        /// <summary>Définitions des outils d’ouverture, lecture, navigation, édition et synchronisation Monaco.</summary>
        /// <value>Définitions JSON offertes au modèle.</value>
        private static object[] MonacoDefinitions => new[] {
            MonacoDefinition("monaco_open", "Open the exact live module in the modern Monaco editor. If loading, retry monaco_read. Never opens the active module by assumption."),
            MonacoDefinition("monaco_read", "Read Monaco draft, baseline, current native source, draft Version, normalized-LF NativeSha256, selection, conflict and automatic synchronization state. read_module continues to read native VBA only. This snapshot does not report compiler diagnostics or breakpoint state: use compile_project/debug_dialog/debug_state; project_symbols inspects native symbols, not unsynchronized drafts."),
            MonacoDefinition("monaco_navigate", "Select an exact range in the current Monaco draft using the Version returned by monaco_read. Does not edit VBA.", "ExpectedVersion", "StartLine", "StartColumn", "EndLine", "EndColumn"),
            MonacoDefinition("monaco_edit", "Replace the full Monaco draft using ExpectedVersion, then synchronize continuously to VBA. Preserves newer user typing by compare-and-swap. Check AppliedToDraft and Synchronized separately: a refused native write preserves the draft. Emits chat diff and guarded rollback. Agent mode and VBE edit policy required.", "ExpectedVersion", "Text"),
            MonacoDefinition("monaco_sync", "Synchronize exactly one draft to VBA using ExpectedVersion and ExpectedSha256=NativeSha256 from monaco_read (LF-normalized SHA, not read_module SHA). Refuses conflicts, running/break/protected projects. Emits chat diff with guarded rollback. Does not save the host document.", "ExpectedVersion", "ExpectedSha256")
        };

        /// <summary>Détermine si l’outil dépend de la synchronisation entre l’éditeur et le code natif.</summary>
        /// <param name="name">Nom de l’outil à classer.</param>
        /// <returns><see langword="true"/> pour les opérations incompatibles avec un brouillon non synchronisé.</returns>
        private static bool NeedsSynchronizedEditor(string name) => name == "compile_project" || !ReadOnlyTools.Contains(name);

        /// <summary>Refuse une mutation native lorsque le brouillon Monaco contient des changements non synchronisés.</summary>
        /// <param name="name">Nom de l’outil dont la mutation est envisagée.</param>
        private void GuardLegacyEditorMutation(string name)
        {
            if (!NeedsSynchronizedEditor(name) || name.StartsWith("monaco_", StringComparison.Ordinal)) return;
            var window = EditorWindow(false);
            if (window != null && !window.IsDisposed && window.HasPendingEditorDraft)
                throw new InvalidOperationException("A Monaco draft has unsynchronized changes. Use monaco_read and resolve/synchronize the draft before native/Git mutations. Drafts are preserved.");
        }

        /// <summary>Capture l’état Monaco, restaure les changements demandés puis réconcilie les documents ouverts.</summary>
        /// <param name="changes">Changements de code issus du transcript.</param><param name="hunk">Index facultatif du bloc à restaurer.</param>
        /// <returns>Réponse de restauration ou échec converti en réponse.</returns>
        internal async Task<Response> RestoreChangesAsync(CodeChange[] changes, int? hunk)
        {
            try
            {
                var window = EditorWindow(false);
                if (window != null && !window.IsDisposed && window.Documents.Any())
                    await window.CaptureForTool();
                GuardLegacyEditorMutation("replace_lines");
                var result = RestoreChanges(changes, hunk);
                if (result.Ok && window != null && !window.IsDisposed) await window.ProcessDocuments(false);
                return result;
            }
            catch (Exception error) { return Response.Failure(error.Message); }
        }

        /// <summary>Valide et exécute un outil Monaco, en appliquant les gardes de projet, d’édition et de thread UI.</summary>
        /// <param name="name">Nom exact de l’outil à appeler.</param><param name="arguments">Objet JSON contenant ses paramètres.</param>
        /// <returns>Réponse JSON de l’outil ou message d’échec sérialisé.</returns>
        private async Task<string> InvokeMonacoAsync(string name, string arguments)
        {
            try
            {
                var definition = MonacoDefinitions.Cast<dynamic>().FirstOrDefault(d => (string)d.function.name == name);
                if (definition == null) throw new ArgumentException("Unknown Monaco tool: " + name);
                var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                var fields = (Dictionary<string, object>)definition.function.parameters.properties;
                foreach (string field in fields.Keys)
                {
                    if (!values.TryGetValue(field, out var value) || value == null) throw new ArgumentException(field + " is required.");
                    bool integer = field == "ExpectedVersion" || field == "StartLine" || field == "StartColumn" || field == "EndLine" || field == "EndColumn";
                    if (integer ? !(value is int) || (int)value < 1 : !(value is string) || (field != "Text" && string.IsNullOrWhiteSpace((string)value)))
                        throw new ArgumentException("Invalid Monaco argument: " + field);
                }
                foreach (string field in values.Keys) if (!fields.ContainsKey(field)) throw new ArgumentException("Unexpected argument: " + field);
                string project = (string)values["Project"], module = (string)values["Module"];
                bool edit = name == "monaco_edit" || name == "monaco_sync";
                if (!edit) RequireProjectRead(project);
                else if (!string.IsNullOrEmpty(BoundProject) && !string.Equals(BoundProject, project, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Cette action vise un autre projet que celui de la conversation.");
                if (edit)
                {
                    if (settings.VbeEditApproval == "ReadOnly") throw new InvalidOperationException("VBE edits are disabled (Read-only mode).");
                    if (settings.VbeEditApproval != "Automatic" && settings.VbeEditApproval != "AskEachTime") throw new InvalidOperationException("Unknown VBE edit policy; action refused.");
                    if (settings.VbeEditApproval == "AskEachTime")
                        using (var approval = new VbeApprovalDialog(name + "\r\n" + project + " / " + module))
                            if (ShowApproval(approval, owner) != DialogResult.Yes) throw new InvalidOperationException("User rejected the edit.");
                }
                var target = EditorModule(project, module);
                var window = EditorWindow(name == "monaco_open");
                if (window == null || window.IsDisposed) throw new InvalidOperationException("Open the modern editor with monaco_open first.");
                if (window.InvokeRequired) throw new InvalidOperationException("Monaco tools must run on the owning VBE UI thread.");
                EditorDocument doc;
                if (name == "monaco_open")
                {
                    doc = await window.OpenModule(target);
                    if (!window.Ready) return json.Serialize(Response.Success(new { Loading = true, doc.Id, Next = "monaco_read" }));
                }
                else doc = window.FindToolDocument(target);
                object result;
                if (name == "monaco_navigate") result = await window.NavigateForTool(doc, (int)values["ExpectedVersion"], (int)values["StartLine"], (int)values["StartColumn"], (int)values["EndLine"], (int)values["EndColumn"]);
                else if (edit)
                {
                    var before = ReadCode(project, module);
                    Action synchronizedSource = () =>
                    {
                        // Capture the native diff before renderer/draft-worker awaits can admit
                        // unrelated user edits. A diff failure must not interrupt reconciliation.
                        try
                        {
                            var after = ReadCode(project, module);
                            if (before.Sha256 != after.Sha256)
                                CodeEdited?.Invoke(new CodeChange(project, module, before.Code, before.Sha256, after.Code, after.Sha256, CodeRollback.Lines(after.Code).Length));
                        }
                        catch (Exception error) { WriteLog("Monaco code diff readback failed: " + error.Message); }
                    };
                    result = name == "monaco_edit"
                        ? await window.EditForTool(doc, (int)values["ExpectedVersion"], (string)values["Text"], synchronizedSource)
                        : await window.SynchronizeForTool(doc, (int)values["ExpectedVersion"], (string)values["ExpectedSha256"], synchronizedSource);
                }
                else result = await window.ReadForTool(doc);
                return json.Serialize(Response.Success(result));
            }
            catch (Exception error) { return json.Serialize(Response.Failure(error.Message)); }
        }
    }
}
