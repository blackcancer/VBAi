using System;

namespace VBAi
{

    /// <summary>Applique les commandes Undo et Redo natives au concepteur d’un formulaire.</summary>
    internal sealed partial class VbeForms
    {

        /// <summary>Exécute une action d’historique si l’arbre et l’état natif correspondent encore à la lecture fournie.</summary>
        /// <param name="request">Projet, formulaire, action et version d’arbre attendue.</param>
        /// <returns>Action exécutée, changements observés et erreurs de lecture avant/après.</returns>
        /// <exception cref="ArgumentException">L’action ou la version de précondition est absente/invalide.</exception>
        /// <exception cref="InvalidOperationException">L’arbre a changé ou l’action Undo/Redo n’est pas disponible.</exception>
        public object NativeHistory(Request request)
        {
            if (request.Action != "undo" && request.Action != "redo") throw new ArgumentException("Action must be undo or redo.");
            if (string.IsNullOrWhiteSpace(request.ExpectedTreeVersion)) throw new ArgumentException("ExpectedTreeVersion from form_tree is required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form hierarchy or native history state changed since it was read.");
            dynamic designer = form.Designer;
            bool available = request.Action == "undo" ? (bool)designer.CanUndo : (bool)designer.CanRedo;
            if (!available) throw new InvalidOperationException("The requested native Designer history action is unavailable.");
            string error = null; object after = null;
            try { if (request.Action == "undo") designer.UndoAction(); else designer.RedoAction(); }
            catch (Exception ex) { error = ex.Message; }
            try { after = Tree(request.Project, request.Form); }
            catch (Exception ex) { error = error ?? ex.Message; }
            var changes = after == null ? new FormHistoryDiff.Change[0] : FormHistoryDiff.Compare((object)before, after);
            return new
            {
                request.Project,
                request.Form,
                request.Action,
                Executed = error == null,
                Verified = error == null && changes.Length > 0,
                VerificationPending = changes.Length == 0,
                DesignerChanges = changes,
                ReadErrorsBefore = FormHistoryDiff.ReadErrorCount((object)before),
                ReadErrorsAfter = after == null ? (int?)null : FormHistoryDiff.ReadErrorCount(after),
                Tree = after,
                NativeError = error,
                Saved = false,
                NextRead = "form_tree",
                Limit = "Targets this Designer via UndoAction/RedoAction; does not enumerate its history. Availability flags alone do not prove a Designer change. Properties with read errors are excluded from differences; nonzero read-error counts mean incomplete inspection. Do not retry automatically if the result is pending."
            };
        }
    }
}
