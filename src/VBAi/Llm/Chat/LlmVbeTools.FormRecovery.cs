using System;
namespace VBAi
{
    /// <summary>Expose les opérations du chat vers l’hôte VBA et les formulaires.</summary>
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Vérifie l’identité, le périmètre et l’état d’une coupe avant de construire sa requête de récupération.</summary>
        /// <param name="change">État de la coupe conservé dans le transcript.</param>
        /// <returns>Requête ciblant le formulaire et le conteneur d’origine.</returns>
        private Request FormRecoveryRequest(FormCutChange change)
        {
            if (change == null || !ReferenceEquals(change.Owner, this) || change.Attempted || change.Restored ||
                !string.Equals(change.Project, BoundProject, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This recovery is not available in the current live conversation scope.");
            return new Request { Project = change.Project, Form = change.Form, ParentPath = change.ParentPath, DesignerClipboardRecoveryId = change.RecoveryId };
        }
        /// <summary>Indique si l’hôte accepte actuellement la récupération de cette coupe.</summary>
        /// <param name="change">État de coupe à vérifier.</param>
        /// <returns><see langword="true"/> si les préconditions de récupération sont satisfaites.</returns>
        internal bool CanRecoverFormCut(FormCutChange change)
        {
            try { return CanRecoverDesignerCut(FormRecoveryRequest(change)); }
            catch { return false; }
        }
        /// <summary>Récupère les contrôles coupés après lecture des versions courantes de sélection et du presse-papiers.</summary>
        /// <param name="change">État de coupe associé au propriétaire et au conteneur d’origine.</param>
        /// <returns>Réponse de l’hôte avec le résultat ou l’erreur de récupération.</returns>
        internal Response RecoverFormCut(FormCutChange change)
        {
            try
            {
                ValidateScope?.Invoke();
                Request request = FormRecoveryRequest(change);
                request.Command = "form_clipboard_state";
                Response state = Execute(request);
                if (!state.Ok) return state;
                request.ExpectedDesignerSelectionVersion = (string)((dynamic)state.Data).SelectionVersion;
                request.ExpectedClipboardVersion = (string)((dynamic)state.Data).ClipboardVersion;
                request.Command = "recover_form_cut";
                Response result = Execute(request);
                if (result.Ok)
                {
                    change.Attempted = (bool)((dynamic)result.Data).RecoveryAttempted;
                    change.Restored = (bool)((dynamic)result.Data).RestoredNamesGeometryAndTabOrder;
                }
                return result;
            }
            catch (Exception ex) { return Response.Failure(ex.Message); }
        }
    }
}
