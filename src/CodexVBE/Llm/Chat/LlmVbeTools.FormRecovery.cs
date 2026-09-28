using System;
namespace CodexVBE
{
    internal sealed partial class LlmVbeTools
    {
        private Request FormRecoveryRequest(FormCutChange change)
        {
            if (change == null || !ReferenceEquals(change.Owner, this) || change.Attempted || change.Restored ||
                !string.Equals(change.Project, BoundProject, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("This recovery is not available in the current live conversation scope.");
            return new Request { Project = change.Project, Form = change.Form, ParentPath = change.ParentPath, DesignerClipboardRecoveryId = change.RecoveryId };
        }
        internal bool CanRecoverFormCut(FormCutChange change)
        {
            try { return session.CanRecoverFormCut(FormRecoveryRequest(change)); }
            catch { return false; }
        }
        internal Response RecoverFormCut(FormCutChange change)
        {
            try
            {
                ValidateScope?.Invoke();
                Request request = FormRecoveryRequest(change);
                request.Command = "form_clipboard_state";
                Response state = session.Execute(request);
                if (!state.Ok) return state;
                request.ExpectedDesignerSelectionVersion = (string)((dynamic)state.Data).SelectionVersion;
                request.ExpectedClipboardVersion = (string)((dynamic)state.Data).ClipboardVersion;
                request.Command = "recover_form_cut";
                Response result = session.Execute(request);
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
