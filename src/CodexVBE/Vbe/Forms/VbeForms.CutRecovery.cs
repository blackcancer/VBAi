using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        internal bool CanRecoverCut(Request request)
        {
            try
            {
                var recovery = clipboardRecoveries.FirstOrDefault(x => x.Id == request.DesignerClipboardRecoveryId);
                return recovery != null && recovery.CutTree != null && !recovery.RecoveryAttempted &&
                    ReferenceEquals((object)GetForm(GetDesignProject(request.Project), request.Form), recovery.Form) &&
                    (request.ParentPath ?? "") == recovery.ParentPath;
            }
            catch { return false; }
        }
        public object RecoverDesignerCut(Request request)
        {
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            var recovery = clipboardRecoveries.FirstOrDefault(x => x.Id == request.DesignerClipboardRecoveryId);
            if (recovery == null || recovery.CutTree == null || recovery.RecoveryAttempted)
                throw new InvalidOperationException("Cut recovery unavailable, expired or already attempted; inspect the current tree.");
            if (!ReferenceEquals((object)form, recovery.Form) || (request.ParentPath ?? "") != recovery.ParentPath)
                throw new InvalidOperationException("Clipboard recovery belongs to another live form or container.");
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            RequireClipboardRevision(request, (string)before.SelectionVersion, (string)before.ClipboardVersion);
            if (FormHistoryDiff.Compare(recovery.CutTree, (object)before.Tree).Length != 0)
                throw new InvalidOperationException("The form changed after the cut; recovery cannot overwrite intervening changes.");
            dynamic container = ClipboardContainer((object)form.Designer, (object)before.Tree, request.ParentPath);
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (dynamic control in container.Controls) names.Add((string)control.Name);
            if (recovery.Boxes.Any(box => names.Contains(box.Path)) || names.Count + recovery.Boxes.Length != recovery.TabOrder.Length)
                throw new InvalidOperationException("Cut controls are already present or the container changed; recovery would duplicate controls.");
            string error = null; object after = null; bool geometryVerified = false;
            try
            {
                WriteDesignerClipboard(recovery.Backup.CreateDataObject(), true);
                if (!recovery.Backup.Matches(ReadDesignerClipboard()))
                    throw new InvalidOperationException("Recovery clipboard readback differs; no paste attempted.");
                if (!(bool)container.CanPaste) throw new InvalidOperationException("The container cannot paste the recovery.");
                // Never repeat a paste after a partial failure.
                recovery.RecoveryAttempted = true;
                container.Paste();
                var controls = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (dynamic control in container.Controls) controls.Add((string)control.Name, (object)control);
                if (!new HashSet<string>(controls.Keys, StringComparer.Ordinal).SetEquals(recovery.TabOrder))
                    throw new InvalidOperationException("Recovered control names differ; inspect the resulting tree.");
                foreach (var box in recovery.Boxes) ApplyBox(controls[box.Path], box);
                for (int i = recovery.TabOrder.Length - 1; i >= 0; i--) ((dynamic)controls[recovery.TabOrder[i]]).TabIndex = 0;
                foreach (var box in recovery.Boxes) VerifyBox(controls[box.Path], box);
                for (int i = 0; i < recovery.TabOrder.Length; i++)
                    if ((int)((dynamic)controls[recovery.TabOrder[i]]).TabIndex != i) throw new InvalidOperationException("Recovered tab order differs.");
                geometryVerified = true;
            }
            catch (Exception ex) { error = ex.Message; }
            try { after = ClipboardState(request.Project, request.Form, request.ParentPath); }
            catch (Exception ex) { error = error ?? ex.Message; }
            object tree = after == null ? null : (object)((dynamic)after).Tree;
            return new { RestoredNamesGeometryAndTabOrder = geometryVerified && error == null, FullPropertyFidelityVerified = false,
                recovery.RecoveryAttempted, recovery.Backup.OmittedFormats, State = after, NativeError = error,
                RemainingDifferences = tree == null ? null : FormHistoryDiff.Compare(recovery.OriginalTree, tree),
                ReadErrorsBefore = FormHistoryDiff.ReadErrorCount(recovery.OriginalTree),
                ReadErrorsAfter = tree == null ? (int?)null : FormHistoryDiff.ReadErrorCount(tree), Saved = false,
                Limit = "Restores original direct control names, bounds and container tab order after this session's cut. Clipboard is replaced by the recovery. Other properties depend on native serialization; read errors and omitted formats preclude full fidelity claims. A paste attempt consumes automatic recovery even on failure; inspect the tree, never retry blindly." };
        }
    }
}
