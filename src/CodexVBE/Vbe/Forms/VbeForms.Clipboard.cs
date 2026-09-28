using System;
using System.Collections.Generic;
using System.Globalization;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        private sealed class ClipboardRecovery
        {
            public string Id, ParentPath; public object Form; public DesignerClipboardBackup Backup;
            public object OriginalTree, CutTree; public FormLayoutBox[] Boxes; public string[] TabOrder;
            public bool RecoveryAttempted;
        }
        private readonly Queue<ClipboardRecovery> clipboardRecoveries = new Queue<ClipboardRecovery>();
        public object RestoreDesignerClipboard(Request request)
        {
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            var recovery = clipboardRecoveries.FirstOrDefault(x => x.Id == request.DesignerClipboardRecoveryId);
            if (recovery == null) throw new InvalidOperationException("Clipboard recovery not found or expired; only the latest 8 cuts in this session are retained.");
            if (!ReferenceEquals((object)form, recovery.Form) || (request.ParentPath ?? "") != recovery.ParentPath)
                throw new InvalidOperationException("Clipboard recovery belongs to another live form or container.");
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            RequireClipboardRevision(request, (string)before.SelectionVersion, (string)before.ClipboardVersion);
            System.Windows.Forms.Clipboard.SetDataObject(recovery.Backup.CreateDataObject(), true);
            bool verified = recovery.Backup.Matches(System.Windows.Forms.Clipboard.GetDataObject());
            return new { Restored = verified, recovery.Backup.OmittedFormats,
                State = ClipboardState(request.Project, request.Form, request.ParentPath), NextAction = "native_form_clipboard paste",
                Limit = "Restores captured clipboard formats only; does not recreate controls. Inspect the current tree before an explicit paste to avoid duplicates. Native paste may change placement; omitted formats are not restored." };
        }
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        public object ClipboardState(string projectName, string formName, string parentPath = null)
        {
            dynamic form = GetForm(GetDesignProject(projectName), formName);
            dynamic tree = Tree(projectName, formName);
            dynamic container = ClipboardContainer((object)form.Designer, (object)tree, parentPath);
            var selected = new List<string>();
            foreach (dynamic control in container.Selected)
            {
                if (selected.Count >= 256) throw new InvalidOperationException("Designer selection exceeds 256 controls.");
                selected.Add((string)control.Name);
            }
            uint sequence = GetClipboardSequenceNumber();
            bool canPaste = (bool)container.CanPaste;
            if (sequence != GetClipboardSequenceNumber()) throw new InvalidOperationException("Clipboard changed during inspection.");
            string selectionVersion = VbeCodeClipboard.Hash(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(
                new { Project = projectName, Form = formName, ParentPath = parentPath ?? "", TreeVersion = (string)tree.TreeVersion, Selected = selected }));
            return new { Project = projectName, Form = formName, ParentPath = parentPath ?? "", Tree = (object)tree, Selected = selected,
                SelectionVersion = selectionVersion, ClipboardVersion = sequence.ToString(CultureInfo.InvariantCulture), CanPaste = canPaste,
                Limit = "Current Designer selection; binary clipboard data is not read or sent to the model. Clipboard revision is a Windows sequence number, not a content hash." };
        }
        public object NativeClipboard(Request request)
        {
            if (request.Action != "copy" && request.Action != "cut" && request.Action != "paste")
                throw new ArgumentException("Action must be copy, cut or paste.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            RequireClipboardRevision(request, (string)before.SelectionVersion, (string)before.ClipboardVersion);
            if (request.Action != "paste" && ((List<string>)before.Selected).Count == 0)
                throw new InvalidOperationException("Select Designer controls before copying or cutting.");
            if (request.Action == "paste" && !(bool)before.CanPaste)
                throw new InvalidOperationException("The Designer cannot paste the current clipboard.");
            dynamic container = ClipboardContainer((object)form.Designer, (object)before.Tree, request.ParentPath);
            string error = null; object after = null; ClipboardRecovery recovery = null;
            try
            {
                if (request.Action == "copy") container.Copy();
                else if (request.Action == "cut")
                {
                    // Capture an eager copy before removal; no delayed COM clipboard object is retained.
                    var boxes = new List<FormLayoutBox>(); var tabs = new SortedDictionary<int, string>();
                    foreach (dynamic control in container.Controls)
                    {
                        string name = (string)control.Name; tabs.Add((int)control.TabIndex, name);
                        if (((List<string>)before.Selected).Contains(name)) boxes.Add(new FormLayoutBox { Path = name,
                            Left = (double)control.Left, Top = (double)control.Top, Width = (double)control.Width, Height = (double)control.Height });
                    }
                    container.Copy();
                    uint sequence = GetClipboardSequenceNumber();
                    var backup = DesignerClipboardBackup.Capture(System.Windows.Forms.Clipboard.GetDataObject());
                    if (sequence != GetClipboardSequenceNumber()) throw new InvalidOperationException("Clipboard changed while preparing recovery; cut was not attempted.");
                    recovery = new ClipboardRecovery { Id = Guid.NewGuid().ToString("N"), Form = (object)form, ParentPath = request.ParentPath ?? "", Backup = backup, OriginalTree = (object)before.Tree, Boxes = boxes.ToArray(), TabOrder = tabs.Values.ToArray() };
                    if (clipboardRecoveries.Count == 8) clipboardRecoveries.Dequeue();
                    clipboardRecoveries.Enqueue(recovery);
                    container.Cut();
                }
                else container.Paste();
            }
            catch (Exception ex) { error = ex.Message; }
            try { after = ClipboardState(request.Project, request.Form, request.ParentPath); }
            catch (Exception ex) { error = error ?? ex.Message; }
            object afterTree = after == null ? null : (object)((dynamic)after).Tree;
            var changes = afterTree == null ? new FormHistoryDiff.Change[0] : FormHistoryDiff.Compare((object)before.Tree, afterTree);
            if (recovery != null && error == null && afterTree != null && changes.Length > 0) recovery.CutTree = afterTree;
            bool clipboardChanged = after != null && (string)((dynamic)after).ClipboardVersion != (string)before.ClipboardVersion;
            return new { request.Project, request.Form, request.Action, Executed = error == null,
                DesignerChangeObserved = changes.Length > 0, ClipboardChanged = clipboardChanged,
                DesignerClipboardRecoveryId = recovery?.Id, RecoveryOmittedFormats = recovery?.Backup.OmittedFormats, RecoveryBytes = recovery?.Backup.ByteCount,
                ClipboardContentVerified = false, Before = (object)before, After = after, DesignerChanges = changes,
                ReadErrorsBefore = FormHistoryDiff.ReadErrorCount((object)before.Tree),
                ReadErrorsAfter = afterTree == null ? (int?)null : FormHistoryDiff.ReadErrorCount(afterTree),
                NativeError = error, Saved = false, NextRead = "form_clipboard_state",
                Limit = "Native selected controls are copied/cut; paste uses this Designer's current selection context. Event-handler code is not transferred. Binary clipboard contents are not inspected. Partial changes are reported without retry or implicit rollback; Copy/Cut/Paste may not enter native undo history; inspect CanUndo before offering native_form_history. Cut stores up to 8 MiB of readable MSForms formats before removal (latest 8 cuts in this session). restore_form_clipboard can republish the recovery; it does not undo or restore geometry automatically." };
        }
        private static object ClipboardContainer(object designer, object tree, string parentPath)
        {
            if (string.IsNullOrEmpty(parentPath)) return designer;
            if (!TreeContainsPath((IEnumerable)((dynamic)tree).Controls, parentPath))
                throw new ArgumentException("ParentPath is not canonical in form_tree.");
            // Validate it exposes a child Controls collection, excluding leaf controls.
            ResolveNestedControls((dynamic)designer, parentPath);
            return ResolveTreeItem(designer, parentPath);
        }
        public object SelectDesignerControls(Request request)
        {
            if (request.Items == null || request.Items.Length > 64 ||
                request.Items.Any(string.IsNullOrWhiteSpace) || request.Items.Distinct(StringComparer.Ordinal).Count() != request.Items.Length)
                throw new ArgumentException("Items must contain 0 to 64 unique direct control names; empty clears this container's selection.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            if (string.IsNullOrWhiteSpace(request.ExpectedDesignerSelectionVersion) ||
                !string.Equals(request.ExpectedDesignerSelectionVersion, (string)before.SelectionVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Designer tree or selection changed; read form_clipboard_state again.");
            dynamic container = ClipboardContainer((object)form.Designer, (object)before.Tree, request.ParentPath);
            var controls = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (dynamic control in container.Controls) controls.Add((string)control.Name, (object)control);
            if (request.Items.Any(name => !controls.ContainsKey(name)))
                throw new ArgumentException("Items contains a name outside the selected container; no selection changed.");
            string error = null; object after = null;
            try
            {
                foreach (var entry in controls) ((dynamic)entry.Value).InSelection = request.Items.Contains(entry.Key, StringComparer.Ordinal);
            }
            catch (Exception ex) { error = ex.Message; }
            try { after = ClipboardState(request.Project, request.Form, request.ParentPath); }
            catch (Exception ex) { error = error ?? ex.Message; }
            bool verified = after != null && error == null &&
                new HashSet<string>((List<string>)((dynamic)after).Selected, StringComparer.Ordinal).SetEquals(request.Items);
            return new { Verified = verified, Before = (object)before, After = after, NativeError = error,
                NextRead = "form_clipboard_state", Limit = "Changes only the specified container selection; selections in other containers are independent. No focus or clipboard mutation. On partial failure re-read before another action." };
        }
        internal static void RequireClipboardRevision(Request request, string selectionVersion, string clipboardVersion)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedDesignerSelectionVersion) ||
                !string.Equals(request.ExpectedDesignerSelectionVersion, selectionVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Designer tree or selection changed; read form_clipboard_state again.");
            if (string.IsNullOrWhiteSpace(request.ExpectedClipboardVersion) || request.ExpectedClipboardVersion != clipboardVersion || clipboardVersion == "0")
                throw new InvalidOperationException("Clipboard changed or is unavailable; read form_clipboard_state again.");
        }
    }
}
