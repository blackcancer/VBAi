using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed partial class ModernEditorWindow
    {
        internal bool HasPendingEditorDraft => documents.Values.Any(d => d.Dirty || d.Conflict);

        internal async Task CaptureForTool()
        {
            if (busy || closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco is busy or unavailable. Retry after it is ready.");
            busy = true;
            try { await CaptureDocuments(); }
            finally { busy = false; }
            if (closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco became unavailable.");
        }

        internal EditorDocument FindToolDocument(IEditorModule module)
        {
            var doc = documents.Values.FirstOrDefault(d => ReferenceEquals(d.Module, module) ||
                (d.Module is EditorVbeModule vm && module is EditorVbeModule other && vm.IsComponent(other.Component)));
            if (doc == null) throw new InvalidOperationException("Open this exact module with monaco_open first.");
            return doc;
        }

        internal async Task<object> ReadForTool(EditorDocument doc)
        {
            await CaptureForTool(); doc.Observe();
            busy = true;
            try
            {
                selected = doc.Id; SelectTab(doc.Id); await Script("select", doc.Id);
                // Chromium captures text, selection and revision together, without another UI await.
                var snapshot = json.Deserialize<Dictionary<string, object>>(await Script("read", doc.Id));
                if (snapshot == null) throw new InvalidOperationException("The Monaco document was closed.");
                string text = (string)snapshot["text"]; int version = (int)snapshot["version"];
                if (version >= versions[doc.Id]) { doc.Edit(text); versions[doc.Id] = version; }
                bool dirty = text != doc.Baseline;
                return new { DocumentId = doc.Id, Version = version, Draft = text, Baseline = doc.Baseline,
                    Native = doc.Native, NativeSha256 = EditorDocument.Hash(doc.Native), Dirty = dirty,
                    Conflict = dirty && doc.Native != doc.Baseline && doc.Native != text,
                    Writable = doc.Writable, Selection = snapshot["selection"],
                    AutomaticSynchronization = true, HostDocumentSaved = false };
            }
            finally { busy = false; }
        }

        internal async Task<object> NavigateForTool(EditorDocument doc, int version, int startLine, int startColumn, int endLine, int endColumn)
        {
            await CaptureForTool(); RequireVersion(doc, version);
            var lines = doc.Text.Split('\n');
            if (startLine < 1 || endLine < startLine || endLine > lines.Length || startColumn < 1 || endColumn < 1 ||
                startColumn > lines[startLine - 1].Length + 1 || endColumn > lines[endLine - 1].Length + 1 ||
                (startLine == endLine && endColumn < startColumn)) throw new ArgumentException("Selection is outside the current draft.");
            selected = doc.Id; SelectTab(doc.Id);
            if (await Script("selectRange", doc.Id, version, startLine, startColumn, endLine, endColumn) != "true")
                throw new InvalidOperationException("The Monaco draft changed before navigation. Read it again.");
            return await ReadForTool(doc);
        }

        private void RequireVersion(EditorDocument doc, int version)
        {
            if (closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco became unavailable.");
            if (!documents.ContainsKey(doc.Id) || versions[doc.Id] != version)
                throw new InvalidOperationException("The Monaco draft changed. Read monaco_read again before editing.");
        }

        internal async Task<object> EditForTool(EditorDocument doc, int version, string text, Action synchronizedSource = null)
        {
            EditorDocument.Validate(text);
            await CaptureForTool(); RequireVersion(doc, version);
            if (!doc.Writable) throw new InvalidOperationException("VBA is running, paused, protected or unavailable.");
            doc.Observe();
            if (doc.Conflict || doc.Native != doc.Baseline) throw new InvalidOperationException("The native module and draft diverged. Wait for synchronization or resolve the conflict before editing.");
            busy = true;
            try
            {
                int applied;
                if (!int.TryParse(await Script("apply", doc.Id, version, EditorDocument.Normalize(text)), out applied) || applied <= 0)
                    throw new InvalidOperationException("The Monaco draft changed during the edit. Nothing was replaced.");
                if (versions[doc.Id] <= applied) { doc.Edit(text); versions[doc.Id] = applied; }
                await CaptureDocuments();
                var plan = await PrepareSynchronization(doc);
                await CaptureDocuments();
                if (versions[doc.Id] != applied)
                {
                    await PrepareSynchronization(doc);
                    return new { AppliedToDraft = true, Synchronized = false, Version = versions[doc.Id], doc.Dirty,
                        Reason = "Newer user typing was preserved. Continuous synchronization will handle it separately." };
                }
                RequireVersion(doc, applied);
                if (doc.Text != plan.After || doc.Baseline != plan.Before)
                    throw new InvalidOperationException("The Monaco draft changed during synchronization preparation.");
                string captured = doc.Text;
                string synchronized;
                try { synchronized = doc.Synchronize(plan); }
                catch (Exception error) { SetStatus(); return new { AppliedToDraft = true, Synchronized = false, Version = versions[doc.Id], doc.Dirty, Error = error.Message }; }
                synchronizedSource?.Invoke();
                int reconciled;
                // COM can pump UI messages: never replace text against a newer received revision.
                if (int.TryParse(await Script("apply", doc.Id, applied, synchronized), out reconciled) && reconciled > 0)
                { doc.Acknowledge(synchronized, captured); versions[doc.Id] = Math.Max(versions[doc.Id], reconciled); }
                await CaptureDocuments(); if (doc.Dirty) await PrepareSynchronization(doc); else Drafts.ClearOwn(doc); SetStatus();
                return new { AppliedToDraft = true, Synchronized = true, AppliedVersion = applied, Version = versions[doc.Id], doc.Dirty,
                    AutomaticSynchronization = true, HostDocumentSaved = false };
            }
            finally { busy = false; }
        }

        internal async Task<object> SynchronizeForTool(EditorDocument doc, int version, string expectedNativeSha256, Action synchronizedSource = null)
        {
            await CaptureForTool(); RequireVersion(doc, version); doc.Observe();
            if (!string.Equals(EditorDocument.Hash(doc.Native), expectedNativeSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The native module changed. Read monaco_read and resolve the conflict first.");
            busy = true;
            try
            {
                var plan = await PrepareSynchronization(doc);
                await CaptureDocuments();
                RequireVersion(doc, version);
                if (doc.Text != plan.After || doc.Baseline != plan.Before)
                    throw new InvalidOperationException("The Monaco draft changed during synchronization preparation.");
                doc.Observe();
                if (!string.Equals(EditorDocument.Hash(doc.Native), expectedNativeSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The native module changed during synchronization preparation.");
                string captured = doc.Text;
                string actual = doc.Synchronize(plan);
                synchronizedSource?.Invoke();
                int applied;
                if (int.TryParse(await Script("apply", doc.Id, version, actual), out applied) && applied > 0)
                { doc.Acknowledge(actual, captured); versions[doc.Id] = Math.Max(versions[doc.Id], applied); }
                await CaptureDocuments();
                if (doc.Dirty) await PrepareSynchronization(doc); else Drafts.ClearOwn(doc);
                SetStatus();
                return new { Synchronized = true, Version = versions[doc.Id], doc.Dirty, Native = doc.Native,
                    NativeSha256 = EditorDocument.Hash(doc.Native), HostDocumentSaved = false };
            }
            finally { busy = false; }
        }
    }
}
