using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Expose aux outils les opérations de lecture, navigation et modification des brouillons Monaco.</summary>
    internal sealed partial class ModernEditorWindow
    {
        /// <summary>Indique si au moins un document contient un brouillon modifié ou un conflit non résolu.</summary>
        /// <value><see langword="true"/> dès qu’un document est modifié ou en conflit.</value>
        internal bool HasPendingEditorDraft => documents.Values.Any(d => d.Dirty || d.Conflict);

        /// <summary>Capture les textes affichés par Monaco avant une opération d’outil et confirme que l’éditeur reste prêt.</summary>
        /// <returns>Tâche terminée lorsque la capture et la vérification de disponibilité sont achevées.</returns>
        /// <exception cref="InvalidOperationException">L’éditeur est occupé, fermé ou indisponible.</exception>
        internal async Task CaptureForTool()
        {
            if (busy || closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco is busy or unavailable. Retry after it is ready.");
            busy = true;
            try { await CaptureDocuments(); }
            finally { busy = false; }
            if (closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco became unavailable.");
        }

        /// <summary>Retrouve le document ouvert correspondant à l’instance de module ou à son identité COM.</summary>
        /// <param name="module">Module demandé par l’outil.</param>
        /// <returns>Document Monaco associé.</returns>
        /// <exception cref="InvalidOperationException">Le module exact n’est pas ouvert dans l’éditeur.</exception>
        internal EditorDocument FindToolDocument(IEditorModule module)
        {
            var doc = documents.Values.FirstOrDefault(d => ReferenceEquals(d.Module, module) ||
                (d.Module is EditorVbeModule vm && module is EditorVbeModule other && vm.IsComponent(other.Component)));
            if (doc == null) throw new InvalidOperationException("Open this exact module with monaco_open first.");
            return doc;
        }

        /// <summary>Capture puis lit le texte, la révision, la sélection et l’état natif d’un document Monaco.</summary>
        /// <param name="doc">Document à lire.</param>
        /// <returns>Objet de résultat avec brouillon, base, empreinte native, sélection et état de synchronisation.</returns>
        /// <exception cref="InvalidOperationException">Le document Monaco a été fermé pendant la lecture.</exception>
        internal async Task<object> ReadForTool(EditorDocument doc)
        {
            await CaptureForTool(); doc.Observe();
            busy = true;
            try
            {
                selected = doc.Id; SelectTab(doc.Id); await SelectEditorDocument(doc.Id);
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
                    AutomaticSynchronization = true, HostDocumentSaveInvoked = false, HostDocumentSaved = ReadDocumentHostSaved(doc) };
            }
            finally { busy = false; }
        }

        /// <summary>Sélectionne une plage dans la révision demandée puis renvoie l’état actualisé du document.</summary>
        /// <param name="doc">Document dans lequel naviguer.</param>
        /// <param name="version">Révision Monaco que l’appelant a lue.</param>
        /// <param name="startLine">Première ligne de la sélection, indexée à un.</param>
        /// <param name="startColumn">Première colonne, indexée à un.</param>
        /// <param name="endLine">Dernière ligne de la sélection, indexée à un.</param>
        /// <param name="endColumn">Colonne de fin, indexée à un.</param>
        /// <returns>État complet relu après le déplacement de sélection.</returns>
        /// <exception cref="ArgumentException">La plage ne tient pas dans le brouillon courant.</exception>
        /// <exception cref="InvalidOperationException">La révision a changé avant la navigation.</exception>
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

        /// <summary>Vérifie que l’éditeur et le document correspondent encore à la révision capturée.</summary>
        /// <param name="doc">Document concerné.</param>
        /// <param name="version">Révision attendue par l’appelant.</param>
        /// <exception cref="InvalidOperationException">L’éditeur est indisponible ou le brouillon a changé.</exception>
        private void RequireVersion(EditorDocument doc, int version)
        {
            if (closing || IsDisposed || !Ready) throw new InvalidOperationException("Monaco became unavailable.");
            if (!documents.ContainsKey(doc.Id) || versions[doc.Id] != version)
                throw new InvalidOperationException("The Monaco draft changed. Read monaco_read again before editing.");
        }

        /// <summary>Applique un texte à la révision lue puis tente sa synchronisation gardée avec le module natif.</summary>
        /// <param name="doc">Document à modifier.</param>
        /// <param name="version">Révision de brouillon attendue.</param>
        /// <param name="text">Nouveau texte complet à appliquer.</param>
        /// <param name="synchronizedSource">Rappel facultatif exécuté après une écriture native réussie.</param>
        /// <returns>Résultat indiquant l’application au brouillon, l’état de synchronisation, la révision et les erreurs récupérables.</returns>
        /// <exception cref="ArgumentException">Le texte dépasse les limites d’un document.</exception>
        /// <exception cref="InvalidOperationException">La version a changé, le module est protégé ou un conflit interdit l’édition.</exception>
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
                    AutomaticSynchronization = true, HostDocumentSaveInvoked = false, HostDocumentSaved = ReadDocumentHostSaved(doc) };
            }
            finally { busy = false; }
        }

        /// <summary>Synchronise le brouillon après vérification de la révision et de l’empreinte du code natif observé.</summary>
        /// <param name="doc">Document à synchroniser.</param>
        /// <param name="version">Révision Monaco attendue.</param>
        /// <param name="expectedNativeSha256">Empreinte SHA-256 native renvoyée lors de la lecture préalable.</param>
        /// <param name="synchronizedSource">Rappel facultatif exécuté après l’écriture native.</param>
        /// <returns>Résultat avec la nouvelle révision, le texte natif, son empreinte et l’état du brouillon.</returns>
        /// <exception cref="InvalidOperationException">Le texte natif ou la révision a changé pendant la préparation.</exception>
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
                    NativeSha256 = EditorDocument.Hash(doc.Native), HostDocumentSaveInvoked = false, HostDocumentSaved = ReadDocumentHostSaved(doc) };
            }
            finally { busy = false; }
        }
    }
}
