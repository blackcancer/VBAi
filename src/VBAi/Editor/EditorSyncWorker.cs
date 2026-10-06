using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Diff immuable calculé en arrière-plan pour synchroniser une capture du document.</summary>
    internal sealed class EditorSyncPlan
    {

        /// <summary>Source native attendue et texte de brouillon capturé.</summary>
        internal readonly string Before, After;

        /// <summary>Remplacement de lignes calculé entre les deux textes.</summary>
        internal readonly Tuple<int, int, string> Patch;

        /// <summary>Identifiant du thread ayant préparé le plan.</summary>
        internal readonly int WorkerThreadId;

        /// <summary>Valide le texte capturé et calcule son remplacement minimal.</summary>
        /// <param name="before">Source native de base.</param>
        /// <param name="after">Texte de brouillon à synchroniser.</param>
        internal EditorSyncPlan(string before, string after)
        {
            Before = before; After = after; EditorDocument.Validate(after);
            Patch = EditorDocument.Difference(before, after);
            WorkerThreadId = Thread.CurrentThread.ManagedThreadId;
        }
    }

    /// <summary>One background worker per editor. Only immutable text snapshots cross this boundary.</summary>
    internal sealed class EditorSyncWorker : IDisposable
    {

        /// <summary>File des tâches traitées par le thread de synchronisation.</summary>
        private readonly BlockingCollection<Action> work = new BlockingCollection<Action>();

        /// <summary>Thread de fond qui consomme la file jusqu’à sa fermeture.</summary>
        private readonly Thread thread;

        /// <summary>Démarre le consommateur de tâches propre à cet éditeur.</summary>
        internal EditorSyncWorker() : this("VBAi editor synchronization") { }

        /// <summary>Names isolated queues so profiler traces distinguish language and synchronization.</summary>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        internal EditorSyncWorker(string name)
        {
            thread = new Thread(Run) { IsBackground = true, Name = name };
            thread.Start();
        }

        /// <summary>Exécute les actions de la file jusqu’à la fin de son alimentation.</summary>
        private void Run()
        {
            try { foreach (var action in work.GetConsumingEnumerable()) action(); }
            finally { work.Dispose(); }
        }

        /// <summary>Exécute une fonction sur le thread de fond et renvoie sa tâche de résultat.</summary>
        /// <typeparam name="T">Type du résultat calculé.</typeparam>
        /// <param name="action">Opération autonome ne dépendant pas des objets UI ou COM.</param>
        /// <returns>Tâche terminée avec le résultat ou l’exception de l’opération.</returns>
        internal Task<T> Evaluate<T>(Func<T> action)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            try { work.Add(() => { try { completion.SetResult(action()); } catch (Exception error) { completion.SetException(error); } }); }
            catch (Exception error) { completion.SetException(error); }
            return completion.Task;
        }

        /// <summary>Capture le document sur le thread appelant, sauvegarde le brouillon puis prépare son plan en arrière-plan.</summary>
        /// <param name="document">Document à capturer sans le transmettre au thread de fond.</param>
        /// <param name="drafts">Magasin qui persiste l’instantané chiffré.</param>
        /// <returns>Tâche contenant le diff prêt pour la synchronisation native.</returns>
        internal Task<EditorSyncPlan> Prepare(EditorDocument document, EditorDraftStore drafts)
        {
            // Capture on the UI thread. The worker never dereferences a document, COM proxy or control.
            var snapshot = new EditorDraft { Key = document.RecoveryKey, Baseline = document.Baseline, Text = document.Text };
            string id = document.Id;
            var completion = new TaskCompletionSource<EditorSyncPlan>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                work.Add(() =>
                {
                    try
                    {
                        drafts.SaveSnapshot(id, snapshot);
                        completion.SetResult(new EditorSyncPlan(snapshot.Baseline, snapshot.Text));
                    }
                    catch (Exception error) { completion.SetException(error); }
                });
            }
            catch (Exception error) { completion.SetException(error); }
            return completion.Task;
        }

        /// <summary>Ferme la file sans bloquer le thread UI et laisse finir les instantanés déjà en attente.</summary>
        public void Dispose()
        {
            // Never wait on a background job from the host UI. Queued snapshots finish safely.
            try { work.CompleteAdding(); } catch (ObjectDisposedException) { }
        }
    }
}
