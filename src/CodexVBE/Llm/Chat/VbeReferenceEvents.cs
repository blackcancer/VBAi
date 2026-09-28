using System;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    // Subscribe and dispose on the VBE STA. Callbacks only invalidate managed state.
    /// <summary>Observe les ajouts et retraits de références du projet VBA courant.</summary>
    internal sealed class VbeReferenceEvents : IDisposable
    {
        /// <summary>IID de l’interface COM d’événements de références VBIDE.</summary>
        private static readonly Guid EventInterface = new Guid("0002E118-0000-0000-C000-000000000046");
        /// <summary>Opération d’abonnement COM.</summary>
        private readonly Action<object, Guid, int, Delegate> add;
        /// <summary>Opération de désabonnement COM.</summary>
        private readonly Action<object, Guid, int, Delegate> remove;
        /// <summary>Callback natif relayé vers l’invalidation du contexte géré.</summary>
        private readonly Action<object> handler;
        /// <summary>Projet COM actuellement observé.</summary>
        private object source;
        /// <summary>Indicateurs des deux abonnements et de la libération.</summary>
        private bool added, removed, disposed;

        /// <summary>Crée l’observateur et prépare les callbacks d’ajout et de retrait.</summary>
        /// <param name="changed">Action appelée pour invalider le contexte après un événement.</param>
        /// <param name="add">Opération d’abonnement facultative.</param>
        /// <param name="remove">Opération de désabonnement facultative.</param>
        internal VbeReferenceEvents(Action changed,
            Action<object, Guid, int, Delegate> add = null,
            Action<object, Guid, int, Delegate> remove = null)
        {
            this.add = add ?? new Action<object, Guid, int, Delegate>(ComEventsHelper.Combine);
            this.remove = remove ?? ((target, iid, id, callback) => ComEventsHelper.Remove(target, iid, id, callback));
            handler = item => { if (!disposed) changed(); };
        }

        /// <summary>Déplace l’abonnement vers le projet fourni; une source identique reste inchangée.</summary>
        /// <param name="next">Projet COM à observer, ou <see langword="null"/> pour arrêter l’observation.</param>
        internal void Observe(object next)
        {
            if (disposed) throw new ObjectDisposedException(nameof(VbeReferenceEvents));
            if (ReferenceEquals(source, next)) return;
            Detach();
            if (next == null) return;
            source = next;
            try
            {
                add(source, EventInterface, 1, handler); added = true;
                add(source, EventInterface, 2, handler); removed = true;
            }
            catch { Detach(); throw; }
        }

        /// <summary>Tente de retirer chaque abonnement actif avant d’effacer la source.</summary>
        private void Detach()
        {
            if (source == null) return;
            // Attempt both removals even when the project has already been closed.
            if (added) TryRemove(1);
            if (removed) TryRemove(2);
            added = removed = false;
            source = null;
        }
        /// <summary>Retire un identifiant d’événement sans interrompre le nettoyage en cas d’erreur COM.</summary>
        /// <param name="id">Identifiant du membre événementiel à retirer.</param>
        private void TryRemove(int id)
        {
            try { remove(source, EventInterface, id, handler); }
            catch (Exception error) { LoadLog.Write("Reference event detach: " + error.Message); }
        }
        /// <summary>Arrête définitivement l’observation et libère les abonnements COM.</summary>
        public void Dispose() { if (disposed) return; disposed = true; Detach(); }
    }
}
