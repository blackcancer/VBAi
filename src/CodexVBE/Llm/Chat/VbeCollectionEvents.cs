using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CodexVBE
{
    // All operations and callbacks run on the owning VBE STA. Never inspect a COM
    // object inside an event callback: removed objects may already be invalid.
    /// <summary>Observe les événements COM de collections de projets ou de composants du VBE.</summary>
    internal sealed class VbeCollectionEvents : IDisposable
    {
        /// <summary>Frontière native de combinaison des événements COM, conservée séparément des gardes de connexion.</summary>
        internal static Action<object, Guid, int, Delegate> CombineNative = ComEventsHelper.Combine;
        /// <summary>Identifiant de l’interface d’événements native choisie pour la collection.</summary>
        private readonly Guid iid;
        /// <summary>Opération qui attache un gestionnaire à un point de connexion.</summary>
        private readonly Action<object, Guid, int, Delegate> add, remove;
        /// <summary>Gestionnaires prévus pour la collection surveillée.</summary>
        private readonly List<Tuple<int, Delegate>> handlers = new List<Tuple<int, Delegate>>();
        /// <summary>Gestionnaires effectivement attachés et donc à détacher.</summary>
        private readonly List<Tuple<int, Delegate>> attached = new List<Tuple<int, Delegate>>();
        /// <summary>Collection COM actuellement observée.</summary>
        private object source;
        /// <summary>Indique que l’observateur a été libéré.</summary>
        private bool disposed;
        /// <summary>Construit les abonnements adaptés à une collection de projets ou de composants.</summary>
        /// <param name="changed">Callback d’invalidation appelé lors d’un événement pertinent.</param>
        /// <param name="components"><see langword="true"/> pour la collection des composants; sinon celle des projets.</param>
        /// <param name="add">Opération d’abonnement facultative, principalement destinée aux tests.</param>
        /// <param name="remove">Opération de désabonnement facultative, principalement destinée aux tests.</param>
        internal VbeCollectionEvents(Action changed, bool components,
            Action<object, Guid, int, Delegate> add = null,
            Action<object, Guid, int, Delegate> remove = null)
        {
            iid = new Guid(components ? "0002E116-0000-0000-C000-000000000046" : "0002E103-0000-0000-C000-000000000046");
            this.add = add ?? Connect;
            this.remove = remove ?? ((target, id, member, callback) => ComEventsHelper.Remove(target, id, member, callback));
            Action<object> item = value => { if (!disposed) changed(); };
            Action<object, string> renamed = (value, oldName) => { if (!disposed) changed(); };
            handlers.Add(Tuple.Create(1, (Delegate)item));
            handlers.Add(Tuple.Create(2, (Delegate)item));
            handlers.Add(Tuple.Create(3, (Delegate)renamed));
            // Reloaded for components, activated for projects. Selection/activation
            // of individual components does not invalidate the symbol inventory.
            handlers.Add(Tuple.Create(components ? 6 : 4, (Delegate)item));
        }
        /// <summary>Vérifie le point de connexion natif avant d’enregistrer le délégué COM.</summary>
        /// <param name="target">Collection COM source.</param><param name="iid">Interface d’événements attendue.</param>
        /// <param name="member">Identifiant du membre événementiel.</param><param name="handler">Délégué à connecter.</param>
        private static void Connect(object target, Guid iid, int member, Delegate handler)
        {
            var container = target as IConnectionPointContainer;
            if (container == null) throw new InvalidOperationException("No native collection event container.");
            IConnectionPoint point;
            container.FindConnectionPoint(ref iid, out point);
            if (point == null) throw new InvalidOperationException("The native event connection point is absent.");
            Guid actual; point.GetConnectionInterface(out actual);
            if (actual != iid) throw new InvalidOperationException("The native event connection interface does not match.");
            CombineNative(target, iid, member, handler);
        }
        /// <summary>Remplace la collection observée et attache ses gestionnaires, avec nettoyage si l’attachement échoue.</summary>
        /// <param name="next">Nouvelle source COM, ou <see langword="null"/> pour se désabonner.</param>
        internal void Observe(object next)
        {
            if (disposed) throw new ObjectDisposedException(nameof(VbeCollectionEvents));
            if (ReferenceEquals(source, next)) return;
            Detach();
            source = next;
            if (next == null) return;
            try
            {
                foreach (var entry in handlers)
                {
                    add(source, iid, entry.Item1, entry.Item2);
                    attached.Add(entry);
                }
            }
            catch { Detach(); throw; }
        }
        /// <summary>Détache tous les gestionnaires déjà connectés et libère la référence à la source.</summary>
        private void Detach()
        {
            foreach (var entry in attached)
            {
                try { remove(source, iid, entry.Item1, entry.Item2); }
                catch (Exception error) { LoadLog.Write("Collection event detach: " + error.Message); }
            }
            attached.Clear(); source = null;
        }
        /// <summary>Arrête l’observation et détache les gestionnaires au plus une fois.</summary>
        public void Dispose() { if (disposed) return; disposed = true; Detach(); }
    }
}
