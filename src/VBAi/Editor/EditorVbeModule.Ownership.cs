using System;
using System.Threading;

namespace VBAi
{
    /// <summary>Owns only explicit COM getter acquisitions; caller-supplied objects remain borrowed.</summary>
    internal sealed partial class EditorVbeModule
    {
        /// <summary>Acquisitions transferred by the adapter producer or by component replacement.</summary>
        private VbaTestComReferences ownedReferences;
        /// <summary>Acquisitions belonging to the retained native pane, separately from adapter bindings.</summary>
        private VbaTestComReferences windowReferences;
        /// <summary>Optional release boundary used by isolated ownership contracts.</summary>
        private readonly Action<object> releaseReference;
        /// <summary>Managed thread that owns this adapter's COM references.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        /// <summary>Records completed disposal without retrying native releases.</summary>
        private bool bindingsDisposed;
        /// <summary>Prevents automatic synchronization from replaying a replacement with failed recovery.</summary>
        private bool mutationOutcomeUncertain;

        /// <summary>Explains why a failed native restoration cannot be retried automatically.</summary>
        internal string MutationUnavailability => mutationOutcomeUncertain
            ? "Native component recovery is incomplete. Inspect the retained recovery export before another write." : null;

        /// <summary>Roots acquisitions after an uncertain native recovery and blocks later automatic writes.</summary>
        /// <param name="references">Acquisitions involved in the failed operation.</param>
        private void RetainForRecovery(VbaTestComReferences references)
        {
            mutationOutcomeUncertain = true;
            references.Retain();
            ownedReferences?.Retain();
            windowReferences?.Retain();
        }

        /// <summary>Transfers an explicit acquisition scope into a newly created adapter.</summary>
        /// <param name="vbe">Borrowed owning VBE.</param>
        /// <param name="project">Project acquired by the supplied scope.</param>
        /// <param name="component">Component acquired by the supplied scope.</param>
        /// <param name="references">Scope whose ownership transfers on successful construction.</param>
        /// <param name="release">Optional release boundary for subsequent getter acquisitions.</param>
        internal EditorVbeModule(object vbe, object project, object component,
            VbaTestComReferences references, Action<object> release = null) : this(vbe, project, component)
        {
            ownedReferences = references ?? throw new ArgumentNullException(nameof(references));
            releaseReference = release;
        }

        /// <summary>Creates a short-lived scope using the adapter's release boundary.</summary>
        /// <returns>A scope that must be disposed on the owning thread.</returns>
        private VbaTestComReferences NewReferences() => new VbaTestComReferences(releaseReference);

        /// <summary>Transfers only the project and component into a longer-lived adapter.</summary>
        /// <param name="vbe">Borrowed owning VBE.</param>
        /// <param name="project">One project acquisition tracked by the producer.</param>
        /// <param name="component">One component acquisition tracked by the producer.</param>
        /// <param name="producer">Scope retaining every other temporary getter acquisition.</param>
        /// <param name="release">Release boundary inherited by the resulting adapter.</param>
        /// <returns>An owned adapter whose bindings are independent of the producer's lifetime.</returns>
        internal static EditorVbeModule TransferBindings(object vbe, object project, object component,
            VbaTestComReferences producer, Action<object> release = null)
        {
            var bindings = new VbaTestComReferences(release);
            try
            {
                producer.Transfer(project);
                bindings.Add(project);
                producer.Transfer(component);
                bindings.Add(component);
                var module = new EditorVbeModule(vbe, project, component, bindings, release);
                bindings = null;
                return module;
            }
            finally { bindings?.Dispose(); }
        }

        /// <summary>Rejects disposed owned adapters and cross-thread native access.</summary>
        private void RequireReferenceOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("VBE module references must be used on their owning thread.");
            if (bindingsDisposed) throw new ObjectDisposedException(nameof(EditorVbeModule));
        }

        /// <summary>Releases the retained pane's getter acquisitions without closing or activating the pane.</summary>
        private void ReleaseWindowReferences()
        {
            RequireReferenceOwner();
            var references = windowReferences;
            windowReferences = null;
            nativeWindow = null;
            references?.Dispose();
        }

        /// <summary>Releases owned acquisitions once without closing the caller's native windows.</summary>
        /// <remarks>Borrowed constructor arguments are never added to the release scope.</remarks>
        public void Dispose()
        {
            if (bindingsDisposed) return;
            RequireReferenceOwner();
            // Disposal failure is terminal for this adapter; a release is never replayed.
            bindingsDisposed = true;
            var windows = windowReferences;
            var bindings = ownedReferences;
            windowReferences = null;
            ownedReferences = null;
            nativeWindow = null;
            try { windows?.Dispose(); }
            finally { bindings?.Dispose(); }
        }

        /// <summary>Adopts one known imported component acquisition after successful validation or recovery.</summary>
        /// <param name="replacement">Acquired imported component.</param>
        /// <param name="producer">Scope currently responsible for that acquisition.</param>
        private void AdoptReplacement(object replacement, VbaTestComReferences producer)
        {
            RequireReferenceOwner();
            if (ownedReferences == null) ownedReferences = NewReferences();
            producer.Transfer(replacement);
            ownedReferences.Add(replacement);
            component = replacement;
            ReleaseWindowReferences();
        }

        /// <summary>Captures an eligible active pane and transfers all getter acquisitions to its adapter.</summary>
        /// <param name="vbe">Borrowed VBE instance.</param>
        /// <param name="followOnly">Requires the exact active code window in unprotected design mode.</param>
        /// <param name="release">Optional release boundary for isolated ownership tests.</param>
        /// <returns>An owned adapter, or null after releasing all rejected acquisitions.</returns>
        internal static EditorVbeModule AcquireActive(object vbe, bool followOnly, Action<object> release = null)
        {
            var references = new VbaTestComReferences(release);
            try
            {
                dynamic host = vbe;
                dynamic activeWindow = followOnly ? references.Add((object)host.ActiveWindow) : null;
                if (followOnly && (activeWindow == null || (int)activeWindow.Type != 0)) return null;
                dynamic pane = references.Add((object)host.ActiveCodePane);
                if (pane == null) return null;
                dynamic code = references.Add((object)pane.CodeModule);
                dynamic component = references.Add((object)code.Parent);
                dynamic collection = references.Add((object)component.Collection);
                dynamic project = references.Add((object)collection.Parent);
                if (followOnly)
                {
                    int type = (int)component.Type;
                    if ((type != 1 && type != 2 && type != 3 && type != 100) ||
                        (int)project.Mode != 2 || (int)project.Protection != 0) return null;
                    object activeProject = references.Add((object)host.ActiveVBProject);
                    object paneWindow = references.Add((object)pane.Window);
                    if (activeProject == null || !SameIdentity(project, activeProject) ||
                        !SameIdentity((object)activeWindow, paneWindow)) return null;
                }
                var adapter = new EditorVbeModule(vbe, (object)project, (object)component, references, release);
                references = null;
                return adapter;
            }
            finally { references?.Dispose(); }
        }

        /// <summary>Acquires a selected component without borrowing the queue producer's temporary scope.</summary>
        /// <param name="vbe">Borrowed VBE instance.</param>
        /// <param name="release">Optional release boundary for isolated ownership tests.</param>
        /// <returns>An owned adapter for a code module, or null after releasing an ineligible selection.</returns>
        internal static EditorVbeModule AcquireSelected(object vbe, Action<object> release = null)
        {
            var references = new VbaTestComReferences(release);
            try
            {
                dynamic component = references.Add((object)((dynamic)vbe).SelectedVBComponent);
                if (component == null || !EditorProjectNavigation.IsCodeComponent((int)component.Type)) return null;
                dynamic collection = references.Add((object)component.Collection);
                object project = references.Add((object)collection.Parent);
                var adapter = new EditorVbeModule(vbe, project, (object)component, references, release);
                references = null;
                return adapter;
            }
            finally { references?.Dispose(); }
        }
    }
}
