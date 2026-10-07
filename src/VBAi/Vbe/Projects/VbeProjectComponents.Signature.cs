using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Captures owner thread, live project identity and stable content before deferred signing.</summary>
        internal Action CaptureSignaturePersistence(string projectName)
        {
            object original = GetProject(projectName);
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            string revision = SignaturePersistenceRevision(projectName);
            return () =>
            {
                if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                    throw new InvalidOperationException("Signature persistence must run on its original VBE thread.");
                object current = GetProject(projectName);
                if (!SameSignatureProject(original, current))
                    throw new InvalidOperationException("The original project identity changed during signing.");
                if (!string.Equals(revision, SignaturePersistenceRevision(projectName), StringComparison.Ordinal))
                    throw new InvalidOperationException("The project content, path or mode changed during signing.");
            };
        }

        /// <summary>Ignores only the Saved flag that assigning a signature is expected to change.</summary>
        private string SignaturePersistenceRevision(string projectName)
        {
            dynamic project = GetProject(projectName);
            dynamic state = ProjectProperties(projectName);
            var content = new List<object>();
            foreach (dynamic component in project.VBComponents)
            {
                dynamic snapshot = ComponentSnapshot(projectName, component);
                content.Add(new { snapshot.Component, snapshot.Type, snapshot.CodeSha256, snapshot.FormVersion,
                    Properties = ((List<VbePropertyInfo>)snapshot.Properties).Where(p => !string.Equals(p.Name, "Saved", StringComparison.OrdinalIgnoreCase)).ToArray(),
                    snapshot.DesignerProperties, snapshot.HostProperties });
            }
            return Hash(json.Serialize(new { Mode = (int)project.Mode,
                Properties = ((List<VbePropertyInfo>)state.Properties).Where(p => !string.Equals(p.Name, "Saved", StringComparison.OrdinalIgnoreCase)).ToArray(),
                state.References, Content = content }));
        }

        /// <summary>Compares managed fixtures by reference and native projects by IUnknown identity.</summary>
        private static bool SameSignatureProject(object original, object current)
        {
            if (ReferenceEquals(original, current)) return true;
            if (original == null || current == null || !Marshal.IsComObject(original) || !Marshal.IsComObject(current)) return false;
            IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
            try
            {
                first = Marshal.GetIUnknownForObject(original);
                second = Marshal.GetIUnknownForObject(current);
                return first == second;
            }
            finally
            {
                if (first != IntPtr.Zero) Marshal.Release(first);
                if (second != IntPtr.Zero) Marshal.Release(second);
            }
        }
    }
}
