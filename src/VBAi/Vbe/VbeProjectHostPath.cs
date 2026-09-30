using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace VBAi
{
    /// <summary>Reads the persisted host path without confusing Word VBA backing storage with its document.</summary>
    internal static class VbeProjectHostPath
    {
        /// <summary>Reads a project path on its owning thread; Word requires a unique PID-verified document identity.</summary>
        internal static string Read(object project, VbeProjectComponents.IOtherHostProbe native = null)
        {
            if (project == null) throw new ArgumentNullException(nameof(project));
            native = native ?? new VbeProjectComponents.NativeOtherHostProbe();
            string path;
            if (native.HostKind == "Word")
            {
                object document = VbeProjectComponents.MatchOtherHost(project, native);
                path = native.State(document).Path;
                if (native.ApplicationProcessId(native.Application()) != (uint)native.CurrentProcessId ||
                    !native.SameProject(project, native.DocumentProject(document)))
                    throw new InvalidOperationException("The Word document identity changed while reading its host path.");
            }
            else path = (string)((dynamic)project).FileName;
            if (string.IsNullOrWhiteSpace(path)) return null;
            if (!Path.IsPathRooted(path)) throw new InvalidOperationException("The host document path is not absolute.");
            return Path.GetFullPath(path);
        }

        /// <summary>Uses the explicit host path when present; older protocol fixtures retain their legacy FileName contract.</summary>
        internal static string FromFields(IDictionary<string, object> fields)
        {
            object path;
            return fields.TryGetValue("HostPath", out path) ||
                (AllowsLegacyPath && fields.TryGetValue("FileName", out path))
                ? Convert.ToString(path) : null;
        }

        /// <summary>Word never accepts a missing canonical-path field as permission to use its VBA backing file.</summary>
        internal static bool AllowsLegacyPath => new VbeProjectComponents.NativeOtherHostProbe().HostKind != "Word";

        /// <summary>Compares live COM identity, balancing only the IUnknown references acquired here.</summary>
        internal static bool SameProject(object first, object second)
        {
            if (first == null || second == null) return false;
            if (ReferenceEquals(first, second)) return true;
            if (!Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try { a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (a != IntPtr.Zero) Marshal.Release(a); if (b != IntPtr.Zero) Marshal.Release(b); }
        }
    }
}
