using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Reads the persisted host path without confusing Word VBA backing storage with its document.</summary>
    internal static class VbeProjectHostPath
    {

        /// <summary>Resolves a persisted project path from the live host on the owning VBE thread.</summary>
        /// <param name="project">The live VBProject whose host document or macro path is required.</param>
        /// <param name="native">Host identity and filesystem probe; null selects the production probe.</param>
        /// <returns>The normalized absolute host path, or null for an unsaved project or unverified Outlook OTM storage.</returns>
        /// <remarks>Word uses a uniquely matched document and rechecks its process and COM project identity.
        /// Outlook's synthetic FileName is accepted only when it identifies an existing rooted .otm file.
        /// Other hosts retain VBProject.FileName; their file existence is not checked here.</remarks>
        /// <exception cref="ArgumentNullException">The project is null.</exception>
        /// <exception cref="InvalidOperationException">Word identity changes or a nonempty accepted path is relative.</exception>
        internal static string Read(object project, VbeProjectComponents.IOtherHostProbe native = null)
        {
            if (project == null) throw new ArgumentNullException(nameof(project));
            native = native ?? new VbeProjectComponents.NativeOtherHostProbe();
            string hostKind = native.HostKind;
            bool outlook;
            using (var process = Process.GetCurrentProcess())
                outlook = IsOutlookPathHost(hostKind, process.ProcessName);
            string path;
            if (hostKind == "Word")
            {
                object document = VbeProjectComponents.MatchOtherHost(project, native);
                path = native.State(document).Path;
                if (native.ApplicationProcessId(native.Application()) != (uint)native.CurrentProcessId ||
                    !native.SameProject(project, native.DocumentProject(document)))
                    throw new InvalidOperationException("The Word document identity changed while reading its host path.");
            }
            else path = (string)((dynamic)project).FileName;
            if (string.IsNullOrWhiteSpace(path)) return null;
            // Outlook may expose its localized project name as a nonexistent absolute
            // FileName under different caller directories. It is not persisted storage.
            if (outlook && (!Path.IsPathRooted(path) ||
                !string.Equals(Path.GetExtension(path), ".otm", StringComparison.OrdinalIgnoreCase) ||
                !native.FileExists(path))) return null;
            if (!Path.IsPathRooted(path)) throw new InvalidOperationException("The host document path is not absolute.");
            return Path.GetFullPath(path);
        }

        /// <summary>Recognizes Outlook storage without extending the separate document-adapter catalogue.</summary>
        /// <param name="documentHostKind">The adapter's host kind; null means no document adapter identified the host.</param>
        /// <param name="processName">The current process name without an extension, used only when the host kind is null.</param>
        /// <returns>True for the exact host kind Outlook, or for an unidentified host running in OUTLOOK (case insensitive).</returns>
        internal static bool IsOutlookPathHost(string documentHostKind, string processName) =>
            documentHostKind == "Outlook" || (documentHostKind == null &&
                string.Equals(processName, "OUTLOOK", StringComparison.OrdinalIgnoreCase));

        /// <summary>Uses the explicit host path when present; older protocol fixtures retain their legacy FileName contract.</summary>
        /// <param name="fields">Non-null project-status fields containing HostPath or a legacy FileName.</param>
        /// <returns>The selected field converted to text, or null when no permitted field is present.</returns>
        /// <remarks>A present HostPath wins even when null; legacy FileName fallback is never used in Word.</remarks>
        internal static string FromFields(IDictionary<string, object> fields)
        {
            object path;
            return fields.TryGetValue("HostPath", out path) ||
                (AllowsLegacyPath && fields.TryGetValue("FileName", out path))
                ? Convert.ToString(path) : null;
        }

        /// <summary>Word never accepts a missing canonical-path field as permission to use its VBA backing file.</summary>
        /// <value>False when the native probe identifies Word; true for other host kinds.</value>
        internal static bool AllowsLegacyPath => new VbeProjectComponents.NativeOtherHostProbe().HostKind != "Word";

        /// <summary>Compares live COM identity, balancing only the IUnknown references acquired here.</summary>
        /// <param name="first">The first live project object; null cannot match.</param>
        /// <param name="second">The second live project object; null cannot match.</param>
        /// <returns>True for the same managed reference or matching COM IUnknown identity; false for null or distinct non-COM objects.</returns>
        /// <remarks>Call on the thread permitted to access these COM objects. Temporary IUnknown references
        /// are released in finally; this method does not release either caller-owned RCW.</remarks>
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
