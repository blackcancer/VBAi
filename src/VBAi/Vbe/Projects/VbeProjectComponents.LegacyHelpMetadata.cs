using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Marks legacy Access/Publisher HelpFile and HelpContextID setter paths as unsupported.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Current executable name reader used to restrict this compatibility guard to Access/Publisher.</summary>
        internal Func<string> LegacyHelpMetadataProcessName = () => Process.GetCurrentProcess().ProcessName;

        /// <summary>Predicate requiring an actual native COM VBProject.</summary>
        internal Func<object, bool> LegacyHelpMetadataNativeProject = Marshal.IsComObject;

        /// <summary>Status reported for legacy COM setters on restricted Access/Publisher Type100 projects.</summary>
        internal const string LegacyHelpMetadataSetterStatus = "HostLegacyWriteUnsupported";

        /// <summary>Explicit asynchronous native General read/write route offered instead of automatic setter fallback.</summary>
        internal const string LegacyHelpMetadataAlternative = "Use read_project_general followed by an explicitly approved asynchronous set_project_general; no automatic fallback or Save is performed.";

        /// <summary>Checks whether a property belongs to the legacy help metadata pair.</summary>
        /// <param name="property">COM property name.</param>
        /// <returns>True only for HelpFile or HelpContextID, compared case-insensitively.</returns>
        internal static bool IsLegacyHelpMetadataProperty(string property) =>
            string.Equals(property, "HelpFile", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(property, "HelpContextID", StringComparison.OrdinalIgnoreCase);

        /// <summary>Checks whether the current process name is Access or Publisher.</summary>
        /// <param name="processName">Executable process name, without requiring a particular letter casing.</param>
        /// <returns>True for MSACCESS or MSPUB only.</returns>
        internal static bool IsLegacyHelpMetadataHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

        /// <summary>Requires both an Access/Publisher host process and a native Type100 host project.</summary>
        /// <param name="project">Candidate VBProject whose host-project type is read from COM.</param>
        /// <returns>True for a native Type100 project in Access/Publisher; a failed Type read propagates.</returns>
        private bool IsRestrictedLegacyHelpProject(object project)
        {
            if (!IsLegacyHelpMetadataHost(LegacyHelpMetadataProcessName()) || !LegacyHelpMetadataNativeProject(project)) return false;
            // A failed Type read propagates. It cannot admit an unverified native project to a setter.
            return Convert.ToInt32(((dynamic)project).Type, CultureInfo.InvariantCulture) == 100;
        }

        /// <summary>Rejects legacy COM writes to restricted help metadata before any setter is invoked.</summary>
        /// <param name="project">Candidate VBProject being modified.</param>
        /// <param name="property">Requested property name.</param>
        private void RequireLegacyHelpMetadataWriteSupported(object project, string property)
        {
            if (IsLegacyHelpMetadataProperty(property) && IsRestrictedLegacyHelpProject(project))
                throw new InvalidOperationException("Legacy COM " + property + " write is unsupported for this Access/Publisher host project; no setter was invoked. " + LegacyHelpMetadataAlternative);
        }

        /// <summary>Annotates matching property descriptions with the truthful unsupported legacy setter status.</summary>
        /// <param name="project">Candidate project checked for restricted native host identity.</param>
        /// <param name="properties">Property catalogue entries whose raw COM read and descriptor state remain unchanged.</param>
        private void DescribeLegacyHelpMetadata(object project, IEnumerable<VbePropertyInfo> properties)
        {
            if (!IsRestrictedLegacyHelpProject(project)) return;
            foreach (var property in properties)
                if (IsLegacyHelpMetadataProperty(property.Name))
                {
                    // ReadOnly describes the COM descriptor, and Value remains the raw COM read.
                    // Neither is replaced with a native General value or a heuristic conversion.
                    property.SetterStatus = LegacyHelpMetadataSetterStatus;
                    property.Display = "Legacy COM write is unsupported. COM readback is not native General acceptance. " + LegacyHelpMetadataAlternative;
                }
        }
    }
}
