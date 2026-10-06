using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Maintains the legacy help metadata process name state for vbe project components.</summary>
        internal Func<string> LegacyHelpMetadataProcessName = () => Process.GetCurrentProcess().ProcessName;

        /// <summary>Maintains the legacy help metadata native project state for vbe project components.</summary>
        internal Func<object, bool> LegacyHelpMetadataNativeProject = Marshal.IsComObject;

        /// <summary>Maintains the legacy help metadata setter status state for vbe project components.</summary>
        internal const string LegacyHelpMetadataSetterStatus = "HostLegacyWriteUnsupported";

        /// <summary>Maintains the legacy help metadata alternative state for vbe project components.</summary>
        internal const string LegacyHelpMetadataAlternative = "Use read_project_general followed by an explicitly approved asynchronous set_project_general; no automatic fallback or Save is performed.";

        /// <summary>Determines whether legacy help metadata property for vbe project components.</summary>
        /// <param name="property">Text that supplies the property value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is legacy help metadata property on vbe project components.</returns>
        internal static bool IsLegacyHelpMetadataProperty(string property) =>
            string.Equals(property, "HelpFile", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(property, "HelpContextID", StringComparison.OrdinalIgnoreCase);

        /// <summary>Determines whether legacy help metadata host for vbe project components.</summary>
        /// <param name="processName">Text that supplies the process name value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is legacy help metadata host on vbe project components.</returns>
        internal static bool IsLegacyHelpMetadataHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

        /// <summary>Determines whether restricted legacy help project for vbe project components.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is restricted legacy help project on vbe project components.</returns>
        private bool IsRestrictedLegacyHelpProject(object project)
        {
            if (!IsLegacyHelpMetadataHost(LegacyHelpMetadataProcessName()) || !LegacyHelpMetadataNativeProject(project)) return false;
            // A failed Type read propagates. It cannot admit an unverified native project to a setter.
            return Convert.ToInt32(((dynamic)project).Type, CultureInfo.InvariantCulture) == 100;
        }

        /// <summary>Requires legacy help metadata write supported for vbe project components.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="property">Text that supplies the property value. Use the format required by the calling operation.</param>
        private void RequireLegacyHelpMetadataWriteSupported(object project, string property)
        {
            if (IsLegacyHelpMetadataProperty(property) && IsRestrictedLegacyHelpProject(project))
                throw new InvalidOperationException("Legacy COM " + property + " write is unsupported for this Access/Publisher host project; no setter was invoked. " + LegacyHelpMetadataAlternative);
        }

        /// <summary>Handles describe legacy help metadata for vbe project components.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="properties">i enumerable&lt;vbe property info&gt; that supplies the properties for this operation.</param>
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
