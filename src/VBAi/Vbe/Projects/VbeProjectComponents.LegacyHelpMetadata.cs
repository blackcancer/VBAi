using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        internal Func<string> LegacyHelpMetadataProcessName = () => Process.GetCurrentProcess().ProcessName;
        internal Func<object, bool> LegacyHelpMetadataNativeProject = Marshal.IsComObject;
        internal const string LegacyHelpMetadataSetterStatus = "HostLegacyWriteUnsupported";
        internal const string LegacyHelpMetadataAlternative = "Use read_project_general followed by an explicitly approved asynchronous set_project_general; no automatic fallback or Save is performed.";

        internal static bool IsLegacyHelpMetadataProperty(string property) =>
            string.Equals(property, "HelpFile", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(property, "HelpContextID", StringComparison.OrdinalIgnoreCase);

        internal static bool IsLegacyHelpMetadataHost(string processName) =>
            string.Equals(processName, "MSACCESS", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(processName, "MSPUB", StringComparison.OrdinalIgnoreCase);

        private bool IsRestrictedLegacyHelpProject(object project)
        {
            if (!IsLegacyHelpMetadataHost(LegacyHelpMetadataProcessName()) || !LegacyHelpMetadataNativeProject(project)) return false;
            // A failed Type read propagates. It cannot admit an unverified native project to a setter.
            return Convert.ToInt32(((dynamic)project).Type, CultureInfo.InvariantCulture) == 100;
        }

        private void RequireLegacyHelpMetadataWriteSupported(object project, string property)
        {
            if (IsLegacyHelpMetadataProperty(property) && IsRestrictedLegacyHelpProject(project))
                throw new InvalidOperationException("Legacy COM " + property + " write is unsupported for this Access/Publisher host project; no setter was invoked. " + LegacyHelpMetadataAlternative);
        }

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
