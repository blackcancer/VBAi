using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace CodexVBE
{
    internal static class VbeControlCatalog
    {
        private const string ControlCategory = "{40FC6ED4-2438-11CF-A3DB-080036F12502}";

        public static object List(IEnumerable<string> nativeProgIds)
        {
            var result = new List<object>();
            foreach (string progId in nativeProgIds.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                result.Add(new { ProgId = progId, Source = "MSForms native", Hosting = "Built in" });
            foreach (string progId in InstalledControls().OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                if (!nativeProgIds.Contains(progId, StringComparer.OrdinalIgnoreCase))
                    result.Add(new { ProgId = progId, Source = "COM CATID_Control x64",
                        Hosting = "Unverified until Controls.Add succeeds" });
            return result;
        }

        public static bool IsCandidate(string progId, ISet<string> nativeProgIds)
        {
            return nativeProgIds.Contains(progId) ||
                InstalledControls().Contains(progId, StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> InstalledControls()
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (RegistryKey classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (RegistryKey clsids = classes.OpenSubKey("CLSID"))
            {
                if (clsids == null) return result;
                foreach (string clsid in clsids.GetSubKeyNames())
                {
                    using (RegistryKey category = clsids.OpenSubKey(clsid + "\\Implemented Categories\\" + ControlCategory))
                    {
                        if (category == null) continue;
                    }
                    using (RegistryKey progIdKey = clsids.OpenSubKey(clsid + "\\ProgID"))
                    {
                        string progId = progIdKey?.GetValue(null) as string;
                        if (!string.IsNullOrWhiteSpace(progId)) result.Add(progId);
                    }
                }
            }
            return result;
        }
    }
}
