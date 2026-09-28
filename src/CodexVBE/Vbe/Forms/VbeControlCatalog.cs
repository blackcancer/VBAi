using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace CodexVBE
{
    /// <summary>Recense les contrôles MSForms natifs et les contrôles COM enregistrés en 64 bits.</summary>
    internal static class VbeControlCatalog
    {
        /// <summary>CATID_Control recherchée dans les catégories COM implémentées.</summary>
        private const string ControlCategory = "{40FC6ED4-2438-11CF-A3DB-080036F12502}";
        /// <summary>Ouvre le catalogue COM 64 bits natif dont la lecture reste sans mutation.</summary>
        internal static Func<RegistryKey> OpenClasses = () => RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64);

        /// <summary>Retourne les contrôles natifs puis les candidats installés, sans dupliquer les ProgID natifs.</summary>
        /// <param name="nativeProgIds">ProgID des contrôles MSForms connus comme natifs.</param>
        /// <returns>Entrées descriptives qui distinguent les contrôles natifs des candidats non vérifiés.</returns>
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

        /// <summary>Indique si le ProgID figure parmi les contrôles natifs ou les contrôles enregistrés.</summary>
        /// <param name="progId">ProgID à vérifier.</param>
        /// <param name="nativeProgIds">ProgID natifs connus.</param>
        /// <returns>Vrai si le ProgID est natif ou inscrit comme contrôle COM en 64 bits.</returns>
        public static bool IsCandidate(string progId, ISet<string> nativeProgIds)
        {
            return nativeProgIds.Contains(progId) ||
                InstalledControls().Contains(progId, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>Énumère les ProgID des classes COM 64 bits qui annoncent la catégorie de contrôle.</summary>
        /// <returns>Ensemble sans doublons des ProgID installés comme contrôles en 64 bits.</returns>
        private static IEnumerable<string> InstalledControls()
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (RegistryKey classes = OpenClasses())
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
