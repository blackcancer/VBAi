using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Records an observed COM-created Excel identity; its handle was acquired by PID, not returned by process creation.</summary>
        internal static void RecordComAttachedIdentity(IDictionary<string, object> receipt, int processId,
            int windowProcessId, long retainedHandle, string processStartedUtc, uint windowThreadId,
            string ownerDesktop, string windowDesktop, string inputDesktop)
        {
            if (receipt == null) throw new ArgumentNullException(nameof(receipt));
            if (processId <= 0 || windowProcessId != processId || retainedHandle <= 0 || windowThreadId == 0 ||
                !DateTime.TryParse(processStartedUtc, out _) || string.IsNullOrWhiteSpace(inputDesktop))
                throw new InvalidOperationException("COM Excel's observed process identity is incomplete or changed.");
            RequireOwnedWindowDesktop(null, ownerDesktop, windowDesktop);
            receipt["ProcessIdentityOrigin"] = "ComActivator_GetProcessById_Handle";
            receipt["RetainedProcessHandle"] = retainedHandle;
            receipt["HostStartedUtc"] = processStartedUtc;
            receipt["ApplicationHwndProcessId"] = windowProcessId;
            receipt["ApplicationWindowThreadId"] = windowThreadId;
            receipt["OwnerThreadDesktop"] = ownerDesktop;
            receipt["ApplicationWindowDesktop"] = windowDesktop;
            receipt["InputDesktopAtObservation"] = inputDesktop;
        }

        /// <summary>Pairs shutdown with the exact retained COM process observation without claiming a creation handle.</summary>
        internal static void CopyComAttachedIdentityToShutdown(IDictionary<string, object> startup,
            IDictionary<string, object> shutdown)
        {
            if (startup == null) throw new ArgumentNullException(nameof(startup));
            if (shutdown == null) throw new ArgumentNullException(nameof(shutdown));
            if (!startup.ContainsKey("ProcessIdentityOrigin")) return;
            var paired = new Dictionary<string, object>();
            foreach (string key in new[] { "ProcessIdentityOrigin", "RetainedProcessHandle", "HostStartedUtc",
                "ApplicationHwndProcessId", "ApplicationWindowThreadId", "OwnerThreadDesktop",
                "ApplicationWindowDesktop", "InputDesktopAtObservation" })
            {
                if (!startup.TryGetValue(key, out object value) || value == null)
                    throw new InvalidOperationException("The COM process identity receipt is incomplete at shutdown: " + key);
                paired[key == "HostStartedUtc" ? "ProcessStartedUtc" : key] = value;
            }
            foreach (var item in paired) shutdown[item.Key] = item.Value;
        }
    }
}
