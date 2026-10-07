using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace VBAi.Tests.Integration
{
    /// <summary>Read-only public token metadata for the exact test process and owned native host.</summary>
    internal static class NativeProcessTokenObservation
    {
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, uint processId);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("advapi32.dll", SetLastError = true)] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool GetTokenInformation(IntPtr token, int informationClass,
            IntPtr information, int size, out int returnedSize);
        [DllImport("advapi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsTokenRestricted(IntPtr token);

        internal static IDictionary<string, object> Read(int processId)
        {
            var result = new Dictionary<string, object>
            {
                ["ProcessId"] = processId,
                ["State"] = "UNVERIFIED",
                ["Scope"] = "Read-only TOKEN_QUERY/process limited information. Public SID/session/integrity/restriction metadata only; no credentials, handles or token mutation."
            };
            IntPtr process = IntPtr.Zero, token = IntPtr.Zero;
            try
            {
                using (var observed = Process.GetProcessById(processId))
                {
                    result["ProcessName"] = observed.ProcessName;
                    result["SessionId"] = observed.SessionId;
                    result["StartedUtc"] = observed.StartTime.ToUniversalTime().ToString("o");
                }
                process = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, (uint)processId);
                if (process == IntPtr.Zero) { result["OpenProcessError"] = Marshal.GetLastWin32Error(); return result; }
                if (!OpenProcessToken(process, 0x8 /* TOKEN_QUERY */, out token))
                { result["OpenTokenError"] = Marshal.GetLastWin32Error(); return result; }
                result["Restricted"] = IsTokenRestricted(token);
                result["UserSid"] = ReadSid(token, 1, result, "UserSid");
                result["IntegritySid"] = ReadSid(token, 25, result, "IntegritySid");
                result["ElevationType"] = ReadInteger(token, 18, result, "ElevationType");
                int? appContainer = ReadInteger(token, 29, result, "IsAppContainer");
                result["IsAppContainer"] = appContainer.HasValue ? (bool?)(appContainer.Value != 0) : null;
                result["TokenSessionId"] = ReadInteger(token, 12, result, "TokenSessionId");
                result["State"] = HasReadFailure(result) ? "PARTIAL" : "READ";
            }
            catch (Exception error) { result["ReadError"] = error.ToString(); }
            finally { if (token != IntPtr.Zero) CloseHandle(token); if (process != IntPtr.Zero) CloseHandle(process); }
            return result;
        }

        private static bool HasReadFailure(IDictionary<string, object> result)
        {
            foreach (string key in result.Keys) if (key.EndsWith("Error", StringComparison.Ordinal)) return true;
            return false;
        }

        private static string ReadSid(IntPtr token, int informationClass, IDictionary<string, object> result, string name)
        {
            return ReadInformation(token, informationClass, result, name, pointer =>
                new SecurityIdentifier(Marshal.ReadIntPtr(pointer)).Value);
        }

        private static int? ReadInteger(IntPtr token, int informationClass, IDictionary<string, object> result, string name)
        {
            string value = ReadInformation(token, informationClass, result, name,
                pointer => Marshal.ReadInt32(pointer).ToString(System.Globalization.CultureInfo.InvariantCulture));
            return value == null ? (int?)null : int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string ReadInformation(IntPtr token, int informationClass, IDictionary<string, object> result,
            string name, Func<IntPtr, string> read)
        {
            int size;
            GetTokenInformation(token, informationClass, IntPtr.Zero, 0, out size);
            int sizingError = Marshal.GetLastWin32Error();
            if (size <= 0 || size > 65536) { result[name + "Error"] = sizingError; return null; }
            IntPtr information = Marshal.AllocHGlobal(size);
            try
            {
                if (!GetTokenInformation(token, informationClass, information, size, out size))
                { result[name + "Error"] = Marshal.GetLastWin32Error(); return null; }
                return read(information);
            }
            finally { Marshal.FreeHGlobal(information); }
        }
    }
}
