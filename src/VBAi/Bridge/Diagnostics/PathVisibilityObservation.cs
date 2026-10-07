using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi
{

    /// <summary>Observes synthetic path attributes and effective-token metadata without reading contents or changing state.</summary>
    internal static class PathVisibilityObservation
    {

        /// <summary>Gets the current process ID used to bind the diagnostic result to its host.</summary>
        /// <value>Win32 process ID of the calling process.</value>
        internal static uint ProcessId => GetCurrentProcessId();

        /// <summary>Gets the current native thread ID used to verify owner-STA execution.</summary>
        /// <value>Win32 thread ID of the calling thread.</value>
        internal static uint ThreadId => GetCurrentThreadId();

        /// <summary>Observes managed and Win32 attributes for exactly four synthetic paths and captures effective-token metadata before and after.</summary>
        /// <param name="syntheticAllowlist">Two owned directories and their two fixed synthetic-file paths; any other count is rejected.</param>
        /// <returns>Process/thread, apartment, token, and per-path attribute observations; no path contents are read.</returns>
        internal static IDictionary<string, object> Read(string[] syntheticAllowlist)
        {
            if (syntheticAllowlist == null || syntheticAllowlist.Length != 4)
                throw new ArgumentException("An exact two-directory/two-file synthetic allowlist is required.");
            var paths = new List<object>();
            var report = new Dictionary<string, object>
            {
                ["ProcessId"] = GetCurrentProcessId(),
                ["NativeThreadId"] = GetCurrentThreadId(),
                ["ManagedThreadId"] = Thread.CurrentThread.ManagedThreadId,
                ["Apartment"] = Thread.CurrentThread.GetApartmentState().ToString(),
                ["Utc"] = DateTime.UtcNow.ToString("o"),
                ["EffectiveTokenBefore"] = EffectiveTokenObservation.Read(),
                ["Scope"] = "Calling-thread synthetic path attributes and effective token; caller must independently establish exact owner-STA identity."
            };
            foreach (string path in syntheticAllowlist)
            {
                var item = new Dictionary<string, object> { ["Path"] = path };
                try { item["DirectoryExists"] = Directory.Exists(path); }
                catch (Exception error) { item["DirectoryExistsError"] = Error(error); }
                try { item["ManagedAttributes"] = (uint)File.GetAttributes(path); }
                catch (Exception error) { item["ManagedAttributesError"] = Error(error); }
                uint attributes = GetFileAttributesW(path);
                int errorCode = Marshal.GetLastWin32Error(); // Immediate capture, before another native call.
                item["NativeAttributes"] = attributes;
                item["NativeSucceeded"] = attributes != uint.MaxValue;
                item["NativeLastError"] = errorCode;
                item["NativeErrorMeaningful"] = attributes == uint.MaxValue;
                paths.Add(item);
            }
            report["Paths"] = paths;
            report["EffectiveTokenAfter"] = EffectiveTokenObservation.Read();
            report["FinalProcessId"] = GetCurrentProcessId(); report["FinalNativeThreadId"] = GetCurrentThreadId();
            return report;
        }

        /// <summary>Converts a managed filesystem exception to non-sensitive type and HRESULT fields.</summary>
        /// <param name="error">Exception raised while reading one path attribute.</param>
        /// <returns>Object containing the exception type name and HRESULT, without the path contents.</returns>
        private static object Error(Exception error) => new { Type = error.GetType().FullName, error.HResult };

        /// <summary>Gets the calling process ID for the diagnostic observation.</summary>
        /// <returns>Current Win32 process ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

        /// <summary>Gets the calling native thread ID for the owner-STA observation.</summary>
        /// <returns>Current Win32 thread ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Reads Win32 file attributes without opening or reading the target contents.</summary>
        /// <param name="path">Synthetic local path from the validated allowlist.</param>
        /// <returns>Attribute bitmask, or INVALID_FILE_ATTRIBUTES on failure; capture GetLastError immediately in that case.</returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern uint GetFileAttributesW(string path);
    }
}
