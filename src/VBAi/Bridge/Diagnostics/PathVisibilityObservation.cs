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

        /// <summary>Gets the process id.</summary>
        /// <value>Current process id exposed by path visibility observation.</value>
        internal static uint ProcessId => GetCurrentProcessId();

        /// <summary>Gets the thread id.</summary>
        /// <value>Current thread id exposed by path visibility observation.</value>
        internal static uint ThreadId => GetCurrentThreadId();

        /// <summary>Reads  for path visibility observation.</summary>
        /// <param name="syntheticAllowlist">string[] that supplies the synthetic allowlist for this operation.</param>
        /// <returns>i dictionary&lt;string, object&gt; produced by the operation for read on path visibility observation.</returns>
        internal static IDictionary<string, object> Read(string[] syntheticAllowlist)
        {
            if (syntheticAllowlist == null || syntheticAllowlist.Length != 4)
                throw new ArgumentException("An exact two-directory/two-file synthetic allowlist is required.");
            var paths = new List<object>();
            var report = new Dictionary<string, object> {
                ["ProcessId"] = GetCurrentProcessId(), ["NativeThreadId"] = GetCurrentThreadId(),
                ["ManagedThreadId"] = Thread.CurrentThread.ManagedThreadId,
                ["Apartment"] = Thread.CurrentThread.GetApartmentState().ToString(),
                ["Utc"] = DateTime.UtcNow.ToString("o"), ["EffectiveTokenBefore"] = EffectiveTokenObservation.Read(),
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

        /// <summary>Handles error for path visibility observation.</summary>
        /// <param name="error">Exception describing the error failure.</param>
        /// <returns>object produced by the operation for error on path visibility observation.</returns>
        private static object Error(Exception error) => new { Type = error.GetType().FullName, error.HResult };

        /// <summary>Returns current process id for path visibility observation.</summary>
        /// <returns>uint produced by the operation for get current process id on path visibility observation.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

        /// <summary>Returns current thread id for path visibility observation.</summary>
        /// <returns>uint produced by the operation for get current thread id on path visibility observation.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Returns file attributes w for path visibility observation.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>uint produced by the operation for get file attributes w on path visibility observation.</returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern uint GetFileAttributesW(string path);
    }
}
