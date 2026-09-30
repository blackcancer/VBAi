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
        internal static uint ProcessId => GetCurrentProcessId();
        internal static uint ThreadId => GetCurrentThreadId();
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
        private static object Error(Exception error) => new { Type = error.GetType().FullName, error.HResult };
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern uint GetFileAttributesW(string path);
    }
}
