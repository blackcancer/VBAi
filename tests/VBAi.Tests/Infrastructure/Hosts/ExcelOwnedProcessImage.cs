using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Reads one retained process handle without enumerating or caching process modules.</summary>
    internal static class ExcelOwnedProcessImage
    {
        internal const int MaximumPathCharacters = 32768;
        internal delegate bool ImageQuery(IntPtr handle, int flags, StringBuilder path, ref int size);

        [DllImport("kernel32.dll", EntryPoint = "QueryFullProcessImageNameW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool QueryImage(IntPtr handle, int flags, StringBuilder path, ref int size);

        internal static string Read(IntPtr retainedHandle)
        {
            return Read(retainedHandle, QueryImage, Marshal.GetLastWin32Error);
        }

        /// <summary>One bounded native read; an unavailable identity is a failure, never authorization to relaunch.</summary>
        internal static string Read(IntPtr retainedHandle, ImageQuery query, Func<int> lastError)
        {
            if (retainedHandle == IntPtr.Zero || retainedHandle == new IntPtr(-1))
                throw new ArgumentException("A retained, valid process handle is required.", nameof(retainedHandle));
            if (query == null) throw new ArgumentNullException(nameof(query));
            if (lastError == null) throw new ArgumentNullException(nameof(lastError));
            var path = new StringBuilder(MaximumPathCharacters);
            int size = path.Capacity;
            if (!query(retainedHandle, 0, path, ref size))
                throw new Win32Exception(lastError(), "QueryFullProcessImageNameW could not verify the retained process executable; preserve the process without fallback.");
            if (size <= 0 || size >= MaximumPathCharacters || size != path.Length)
                throw new InvalidOperationException("The retained process image path is empty, truncated, or inconsistent; identity remains unverified.");
            return ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(path.ToString());
        }
    }
}
