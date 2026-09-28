using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;

namespace CodexVBE
{
    /// <summary>Resolves Excel through a document window belonging to the VBE host process.</summary>
    internal static class ExcelOwnedApplication
    {
        /// <summary>Uses Office NativeOM before the registered application fallback; callers retain their PID guard.</summary>
        /// <param name="processId">Exact owning process.</param>
        /// <param name="registeredApplication">Compatibility fallback when no native document is available.</param>
        /// <returns>Application for an owned Excel document, or the fallback result.</returns>
        internal static object Resolve(int processId, Func<object> registeredApplication)
        {
            var documents = new List<IntPtr>();
            VbeDebugWindows.EnumWindows((window, parameter) => {
                uint owner;
                VbeDebugWindows.GetWindowThreadProcessId(window, out owner);
                if (owner != (uint)processId) return true;
                VbeDebugWindows.EnumChildWindows(window, (child, childParameter) => {
                    VbeDebugWindows.GetWindowThreadProcessId(child, out owner);
                    if (owner != (uint)processId) return true;
                    var name = new StringBuilder(256);
                    VbeDebugWindows.GetClassName(child, name, name.Capacity);
                    if (name.ToString() == "EXCEL7") documents.Add(child);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            foreach (IntPtr document in documents)
            {
                try
                {
                    Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                    object window;
                    int result = VbeDebugWindows.AccessibleObjectFromWindow(document, 0xFFFFFFF0, ref dispatch, out window);
                    if (result != 0 || window == null) continue;
                    dynamic application = ((dynamic)window).Application;
                    uint owner;
                    VbeDebugWindows.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(application.Hwnd)), out owner);
                    if (owner == (uint)processId) return application;
                }
                catch (COMException) { }
                catch (RuntimeBinderException) { }
            }
            return registeredApplication();
        }
    }
}
