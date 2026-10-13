using Microsoft.CSharp.RuntimeBinder;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Resolves Excel through a document window belonging to the VBE host process.</summary>
    internal static class ExcelOwnedApplication
    {

        /// <summary>Uses Office NativeOM before the registered application fallback; callers retain their PID guard.</summary>
        /// <param name="processId">Exact owning process.</param>
        /// <param name="registeredApplication">Compatibility fallback when no native document is available.</param>
        /// <returns>Application for an owned Excel document, or the fallback result.</returns>
        internal static object Resolve(int processId, Func<object> registeredApplication)
            => Resolve(processId, registeredApplication, null);

        /// <summary>Resolves an owned application while balancing every temporary native interface acquisition.</summary>
        /// <param name="processId">Exact owning process.</param>
        /// <param name="registeredApplication">Compatibility fallback when no native document is available.</param>
        /// <param name="releaseAcquired">Optional release boundary for each acquired native interface; borrowed arguments are never released.</param>
        /// <returns>Matched application acquisition transferred to the caller, or the fallback result.</returns>
        /// <remarks>Only the matched application acquisition transfers to the caller. Document-window
        /// acquisitions and rejected applications are released before returning or using the fallback.</remarks>
        internal static object Resolve(int processId, Func<object> registeredApplication, Action<object> releaseAcquired)
        {
            var documents = new List<IntPtr>();
            VbeDebugWindows.EnumWindows((window, parameter) =>
            {
                VbeDebugWindows.GetWindowThreadProcessId(window, out uint owner);
                if (owner != (uint)processId) return true;
                VbeDebugWindows.EnumChildWindows(window, (child, childParameter) =>
                {
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
                using (var references = new VbaTestComReferences(releaseAcquired))
                {
                    try
                    {
                        Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                        int result = VbeDebugWindows.AccessibleObjectFromWindow(document, 0xFFFFFFF0, ref dispatch, out object window);
                        references.Add(window);
                        if (result != 0 || window == null) continue;
                        dynamic application = references.Add((object)((dynamic)window).Application);
                        VbeDebugWindows.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(application.Hwnd)), out uint owner);
                        if (owner == (uint)processId)
                        {
                            references.Transfer((object)application);
                            return application;
                        }
                    }
                    catch (COMException) { }
                    catch (RuntimeBinderException) { }
                }
            }
            return registeredApplication();
        }
    }
}
