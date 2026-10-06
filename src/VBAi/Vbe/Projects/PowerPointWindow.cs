using System;

namespace VBAi
{

    /// <summary>Uses the published Office vtable for HWND, a restricted member unavailable through IDispatch.</summary>
    internal static class PowerPointWindow
    {

        /// <summary>Reads PowerPoint's native main-window handle through its published COM vtable.</summary>
        /// <param name="application">PowerPoint Application COM object.</param>
        /// <returns>PowerPoint HWND as an IntPtr.</returns>
        /// <exception cref="ArgumentNullException">The application object is null.</exception>
        internal static IntPtr Read(object application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            // No-PIA embedding preserves the official COM layout without a runtime assembly dependency.
            return new IntPtr(((Microsoft.Office.Interop.PowerPoint._Application)application).HWND);
        }
    }
}
