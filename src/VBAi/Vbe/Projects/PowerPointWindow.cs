using System;

namespace VBAi
{

    /// <summary>Uses the published Office vtable for HWND, a restricted member unavailable through IDispatch.</summary>
    internal static class PowerPointWindow
    {

        /// <summary>Reads  for power point window.</summary>
        /// <param name="application">object that supplies the application for this operation.</param>
        /// <returns>int ptr produced by the operation for read on power point window.</returns>
        internal static IntPtr Read(object application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            // No-PIA embedding preserves the official COM layout without a runtime assembly dependency.
            return new IntPtr(((Microsoft.Office.Interop.PowerPoint._Application)application).HWND);
        }
    }
}
