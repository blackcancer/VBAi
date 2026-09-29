using System;

namespace VBAi
{
    /// <summary>Uses the published Office vtable for HWND, a restricted member unavailable through IDispatch.</summary>
    internal static class PowerPointWindow
    {
        internal static IntPtr Read(object application)
        {
            if (application == null) throw new ArgumentNullException(nameof(application));
            // No-PIA embedding preserves the official COM layout without a runtime assembly dependency.
            return new IntPtr(((Microsoft.Office.Interop.PowerPoint._Application)application).HWND);
        }
    }
}
