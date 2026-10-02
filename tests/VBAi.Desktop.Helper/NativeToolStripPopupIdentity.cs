using System;

namespace VBAi.Desktop.Helper
{
    /// <summary>Recognizes the observed UIA roles of an owned WinForms dropdown, without authorizing its actions.</summary>
    public static class NativeToolStripPopupIdentity
    {
        /// <summary>Requires a WinForms popup class; the legacy ToolBar role additionally requires the dropdown class.</summary>
        public static bool Matches(string role, string nativeClass)
        {
            if (nativeClass == null || !nativeClass.StartsWith("WindowsForms10.Window.", StringComparison.Ordinal))
                return false;
            return role == "ControlType.Menu" || role == "ControlType.ToolBar" &&
                nativeClass.StartsWith("WindowsForms10.Window.20808.app.", StringComparison.Ordinal);
        }
    }
}
