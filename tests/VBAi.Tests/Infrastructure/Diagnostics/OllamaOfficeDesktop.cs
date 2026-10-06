using System;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Explicit Q028 real-desktop containment without changing private-desktop contracts.</summary>
    internal static class OllamaOfficeDesktop
    {
        internal const string MainEnvironment = "VBAi_RUN_OLLAMA_OFFICE_MAIN_DESKTOP_TESTS";
        internal static bool MainEnabled => Environment.GetEnvironmentVariable(MainEnvironment) == "1";
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flag);

        internal static void RequireMainOptIn(string scenario, string required, string configured)
        {
            if (scenario != "1" || !string.IsNullOrEmpty(required) || !string.IsNullOrEmpty(configured))
                throw new InvalidOperationException("Q028 Main requires its scenario opt-in and both private-desktop variables empty.");
        }

        internal static void Require(string privateDesktop)
        {
            if (!MainEnabled) { IsolatedTestDesktop.RequireCurrent(privateDesktop); return; }
            RequireMainOptIn(Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_OFFICE_TESTS"),
                Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP"),
                Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"));
            if (!string.IsNullOrEmpty(privateDesktop) && privateDesktop != "Default")
                throw new InvalidOperationException("Main Q028 cannot adopt a private-desktop descriptor.");
            IsolatedTestDesktop.RequireMainCurrent();
        }

        internal static void RequireWindow(string desktop, uint pid, bool requireVisible, IntPtr window)
        {
            Require(desktop);
            if (!MainEnabled) { IsolatedTestDesktop.RequireOfficeWindowInventory(desktop, pid, requireVisible, window); return; }
            if (pid == 0 || window == IntPtr.Zero) throw new InvalidOperationException("An exact owned main-desktop window is required.");
            // The complete native Default inventory binds the root. Action-specific
            // HWND/PID/thread/parent and visibility checks remain with each caller.
            var root = GetAncestor(window, 2);
            var inventory = IsolatedTestDesktop.ReadMainWindows(pid);
            IsolatedTestDesktop.RequireMainInventory(inventory, pid, root, false, false);
            if (requireVisible && !Array.Exists(inventory.Windows, row => row.Handle == root.ToInt64() && row.Visible))
                throw new InvalidOperationException("The original owned main root is not visible.");
        }
    }
}
