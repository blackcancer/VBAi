using System;
using System.Runtime.InteropServices;
namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Sends native UI-state messages only to controls created by the test process.</summary>
    internal static class NativeUiState
    {
        /// <summary>Updates focus cue visibility without generating keyboard input.</summary>
        [DllImport("user32.dll")]
        internal static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    }
}
